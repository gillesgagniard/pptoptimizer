package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"image/png"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"golang.org/x/image/tiff"
)

type Relationship struct {
	Id         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

type Relationships struct {
	XMLName      xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationship []Relationship
}

func (r *Relationships) ReplaceTarget(oldbasename string, newbasename string) {
	for i, rel := range r.Relationship {
		if strings.HasSuffix(rel.Target, oldbasename) {
			r.Relationship[i].Target = strings.Replace(rel.Target, oldbasename, newbasename, 1)
		}
	}
}

type Media struct {
	size uint64
	data []byte
}

type PowerpointDoc struct {
	sourceFileReader *zip.ReadCloser
	medias           map[string]Media
	slideRels        []Relationships
	slideLayoutRels  []Relationships
	slideMasterRels  []Relationships
}

var reSlideNumber = regexp.MustCompile("/[a-zA-Z]+([0-9]+).xml.rels")
var xmlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"

func NewPowerpointDoc() *PowerpointDoc {
	pptx := PowerpointDoc{}
	pptx.medias = make(map[string]Media)
	return &pptx
}

func (p *PowerpointDoc) Close() {
	if p.sourceFileReader != nil {
		p.sourceFileReader.Close()
	}
}

func updateRelationships(rels []Relationships, pos int, r Relationships) []Relationships {
	// increase length if needed
	newrels := rels
	if pos > len(rels) {
		newrels = make([]Relationships, pos)
		copy(newrels, rels)
	}
	newrels[pos-1] = r
	return newrels
}

func parseRelationships(rels []Relationships, reltype string, f *zip.File) []Relationships {
	if strings.HasPrefix(f.Name, fmt.Sprintf("ppt/%ss/_rels/", reltype)) {
		matches := reSlideNumber.FindStringSubmatch(f.Name)
		slideNumber, _ := strconv.Atoi(matches[1]) // matches[0] is the whole match
		relf, err := f.Open()
		if err != nil {
			log.Fatal(err)
		}
		defer relf.Close()
		if err != nil {
			log.Fatal(err)
		}
		relfxml, err := ioutil.ReadAll(relf)
		if err != nil {
			log.Fatal(err)
		}
		rel := Relationships{}
		err = xml.Unmarshal(relfxml, &rel)
		if err != nil {
			log.Fatal(err)
		}
		rels = updateRelationships(rels, slideNumber, rel)
	}
	return rels
}

func saveRelationships(rels []Relationships, reltype string, outz *zip.Writer) {
	for i, r := range rels {
		fo, err := outz.Create(fmt.Sprintf("ppt/%ss/_rels/%s%d.xml.rels", reltype, reltype, i+1))
		if err != nil {
			log.Fatal(err)
		}
		xmlout, _ := xml.Marshal(r)
		//log.Println("new", reltype, "rels", i+1, string(xmlout))
		log.Println("new", reltype, "rels", i+1)
		fo.Write([]byte(xmlHeader))
		fo.Write(xmlout)
	}
}

func (p *PowerpointDoc) ParseFile(f string) error {
	r, err := zip.OpenReader(f)
	if err != nil {
		log.Fatalln("pptx is an invalid zip file", err)
		return err
	}
	p.sourceFileReader = r

	// parse media files
	for _, f := range p.sourceFileReader.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			p.medias[f.Name] = Media{size: f.UncompressedSize64}
		}
	}

	// parse slide relationships
	for _, f := range p.sourceFileReader.File {
		p.slideRels = parseRelationships(p.slideRels, "slide", f)
		p.slideLayoutRels = parseRelationships(p.slideLayoutRels, "slideLayout", f)
		p.slideMasterRels = parseRelationships(p.slideMasterRels, "slideMaster", f)
	}

	return nil
}

func (p *PowerpointDoc) SaveFile(f string) error {
	log.Println("save pptx", f)
	outf, err := os.Create(f)
	if err != nil {
		log.Fatal(err)
	}
	defer outf.Close()
	outz := zip.NewWriter(outf)
	defer outz.Close()

	for _, f := range p.sourceFileReader.File {
		if strings.HasPrefix(f.Name, "ppt/slides/_rels/") || strings.HasPrefix(f.Name, "ppt/slideLayouts/_rels/") || strings.HasPrefix(f.Name, "ppt/slideMasters/_rels/") {
			log.Println("skip", f.Name)
		} else if _, ok := p.medias[f.Name]; strings.HasPrefix(f.Name, "ppt/media/") && !ok {
			log.Println("media", f.Name, "has been removed, skip it")
		} else {
			log.Println("copy file", f.Name)
			fi, err := f.Open()
			if err != nil {
				log.Fatal(err)
			}
			fo, err := outz.Create(f.Name)
			if err != nil {
				log.Fatal(err)
			}
			idata, _ := ioutil.ReadAll(fi)
			fo.Write(idata)
			fi.Close()
		}
	}

	// add new media files
	for k, m := range p.medias {
		if m.data != nil {
			log.Println("add new media file", k, m.size)
			fo, err := outz.Create(k)
			if err != nil {
				log.Fatal(err)
			}
			fo.Write(m.data)
		}
	}

	// now rewrite all slide rels
	saveRelationships(p.slideRels, "slide", outz)
	saveRelationships(p.slideLayoutRels, "slideLayout", outz)
	saveRelationships(p.slideMasterRels, "slideMaster", outz)

	return nil
}

func (p *PowerpointDoc) GetSlideMediaSize() {
	for i, r := range p.slideRels {
		slideSize := uint64(0)
		for _, r2 := range r.Relationship {
			if r2.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
				slideSize += p.medias["ppt/media/"+filepath.Base(r2.Target)].size
			}
		}
		log.Println("slide", i+1, "total media size", slideSize)
	}
}

func (p *PowerpointDoc) ConvertPictures() {
	for _, f := range p.sourceFileReader.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			log.Println("media file", f.Name, f.UncompressedSize64)
			if strings.ToLower(filepath.Ext(f.Name)) == ".tiff" {
				log.Println("converting media", f.Name, f.UncompressedSize64, "to png ...")
				tiffFile, err := f.Open()
				if err != nil {
					log.Fatal(err)
				}
				defer tiffFile.Close()
				tiffimg, err := tiff.Decode(tiffFile)
				if err != nil {
					log.Fatal(err)
				}
				pngout := bytes.NewBuffer(nil)
				err = png.Encode(pngout, tiffimg)
				if err != nil {
					log.Fatal(err)
				}
				newfilename := strings.Replace(f.Name, ".tiff", ".png", 1)
				p.medias[newfilename] = Media{size: uint64(pngout.Len()), data: pngout.Bytes()}
				delete(p.medias, f.Name)
				for i := range p.slideRels {
					p.slideRels[i].ReplaceTarget(filepath.Base(f.Name), strings.Replace(filepath.Base(f.Name), ".tiff", ".png", 1))
				}
				for i := range p.slideLayoutRels {
					p.slideLayoutRels[i].ReplaceTarget(filepath.Base(f.Name), strings.Replace(filepath.Base(f.Name), ".tiff", ".png", 1))
				}
				for i := range p.slideMasterRels {
					p.slideMasterRels[i].ReplaceTarget(filepath.Base(f.Name), strings.Replace(filepath.Base(f.Name), ".tiff", ".png", 1))
				}
				log.Println("converted media", newfilename, p.medias[newfilename].size)
			}
		}
	}
}
