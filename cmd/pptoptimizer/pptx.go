package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"image/png"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	log "github.com/sirupsen/logrus"

	"github.com/beevik/etree"
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

type Types struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`
	Default []struct {
		Extension   string `xml:"Extension,attr"`
		ContentType string `xml:"ContentType,attr"`
	}
	Override []struct {
		PartName    string `xml:"PartName,attr"`
		ContentType string `xml:"ContentType,attr"`
	}
}

type PowerpointDoc struct {
	sourceFileReader *zip.ReadCloser
	medias           map[string]Media
	slideRels        []Relationships
	slideLayoutRels  []Relationships
	slideMasterRels  []Relationships
	slideMasters     []*etree.Document
	contentTypes     Types
}

var reSlideNumber = regexp.MustCompile(`/[a-zA-Z]+([0-9]+)\.xml`)
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

func updateSlideMasters(sms []*etree.Document, pos int, sm *etree.Document) []*etree.Document {
	// increase length if needed
	newsms := sms
	if pos > len(sms) {
		newsms = make([]*etree.Document, pos)
		copy(newsms, sms)
	}
	newsms[pos-1] = sm
	return newsms
}

func getSlideNumberFromFilename(fname string) (int, error) {
	matches := reSlideNumber.FindStringSubmatch(fname)
	if len(matches) != 2 {
		return 0, errors.New("invalid file name " + fname)
	}
	slideNumber, err := strconv.Atoi(matches[1]) // matches[0] is the whole match
	if err != nil {
		return 0, errors.New("invalid slide number " + fname)
	}
	return slideNumber, nil
}

func parseRelationships(rels []Relationships, reltype string, f *zip.File) []Relationships {
	if strings.HasPrefix(f.Name, fmt.Sprintf("ppt/%ss/_rels/", reltype)) {
		relf, err := f.Open()
		if err != nil {
			log.Fatal(err)
		}
		defer relf.Close()
		relfxml, err := ioutil.ReadAll(relf)
		if err != nil {
			log.Fatal(err)
		}
		rel := Relationships{}
		err = xml.Unmarshal(relfxml, &rel)
		if err != nil {
			log.Fatal(err)
		}
		slideNumber, _ := getSlideNumberFromFilename(f.Name)
		rels = updateRelationships(rels, slideNumber, rel)
	}
	return rels
}

func saveRelationships(rels []Relationships, reltype string, outz *zip.Writer) {
	for i, r := range rels {
		if len(r.Relationship) == 0 { // skip empty
			continue
		}
		fo, err := outz.Create(fmt.Sprintf("ppt/%ss/_rels/%s%d.xml.rels", reltype, reltype, i+1))
		if err != nil {
			log.Fatal(err)
		}
		xmlout, _ := xml.Marshal(r)
		log.Debugln("new", reltype, "rels", i+1)
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

	// parse archive contents
	for _, f := range p.sourceFileReader.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			p.medias[f.Name] = Media{size: f.UncompressedSize64}
		} else if f.Name == "[Content_Types].xml" {
			ctf, err := f.Open()
			if err != nil {
				log.Fatal(err)
			}
			defer ctf.Close()
			ctxml, err := ioutil.ReadAll(ctf)
			if err != nil {
				log.Fatal(err)
			}
			err = xml.Unmarshal(ctxml, &p.contentTypes)
			if err != nil {
				log.Fatal(err)
			}
		} else if strings.HasPrefix(f.Name, "ppt/slideMasters/slideMaster") {
			smf, err := f.Open()
			if err != nil {
				log.Fatal(err)
			}
			defer smf.Close()
			doc := etree.NewDocument()
			if _, err := doc.ReadFrom(smf); err != nil {
				log.Fatal(err)
			}
			masterNumber, _ := getSlideNumberFromFilename(f.Name)
			p.slideMasters = updateSlideMasters(p.slideMasters, masterNumber, doc)
		} else {
			p.slideRels = parseRelationships(p.slideRels, "slide", f)
			p.slideLayoutRels = parseRelationships(p.slideLayoutRels, "slideLayout", f)
			p.slideMasterRels = parseRelationships(p.slideMasterRels, "slideMaster", f)
		}
	}

	return nil
}

func (p *PowerpointDoc) SaveFile(f string) error {
	log.Debugln("save pptx", f)
	outf, err := os.Create(f)
	if err != nil {
		log.Fatal(err)
	}
	defer outf.Close()
	outz := zip.NewWriter(outf)
	defer outz.Close()

	for _, f := range p.sourceFileReader.File {
		if f.Name == "[Content_Types].xml" || strings.HasPrefix(f.Name, "ppt/slides/_rels/") || strings.HasPrefix(f.Name, "ppt/slideLayouts/_rels/") || strings.HasPrefix(f.Name, "ppt/slideMasters/") {
			log.Debugln("skip", f.Name)
		} else if _, ok := p.medias[f.Name]; strings.HasPrefix(f.Name, "ppt/media/") && !ok {
			log.Debugln("media", f.Name, "has been removed, skip it")
		} else {
			log.Debugln("copy file", f.Name)
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
			log.Debugln("add new media file", k, m.size)
			fo, err := outz.Create(k)
			if err != nil {
				log.Fatal(err)
			}
			fo.Write(m.data)
		}
	}

	// rewrite all slide rels
	saveRelationships(p.slideRels, "slide", outz)
	saveRelationships(p.slideLayoutRels, "slideLayout", outz)
	saveRelationships(p.slideMasterRels, "slideMaster", outz)

	// rewrite content types
	fo, err := outz.Create("[Content_Types].xml")
	if err != nil {
		log.Fatal(err)
	}
	xmlout, _ := xml.Marshal(p.contentTypes)
	fo.Write([]byte(xmlHeader))
	fo.Write(xmlout)

	// rewrite slide masters
	for i, sm := range p.slideMasters {
		fo, err := outz.Create(fmt.Sprintf("ppt/slideMasters/slideMaster%d.xml", i+1))
		if err != nil {
			log.Fatal(err)
		}
		sm.WriteTo(fo)
	}

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
		log.Debugln("slide", i+1, "total media size", slideSize)
	}
}

func (p *PowerpointDoc) ConvertPictures() {
	for _, f := range p.sourceFileReader.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			log.Debugln("media file", f.Name, f.UncompressedSize64)
			if strings.ToLower(filepath.Ext(f.Name)) == ".tiff" {
				log.Infoln("converting media", f.Name, f.UncompressedSize64, "to png ...")
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
				log.Infoln("converted media", newfilename, p.medias[newfilename].size)
			}
		}
	}
}

func (p *PowerpointDoc) FindUsedLayouts() []bool {
	usedSlideLayouts := make([]bool, len(p.slideLayoutRels))
	for _, rels := range p.slideRels {
		for _, rel := range rels.Relationship {
			if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" {
				layoutNumber, _ := getSlideNumberFromFilename(rel.Target)
				usedSlideLayouts[layoutNumber-1] = true
			}
		}
	}
	return usedSlideLayouts
}

func removeLayoutFromMaster(master *etree.Document, id string) {
	for _, e := range master.FindElements(fmt.Sprintf("//p:sldLayoutId[@r:id='%s']", id)) {
		log.Debugln("found layout id", id, "in master -> remove")
		if result := e.Parent().RemoveChild(e); result == nil {
			log.Fatal("cannot remove child !!", *e.Parent(), *e)
		}
	}
}

func (p *PowerpointDoc) RemoveUnusedLayouts() {
	usedSlideLayouts := p.FindUsedLayouts()
	for i, b := range usedSlideLayouts {
		if !b { // unused -> remove
			log.Infoln("remove unused slide layout", i+1)

			// remove from content types
			for j, o := range p.contentTypes.Override {
				if o.PartName == fmt.Sprintf("/ppt/slideLayouts/slideLayout%d.xml", i+1) {
					copy(p.contentTypes.Override[j:], p.contentTypes.Override[j+1:])
					p.contentTypes.Override = p.contentTypes.Override[:len(p.contentTypes.Override)-1]
					break
				}
			}

			// remove from slide master
			for j, relsm := range p.slideMasterRels {
				for k, relm := range relsm.Relationship {
					if relm.Target == fmt.Sprintf("../slideLayouts/slideLayout%d.xml", i+1) {
						layoutid := relm.Id
						// remove layout reference in slide master itself
						removeLayoutFromMaster(p.slideMasters[j], layoutid)
						copy(p.slideMasterRels[j].Relationship[k:], p.slideMasterRels[j].Relationship[k+1:])
						p.slideMasterRels[j].Relationship = p.slideMasterRels[j].Relationship[:len(p.slideMasterRels[j].Relationship)-1]
						break
					}
				}
			}

			// remove slide layout itself
			p.slideLayoutRels[i] = Relationships{}
		}
	}
}

func (p *PowerpointDoc) FindUsedMedias() map[string]bool {
	usedMedias := make(map[string]bool)
	allrels := append(p.slideRels, p.slideLayoutRels...)
	allrels = append(allrels, p.slideMasterRels...)
	for _, rels := range allrels {
		for _, rel := range rels.Relationship {
			if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" {
				usedMedias["ppt/media/"+filepath.Base(rel.Target)] = true
			}
		}
	}
	return usedMedias
}

func (p *PowerpointDoc) RemoveUnusedMedias() {
	usedMedias := p.FindUsedMedias()
	for k := range p.medias {
		if _, ok := usedMedias[k]; !ok {
			log.Infoln("remove unused media", k)
			delete(p.medias, k)
		}
	}
}
