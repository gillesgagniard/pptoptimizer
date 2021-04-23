package main

import (
	"flag"
	"os"
	"path/filepath"
	"strings"

	log "github.com/sirupsen/logrus"
)

func main() {
	flagVerbose := flag.Bool("v", false, "verbose logging")
	flagInputFile := flag.String("f", "", "pptx input file")
	flagConvertBitmaps := flag.Bool("convert", true, "convert uncompressed pictures such as TIFF to PNG (lossless)")
	flagCleanLayouts := flag.Bool("layouts", false, "remove all unused layouts, masters, and their media files")
	flagAllOptimizations := flag.Bool("a", false, "apply all optimizations")
	flag.Parse()

	if *flagVerbose {
		log.SetLevel(log.DebugLevel)
	}

	oldinfo, err := os.Stat(*flagInputFile)
	if err != nil {
		log.Fatalln("cannot open input file:", err)
	}

	p := NewPowerpointDoc()
	defer p.Close()
	p.ParseFile(*flagInputFile)

	if *flagConvertBitmaps || *flagAllOptimizations {
		p.ConvertPictures()
	}
	if *flagCleanLayouts || *flagAllOptimizations {
		p.RemoveUnusedLayouts()
		p.RemoveUnusedMasters()
		p.RemoveUnusedMedias()
	}

	outputFileName := strings.Replace(*flagInputFile, filepath.Ext(*flagInputFile), ".new.pptx", 1)
	p.SaveFile(outputFileName)

	newinfo, err := os.Stat(outputFileName)
	if err != nil {
		log.Fatal(err)
	}

	log.Infoln("size", *flagInputFile, oldinfo.Size(), outputFileName, newinfo.Size())
}
