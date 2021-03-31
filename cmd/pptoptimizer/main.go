package main

import (
	"flag"
	"log"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	flagInputFile := flag.String("f", "", "pptx input file")
	flag.Parse()

	oldinfo, err := os.Stat(*flagInputFile)
	if err != nil {
		log.Fatal(err)
	}

	p := NewPowerpointDoc()
	defer p.Close()
	p.ParseFile(*flagInputFile)

	//p.GetSlideMediaSize()
	p.ConvertPictures()
	//p.GetSlideMediaSize()

	outputFileName := strings.Replace(*flagInputFile, filepath.Ext(*flagInputFile), ".new.pptx", 1)
	p.SaveFile(outputFileName)

	newinfo, err := os.Stat(outputFileName)
	if err != nil {
		log.Fatal(err)
	}

	log.Println("size", *flagInputFile, oldinfo.Size(), outputFileName, newinfo.Size())
}
