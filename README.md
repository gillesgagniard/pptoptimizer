
# PPT Optimizer

A simple command line tool to attempt at reducing the size of your humongous pptx file.

It doesn't touch your original file and generates a new one, so it should be reasonably safe. Use at your own risk!

## Features

- Convert TIFF files to PNG (lossless)
- Remove unused slide layouts and masters
- Remove unused associated medias

## Usage

    go install ./cmd/pptoptimizer
    pptoptimizer -h
    pptoptimizer -f myhugepresentation.pptx -a

This command creates a hopefully smaller `myhugepresentation.new.pptx`, applying all possible optimizations.
By default, the only optimization applied is conversion of TIFF pictures to PNG.
