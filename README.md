
# PPT Optimizer

A simple command line tool to attempt at reducing the size of your humongous pptx file.

## Features

- Convert TIFF files to PNG (lossless)

## Usage

    go install ./cmd/pptoptimizer
    pptoptimizer -f myhugepresentation.pptx

This command creates a hopefully smaller `myhugepresentation.new.pptx`.
