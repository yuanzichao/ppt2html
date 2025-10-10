package main

import (
	"flag"
	"fmt"
	"log"
	"os"

	"ppt2html/internal/pptx"
)

func main() {
	input := flag.String("input", "", "Path to the PPT or PPTX file to convert")
	output := flag.String("output", "", "Directory to store the generated HTML and assets")
	flag.Parse()

	if *input == "" || *output == "" {
		fmt.Fprintf(os.Stderr, "Usage: %s -input <file.pptx> -output <output-dir>\n", os.Args[0])
		flag.PrintDefaults()
		os.Exit(1)
	}

	if err := pptx.ConvertToHTML(*input, *output); err != nil {
		log.Fatalf("conversion failed: %v", err)
	}

	fmt.Printf("HTML presentation generated at %s\n", *output)
}
