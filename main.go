package main

import (
	"bytes"
	"encoding/base64"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"mime"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strings"
)

type converter struct {
	inputPath  string
	outputDir  string
	baseName   string
	imagesDir  string
	svgTempDir string
}

func newConverter(inputPath, outputDir string) (*converter, error) {
	if inputPath == "" {
		return nil, errors.New("input path is required")
	}
	if outputDir == "" {
		return nil, errors.New("output directory is required")
	}

	base := strings.TrimSuffix(filepath.Base(inputPath), filepath.Ext(inputPath))
	imagesDir := filepath.Join(outputDir, "images")

	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		return nil, fmt.Errorf("failed to create output directory: %w", err)
	}
	if err := os.MkdirAll(imagesDir, 0o755); err != nil {
		return nil, fmt.Errorf("failed to create images directory: %w", err)
	}

	tempDir, err := os.MkdirTemp("", "ppt2html-svg-")
	if err != nil {
		return nil, fmt.Errorf("failed to create temporary directory: %w", err)
	}

	return &converter{
		inputPath:  inputPath,
		outputDir:  outputDir,
		baseName:   base,
		imagesDir:  imagesDir,
		svgTempDir: tempDir,
	}, nil
}

func (c *converter) cleanup() {
	if c.svgTempDir != "" {
		_ = os.RemoveAll(c.svgTempDir)
	}
}

func (c *converter) convertToSVG() ([]string, error) {
	cmd := exec.Command("libreoffice", "--headless", "--convert-to", "svg", "--outdir", c.svgTempDir, c.inputPath)
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr

	if err := cmd.Run(); err != nil {
		return nil, fmt.Errorf("failed to convert to svg: %w", err)
	}

	entries, err := os.ReadDir(c.svgTempDir)
	if err != nil {
		return nil, fmt.Errorf("failed to read svg directory: %w", err)
	}

	var svgFiles []string
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		if strings.EqualFold(filepath.Ext(entry.Name()), ".svg") {
			svgFiles = append(svgFiles, filepath.Join(c.svgTempDir, entry.Name()))
		}
	}

	sort.Strings(svgFiles)

	if len(svgFiles) == 0 {
		return nil, errors.New("no svg files were generated")
	}

	return svgFiles, nil
}

func (c *converter) processSVG(svgPath string, pageIndex int) error {
	content, err := os.ReadFile(svgPath)
	if err != nil {
		return fmt.Errorf("failed to read svg file %s: %w", svgPath, err)
	}

	processedSVG, err := c.extractImages(content, pageIndex)
	if err != nil {
		return fmt.Errorf("failed to process images in %s: %w", svgPath, err)
	}

	htmlFile := filepath.Join(c.outputDir, fmt.Sprintf("%s-page-%d.html", c.baseName, pageIndex+1))
	if err := writeHTML(htmlFile, c.baseName, processedSVG); err != nil {
		return fmt.Errorf("failed to write html file %s: %w", htmlFile, err)
	}

	return nil
}

func (c *converter) extractImages(svgContent []byte, pageIndex int) ([]byte, error) {
	decoder := xml.NewDecoder(bytes.NewReader(svgContent))
	var buffer bytes.Buffer
	encoder := xml.NewEncoder(&buffer)
	encoder.Indent("", "  ")

	imageCounter := 0

	for {
		token, err := decoder.Token()
		if err != nil {
			if errors.Is(err, io.EOF) {
				break
			}
			return nil, fmt.Errorf("failed to parse svg: %w", err)
		}

		switch elem := token.(type) {
		case xml.StartElement:
			if elem.Name.Local == "image" {
				for i, attr := range elem.Attr {
					if (attr.Name.Local == "href" && (attr.Name.Space == "xlink" || attr.Name.Space == "")) || attr.Name.Local == "href" {
						savedAttr, err := c.saveImage(attr.Value, pageIndex, imageCounter)
						if err != nil {
							return nil, err
						}
						if savedAttr != "" {
							elem.Attr[i].Value = savedAttr
							if attr.Name.Space == "" {
								elem.Attr[i].Name.Space = "xlink"
							}
							imageCounter++
						}
					}
				}
			}
			if err := encoder.EncodeToken(elem); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if err := encoder.EncodeToken(elem); err != nil {
				return nil, err
			}
		case xml.CharData:
			if err := encoder.EncodeToken(elem); err != nil {
				return nil, err
			}
		case xml.Comment:
			if err := encoder.EncodeToken(elem); err != nil {
				return nil, err
			}
		case xml.Directive:
			if err := encoder.EncodeToken(elem); err != nil {
				return nil, err
			}
		case xml.ProcInst:
			// Skip XML processing instructions to avoid invalid HTML content.
			continue
		default:
			if err := encoder.EncodeToken(elem); err != nil {
				return nil, err
			}
		}
	}

	if err := encoder.Flush(); err != nil {
		return nil, err
	}

	return buffer.Bytes(), nil
}

func (c *converter) saveImage(dataURI string, pageIndex, imageIndex int) (string, error) {
	if !strings.HasPrefix(dataURI, "data:") {
		return "", nil
	}

	metaAndData := strings.SplitN(dataURI[5:], ",", 2)
	if len(metaAndData) != 2 {
		return "", fmt.Errorf("invalid data URI format")
	}

	meta := strings.TrimSpace(metaAndData[0])
	dataPart := strings.TrimSpace(metaAndData[1])

	if !strings.Contains(meta, ";base64") {
		return "", fmt.Errorf("unsupported data URI encoding: %s", meta)
	}

	mimeType := strings.Split(meta, ";")[0]

	decoded, err := base64.StdEncoding.DecodeString(dataPart)
	if err != nil {
		return "", fmt.Errorf("failed to decode base64 image: %w", err)
	}

	ext := fileExtensionForMime(mimeType)
	imageName := fmt.Sprintf("%s-page-%d-image-%d%s", c.baseName, pageIndex+1, imageIndex+1, ext)
	imagePath := filepath.Join(c.imagesDir, imageName)

	if err := os.WriteFile(imagePath, decoded, 0o644); err != nil {
		return "", fmt.Errorf("failed to write image %s: %w", imagePath, err)
	}

	return filepath.ToSlash(filepath.Join("images", imageName)), nil
}

func fileExtensionForMime(mimeType string) string {
	if mimeType == "" {
		return ".bin"
	}

	if !strings.Contains(mimeType, "/") {
		return ".bin"
	}

	exts, err := mime.ExtensionsByType(mimeType)
	if err == nil && len(exts) > 0 {
		return exts[0]
	}

	switch mimeType {
	case "image/jpeg":
		return ".jpg"
	case "image/png":
		return ".png"
	case "image/gif":
		return ".gif"
	case "image/svg+xml":
		return ".svg"
	default:
		parts := strings.Split(mimeType, "/")
		return "." + parts[len(parts)-1]
	}
}

func writeHTML(htmlPath, title string, svgContent []byte) error {
	var builder strings.Builder
	builder.WriteString("<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n")
	builder.WriteString(fmt.Sprintf("<title>%s</title>\n", escapeHTMLTitle(title)))
	builder.WriteString("<style>body{margin:0;padding:0;background:#fff;}svg{display:block;margin:auto;height:100vh;width:100vw;}</style>\n")
	builder.WriteString("</head>\n<body>\n")
	builder.Write(svgContent)
	builder.WriteString("\n</body>\n</html>\n")

	return os.WriteFile(htmlPath, []byte(builder.String()), 0o644)
}

func escapeHTMLTitle(title string) string {
	replacer := strings.NewReplacer(
		"&", "&amp;",
		"<", "&lt;",
		">", "&gt;",
		"\"", "&quot;",
		"'", "&#39;",
	)
	return replacer.Replace(title)
}

func run(inputPath, outputDir string) error {
	conv, err := newConverter(inputPath, outputDir)
	if err != nil {
		return err
	}
	defer conv.cleanup()

	svgFiles, err := conv.convertToSVG()
	if err != nil {
		return err
	}

	for idx, svg := range svgFiles {
		if err := conv.processSVG(svg, idx); err != nil {
			return err
		}
	}

	return nil
}

func main() {
	if len(os.Args) != 3 {
		fmt.Fprintf(os.Stderr, "Usage: %s <input-file> <output-directory>\n", filepath.Base(os.Args[0]))
		os.Exit(1)
	}

	input := os.Args[1]
	output := os.Args[2]

	if err := run(input, output); err != nil {
		fmt.Fprintln(os.Stderr, "Error:", err)
		os.Exit(1)
	}
}
