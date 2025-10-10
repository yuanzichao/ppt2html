package pptx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"html"
	"io"
	"io/ioutil"
	"math"
	"os"
	"path"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
)

const (
	presentationNS        = "http://schemas.openxmlformats.org/presentationml/2006/main"
	drawingNS             = "http://schemas.openxmlformats.org/drawingml/2006/main"
	relationshipsNS       = "http://schemas.openxmlformats.org/package/2006/relationships"
	officeRelationshipsNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

// ConvertToHTML converts a PPT/PPTX file into a static HTML representation.
func ConvertToHTML(inputPath, outputDir string) error {
	reader, err := zip.OpenReader(inputPath)
	if err != nil {
		return fmt.Errorf("open pptx: %w", err)
	}
	defer reader.Close()

	files := map[string]*zip.File{}
	for _, f := range reader.File {
		files[f.Name] = f
	}

	presentationFile, ok := files["ppt/presentation.xml"]
	if !ok {
		return fmt.Errorf("ppt/presentation.xml not found")
	}

	presData, err := parsePresentation(presentationFile)
	if err != nil {
		return fmt.Errorf("parse presentation: %w", err)
	}

	rels, err := parseRelationships(files, "ppt/_rels/presentation.xml.rels")
	if err != nil {
		return fmt.Errorf("parse presentation rels: %w", err)
	}

	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		return fmt.Errorf("create output dir: %w", err)
	}

	assetsDir := filepath.Join(outputDir, "assets")
	if err := os.MkdirAll(assetsDir, 0o755); err != nil {
		return fmt.Errorf("create assets dir: %w", err)
	}

	var slidesHTML []string
	for idx, slideRel := range presData.SlideRels {
		target, ok := rels[slideRel]
		if !ok {
			return fmt.Errorf("missing relationship for slide %s", slideRel)
		}
		slidePath := path.Clean(path.Join("ppt", target))
		slideFile, ok := files[slidePath]
		if !ok {
			return fmt.Errorf("missing slide file %s", slidePath)
		}
		slideRelsPath := path.Join(path.Dir(slidePath), "_rels", path.Base(slidePath)+".rels")
		slideRelationships, _ := parseRelationships(files, slideRelsPath)

		slide, err := parseSlide(slideFile)
		if err != nil {
			return fmt.Errorf("parse slide %s: %w", slidePath, err)
		}

		html, err := renderSlide(idx+1, slide, slideRelationships, files, assetsDir, presData.SlideSize)
		if err != nil {
			return fmt.Errorf("render slide %d: %w", idx+1, err)
		}
		slidesHTML = append(slidesHTML, html)
	}

	fullHTML := buildDocument(slidesHTML, presData.SlideSize)
	if err := ioutil.WriteFile(filepath.Join(outputDir, "index.html"), []byte(fullHTML), 0o644); err != nil {
		return fmt.Errorf("write html: %w", err)
	}

	return nil
}

type presentationData struct {
	SlideSize slideSize
	SlideRels []string
}

type slideSize struct {
	CX float64
	CY float64
}

func parsePresentation(f *zip.File) (*presentationData, error) {
	r, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()

	decoder := xml.NewDecoder(r)
	decoder.Strict = false

	size := slideSize{
		CX: 960,
		CY: 540,
	}
	var slideRels []string

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}

		start, ok := token.(xml.StartElement)
		if !ok {
			continue
		}

		if start.Name.Space == presentationNS && start.Name.Local == "sldSz" {
			for _, attr := range start.Attr {
				switch attr.Name.Local {
				case "cx":
					if v, parseErr := strconv.ParseInt(strings.TrimSpace(attr.Value), 10, 64); parseErr == nil && v > 0 {
						size.CX = emuToPixels(v)
					}
				case "cy":
					if v, parseErr := strconv.ParseInt(strings.TrimSpace(attr.Value), 10, 64); parseErr == nil && v > 0 {
						size.CY = emuToPixels(v)
					}
				}
			}
			continue
		}

		if start.Name.Space == presentationNS && start.Name.Local == "sldId" {
			for _, attr := range start.Attr {
				if attr.Name.Local != "id" || attr.Name.Space != officeRelationshipsNS {
					continue
				}
				relID := strings.TrimSpace(attr.Value)
				if relID != "" {
					slideRels = append(slideRels, relID)
				}
			}
		}
	}

	return &presentationData{SlideSize: size, SlideRels: slideRels}, nil
}

func parseRelationships(files map[string]*zip.File, relPath string) (map[string]string, error) {
	relFile, ok := files[relPath]
	if !ok {
		return map[string]string{}, nil
	}
	r, err := relFile.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()

	decoder := xml.NewDecoder(r)
	decoder.Strict = false

	rels := make(map[string]string)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}

		start, ok := token.(xml.StartElement)
		if !ok {
			continue
		}
		if start.Name.Local != "Relationship" {
			continue
		}

		var id, target string
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "Id":
				id = strings.TrimSpace(attr.Value)
			case "Target":
				target = strings.TrimSpace(attr.Value)
			}
		}
		if id == "" || target == "" {
			continue
		}
		rels[id] = target
	}

	return rels, nil
}

type slide struct {
	Background *background
	Shapes     []shape
	Pictures   []picture
}

type background struct {
	Color string
}

type shape struct {
	Name      string
	Transform transform
	FillColor string
	LineColor string
	LineWidth float64
	Text      []paragraph
}

type picture struct {
	Name      string
	Transform transform
	Embed     string
}

type transform struct {
	X, Y, CX, CY float64
}

type paragraph struct {
	Align string
	Runs  []textRun
}

type textRun struct {
	Text      string
	FontSize  float64
	Bold      bool
	Italic    bool
	Underline bool
	Color     string
}

func parseSlide(f *zip.File) (*slide, error) {
	r, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer r.Close()

	decoder := xml.NewDecoder(r)
	decoder.Strict = false

	type rawSlide struct {
		XMLName xml.Name `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}sld"`
		CSld    struct {
			Background *struct {
				BgPr *struct {
					SolidFill *struct {
						Srgb *struct {
							Val string `xml:"val,attr"`
						} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr"`
					} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill"`
				} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}bgPr"`
			} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}bg"`
			SpTree struct {
				Shapes []struct {
					NVSpPr struct {
						CNvPr struct {
							Name string `xml:"name,attr"`
						} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr"`
					} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}nvSpPr"`
					SpPr   *shapeProps `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}spPr"`
					TxBody *textBody   `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}txBody"`
				} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}sp"`
				Pictures []struct {
					NVPicPr struct {
						CNvPr struct {
							Name string `xml:"name,attr"`
						} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr"`
					} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}nvPicPr"`
					BlipFill *struct {
						Blip struct {
							Embed string `xml:"{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed,attr"`
						} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}blip"`
					} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}blipFill"`
					SpPr *shapeProps `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}spPr"`
				} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}pic"`
			} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}spTree"`
		} `xml:"{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"`
	}

	var data rawSlide
	if err := decoder.Decode(&data); err != nil {
		return nil, err
	}

	result := &slide{}

	if data.CSld.Background != nil && data.CSld.Background.BgPr != nil && data.CSld.Background.BgPr.SolidFill != nil && data.CSld.Background.BgPr.SolidFill.Srgb != nil {
		result.Background = &background{Color: formatColor(data.CSld.Background.BgPr.SolidFill.Srgb.Val)}
	}

	for _, rawShape := range data.CSld.SpTree.Shapes {
		sh := shape{Name: rawShape.NVSpPr.CNvPr.Name}
		if rawShape.SpPr != nil {
			sh.Transform = rawShape.SpPr.Transform()
			sh.FillColor = rawShape.SpPr.FillColor()
			sh.LineColor, sh.LineWidth = rawShape.SpPr.LineStyle()
		}
		if rawShape.TxBody != nil {
			sh.Text = rawShape.TxBody.Paragraphs()
		}
		if len(sh.Text) == 0 && sh.FillColor == "" && sh.LineColor == "" {
			continue
		}
		result.Shapes = append(result.Shapes, sh)
	}

	for _, rawPic := range data.CSld.SpTree.Pictures {
		pic := picture{Name: rawPic.NVPicPr.CNvPr.Name}
		if rawPic.SpPr != nil {
			pic.Transform = rawPic.SpPr.Transform()
		}
		if rawPic.BlipFill != nil {
			pic.Embed = rawPic.BlipFill.Blip.Embed
		}
		result.Pictures = append(result.Pictures, pic)
	}

	return result, nil
}

type shapeProps struct {
	Xfrm *struct {
		Off struct {
			X string `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}off>x,attr"`
			Y string `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}off>y,attr"`
		}
		Ext struct {
			CX string `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}ext>cx,attr"`
			CY string `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}ext>cy,attr"`
		}
	} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm"`
	SolidFill *struct {
		Srgb struct {
			Val string `xml:"val,attr"`
		} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr"`
	} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill"`
	Line *struct {
		W         string `xml:"w,attr"`
		SolidFill *struct {
			Srgb struct {
				Val string `xml:"val,attr"`
			} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr"`
		} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill"`
	} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}ln"`
}

func (p *shapeProps) Transform() transform {
	tr := transform{}
	if p == nil || p.Xfrm == nil {
		return tr
	}
	if p.Xfrm.Off.X != "" {
		tr.X = emuToPixels(mustParseInt(p.Xfrm.Off.X))
	}
	if p.Xfrm.Off.Y != "" {
		tr.Y = emuToPixels(mustParseInt(p.Xfrm.Off.Y))
	}
	if p.Xfrm.Ext.CX != "" {
		tr.CX = emuToPixels(mustParseInt(p.Xfrm.Ext.CX))
	}
	if p.Xfrm.Ext.CY != "" {
		tr.CY = emuToPixels(mustParseInt(p.Xfrm.Ext.CY))
	}
	return tr
}

func (p *shapeProps) FillColor() string {
	if p == nil || p.SolidFill == nil {
		return ""
	}
	return formatColor(p.SolidFill.Srgb.Val)
}

func (p *shapeProps) LineStyle() (string, float64) {
	if p == nil || p.Line == nil {
		return "", 0
	}
	color := ""
	if p.Line.SolidFill != nil {
		color = formatColor(p.Line.SolidFill.Srgb.Val)
	}
	width := 0.0
	if p.Line.W != "" {
		width = math.Max(1, emuToPixels(mustParseInt(p.Line.W)))
	}
	return color, width
}

type textBody struct {
	Paras []struct {
		PPr *struct {
			Algn string `xml:"algn,attr"`
		} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}pPr"`
		Runs []struct {
			RPr  *runProperties `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"`
			Text struct {
				Data string `xml:",chardata"`
			} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}t"`
			Breaks []struct{} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}br"`
		} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}r"`
		Fields []struct {
			RPr  *runProperties `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"`
			Text struct {
				Data string `xml:",chardata"`
			} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}t"`
		} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}fld"`
	} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}p"`
}

type runProperties struct {
	Sz        int    `xml:"sz,attr"`
	Bold      bool   `xml:"b,attr"`
	Italic    bool   `xml:"i,attr"`
	Underline string `xml:"u,attr"`
	SolidFill *struct {
		Srgb struct {
			Val string `xml:"val,attr"`
		} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr"`
	} `xml:"{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill"`
}

func (tb *textBody) Paragraphs() []paragraph {
	if tb == nil {
		return nil
	}
	var result []paragraph
	for _, p := range tb.Paras {
		para := paragraph{}
		if p.PPr != nil {
			para.Align = mapAlignment(p.PPr.Algn)
		}
		for _, run := range p.Runs {
			if run.Text.Data == "" && len(run.Breaks) == 0 {
				continue
			}
			tr := textRun{Text: run.Text.Data}
			if run.RPr != nil {
				if run.RPr.Sz != 0 {
					tr.FontSize = fontSizeToPixels(run.RPr.Sz)
				}
				tr.Bold = run.RPr.Bold
				tr.Italic = run.RPr.Italic
				tr.Underline = run.RPr.Underline != ""
				if run.RPr.SolidFill != nil {
					tr.Color = formatColor(run.RPr.SolidFill.Srgb.Val)
				}
			}
			para.Runs = append(para.Runs, tr)
			for range run.Breaks {
				para.Runs = append(para.Runs, textRun{Text: "<br/>", FontSize: tr.FontSize})
			}
		}
		for _, fld := range p.Fields {
			tr := textRun{Text: fld.Text.Data}
			if fld.RPr != nil {
				if fld.RPr.Sz != 0 {
					tr.FontSize = fontSizeToPixels(fld.RPr.Sz)
				}
				tr.Bold = fld.RPr.Bold
				tr.Italic = fld.RPr.Italic
				tr.Underline = fld.RPr.Underline != ""
				if fld.RPr.SolidFill != nil {
					tr.Color = formatColor(fld.RPr.SolidFill.Srgb.Val)
				}
			}
			para.Runs = append(para.Runs, tr)
		}
		if len(para.Runs) > 0 {
			result = append(result, para)
		}
	}
	return result
}

func renderSlide(index int, s *slide, rels map[string]string, files map[string]*zip.File, assetsDir string, size slideSize) (string, error) {
	var buf bytes.Buffer
	slideWidth := size.CX
	slideHeight := size.CY

	bgStyle := ""
	if s.Background != nil && s.Background.Color != "" {
		bgStyle = fmt.Sprintf("background-color:%s;", s.Background.Color)
	}

	fmt.Fprintf(&buf, `<section class="slide" style="width:%.2fpx;height:%.2fpx;%s">`, slideWidth, slideHeight, bgStyle)

	sort.Slice(s.Shapes, func(i, j int) bool {
		if s.Shapes[i].Transform.Y == s.Shapes[j].Transform.Y {
			return s.Shapes[i].Transform.X < s.Shapes[j].Transform.X
		}
		return s.Shapes[i].Transform.Y < s.Shapes[j].Transform.Y
	})

	for _, shape := range s.Shapes {
		renderShape(&buf, shape)
	}

	for imgIdx, pic := range s.Pictures {
		if pic.Embed == "" {
			continue
		}
		target, ok := rels[pic.Embed]
		if !ok {
			continue
		}
		sourcePath := path.Clean(path.Join("ppt/slides", target))
		if !strings.HasPrefix(sourcePath, "ppt/") {
			sourcePath = path.Clean(path.Join("ppt/slides", "../", target))
		}
		sourcePath = path.Clean(sourcePath)
		if !strings.HasPrefix(sourcePath, "ppt/") {
			sourcePath = path.Join("ppt", target)
		}
		file, ok := files[sourcePath]
		if !ok {
			continue
		}
		ext := strings.ToLower(path.Ext(sourcePath))
		if ext == "" {
			ext = ".bin"
		}
		assetName := fmt.Sprintf("slide%d_image%d%s", index, imgIdx+1, ext)
		assetPath := filepath.Join(assetsDir, assetName)
		if err := copyZipFile(file, assetPath); err != nil {
			return "", err
		}
		fmt.Fprintf(&buf, `<img class="picture" src="assets/%s" alt="%s" style="position:absolute;left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx;"/>`, html.EscapeString(assetName), html.EscapeString(pic.Name), pic.Transform.X, pic.Transform.Y, pic.Transform.CX, pic.Transform.CY)
	}

	buf.WriteString("</section>")
	return buf.String(), nil
}

func renderShape(buf *bytes.Buffer, shape shape) {
	style := fmt.Sprintf("left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx;", shape.Transform.X, shape.Transform.Y, shape.Transform.CX, shape.Transform.CY)
	if shape.FillColor != "" {
		style += fmt.Sprintf("background-color:%s;", shape.FillColor)
	}
	if shape.LineColor != "" && shape.LineWidth > 0 {
		style += fmt.Sprintf("border:%.2fpx solid %s;", shape.LineWidth, shape.LineColor)
	}

	buf.WriteString(`<div class="shape" style="position:absolute;` + style + `">`)
	buf.WriteString(`<div class="shape-content" style="width:100%;height:100%;display:flex;flex-direction:column;justify-content:flex-start;">`)
	for _, para := range shape.Text {
		align := para.Align
		if align == "" {
			align = "left"
		}
		buf.WriteString(`<div class="paragraph" style="margin:0;flex:0 0 auto;text-align:` + align + `;white-space:pre-wrap;">`)
		for _, run := range para.Runs {
			if run.Text == "<br/>" {
				buf.WriteString("<br/>")
				continue
			}
			styleParts := []string{}
			if run.FontSize > 0 {
				styleParts = append(styleParts, fmt.Sprintf("font-size:%.2fpx", run.FontSize))
			}
			if run.Color != "" {
				styleParts = append(styleParts, fmt.Sprintf("color:%s", run.Color))
			}
			if run.Bold {
				styleParts = append(styleParts, "font-weight:700")
			}
			if run.Italic {
				styleParts = append(styleParts, "font-style:italic")
			}
			if run.Underline {
				styleParts = append(styleParts, "text-decoration:underline")
			}
			styleAttr := ""
			if len(styleParts) > 0 {
				styleAttr = ` style="` + strings.Join(styleParts, ";") + `"`
			}
			buf.WriteString(`<span` + styleAttr + `>` + escapeText(run.Text) + `</span>`)
		}
		buf.WriteString(`</div>`)
	}
	buf.WriteString(`</div></div>`)
}

func buildDocument(slides []string, size slideSize) string {
	var buf bytes.Buffer
	buf.WriteString("<!DOCTYPE html><html><head><meta charset=\"utf-8\"/>")
	buf.WriteString("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"/>")
	buf.WriteString("<title>PPT to HTML</title>")
	buf.WriteString(`<style>
        body { margin:0; background:#f3f4f6; font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif; }
        .deck { max-width: `)
	buf.WriteString(fmt.Sprintf("%.0fpx", size.CX+40))
	buf.WriteString(`; margin: 0 auto; padding:20px; box-sizing:border-box; }
        .slide { position:relative; margin:20px auto; box-shadow:0 10px 30px rgba(0,0,0,0.12); border-radius:8px; overflow:hidden; background:white; }
        .shape { position:absolute; box-sizing:border-box; padding:8px; }
        .shape-content { width:100%; height:100%; }
        .paragraph { font-size:16px; line-height:1.3; }
        img.picture { object-fit:cover; }
    </style>`)
	buf.WriteString("</head><body><main class=\"deck\">")
	for _, slide := range slides {
		buf.WriteString(slide)
	}
	buf.WriteString("</main></body></html>")
	return buf.String()
}

func escapeText(text string) string {
	return strings.ReplaceAll(html.EscapeString(text), "\n", "<br/>")
}

func formatColor(hex string) string {
	if hex == "" {
		return ""
	}
	hex = strings.TrimSpace(hex)
	if len(hex) == 6 {
		return "#" + strings.ToUpper(hex)
	}
	if len(hex) == 8 {
		return "#" + strings.ToUpper(hex[2:])
	}
	return "#" + strings.ToUpper(hex)
}

func emuToPixels(v int64) float64 {
	return float64(v) * 96.0 / 914400.0
}

func fontSizeToPixels(v int) float64 {
	pt := float64(v) / 100.0
	return pt * 96.0 / 72.0
}

func mustParseInt(s string) int64 {
	if s == "" {
		return 0
	}
	v, err := strconv.ParseInt(s, 10, 64)
	if err != nil {
		return 0
	}
	return v
}

func copyZipFile(f *zip.File, dest string) error {
	if err := os.MkdirAll(filepath.Dir(dest), 0o755); err != nil {
		return err
	}
	src, err := f.Open()
	if err != nil {
		return err
	}
	defer src.Close()
	out, err := os.Create(dest)
	if err != nil {
		return err
	}
	defer out.Close()
	if _, err := io.Copy(out, src); err != nil {
		return err
	}
	return nil
}

func mapAlignment(algn string) string {
	switch algn {
	case "ctr":
		return "center"
	case "r":
		return "right"
	case "just":
		return "justify"
	default:
		return "left"
	}
}
