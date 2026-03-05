package doc

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
)

var errDocxInvalid = errors.New("word/document.xml not found in docx file")

// ParseDocx extracts plain text from a Microsoft Word .docx file.
// It accepts an io.Reader containing the .docx file data and returns
// a reader with the extracted plain text content.
func ParseDocx(r io.Reader) (io.Reader, error) {
	ra, size, err := toMemoryBuffer(r)
	if err != nil {
		return nil, wrapError(err)
	}
	defer ra.Close()

	zr, err := zip.NewReader(ra, size)
	if err != nil {
		return nil, wrapError(err)
	}

	for _, f := range zr.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				return nil, wrapError(err)
			}
			defer rc.Close()
			return parseDocxXML(rc)
		}
	}

	return nil, wrapError(errDocxInvalid)
}

// parseDocxXML extracts text from a word/document.xml reader.
// It outputs paragraphs separated by newlines.
func parseDocxXML(r io.Reader) (io.Reader, error) {
	var buf bytes.Buffer
	decoder := xml.NewDecoder(r)

	const wNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

	inText := false

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			if t.Name.Space == wNS {
				switch t.Name.Local {
				case "t":
					inText = true
				case "br", "cr":
					buf.WriteByte('\n')
				case "tab":
					buf.WriteByte('\t')
				}
			}
		case xml.EndElement:
			if t.Name.Space == wNS {
				switch t.Name.Local {
				case "t":
					inText = false
				case "p":
					buf.WriteByte('\n')
				}
			}
		case xml.CharData:
			if inText {
				buf.Write([]byte(t))
			}
		}
	}

	return &buf, nil
}
