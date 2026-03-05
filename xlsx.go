package doc

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"strconv"
	"strings"
)

var errXlsxInvalid = errors.New("no worksheet files found in xlsx file")

// ParseXlsx extracts plain text from a Microsoft Excel .xlsx file.
// It accepts an io.Reader containing the .xlsx file data and returns
// a reader with the extracted plain text content, with cells separated
// by tabs and rows separated by newlines.
func ParseXlsx(r io.Reader) (io.Reader, error) {
	ra, size, err := toMemoryBuffer(r)
	if err != nil {
		return nil, wrapError(err)
	}
	defer ra.Close()

	zr, err := zip.NewReader(ra, size)
	if err != nil {
		return nil, wrapError(err)
	}

	sharedStrings, err := parseXlsxSharedStrings(zr)
	if err != nil {
		return nil, wrapError(err)
	}

	var buf bytes.Buffer
	sheetFound := false
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/sheet") && strings.HasSuffix(f.Name, ".xml") {
			sheetFound = true
			rc, err := f.Open()
			if err != nil {
				return nil, wrapError(err)
			}
			err = parseXlsxSheet(rc, &buf, sharedStrings)
			rc.Close()
			if err != nil {
				return nil, wrapError(err)
			}
		}
	}

	if !sheetFound {
		return nil, wrapError(errXlsxInvalid)
	}

	return &buf, nil
}

// parseXlsxSharedStrings reads the shared strings table from xl/sharedStrings.xml.
func parseXlsxSharedStrings(zr *zip.Reader) ([]string, error) {
	for _, f := range zr.File {
		if f.Name == "xl/sharedStrings.xml" {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return parseSharedStringsXML(rc)
		}
	}
	// It's valid to have no shared strings (e.g., numeric-only spreadsheet)
	return nil, nil
}

// parseSharedStringsXML parses the XML and returns a slice of shared strings.
func parseSharedStringsXML(r io.Reader) ([]string, error) {
	const ssNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

	var result []string
	var currentStr strings.Builder
	inT := false

	decoder := xml.NewDecoder(r)
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
			if t.Name.Space == ssNS {
				switch t.Name.Local {
				case "si":
					currentStr.Reset()
				case "t":
					inT = true
				}
			}
		case xml.EndElement:
			if t.Name.Space == ssNS {
				switch t.Name.Local {
				case "si":
					result = append(result, currentStr.String())
				case "t":
					inT = false
				}
			}
		case xml.CharData:
			if inT {
				currentStr.Write([]byte(t))
			}
		}
	}

	return result, nil
}

// parseXlsxSheet parses a worksheet XML file and writes its text content to buf.
func parseXlsxSheet(r io.Reader, buf *bytes.Buffer, sharedStrings []string) error {
	const ssNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

	decoder := xml.NewDecoder(r)

	var cellType string  // current cell type attribute
	var cellValue strings.Builder
	inV := false        // inside <v> (value) element
	inIs := false       // inside <is> (inline string) element
	inT := false        // inside <t> within <is>
	firstCellInRow := true
	firstRow := true

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		switch t := token.(type) {
		case xml.StartElement:
			if t.Name.Space == ssNS {
				switch t.Name.Local {
				case "row":
					if !firstRow {
						buf.WriteByte('\n')
					}
					firstRow = false
					firstCellInRow = true
				case "c":
					cellType = ""
					cellValue.Reset()
					for _, attr := range t.Attr {
						if attr.Name.Local == "t" {
							cellType = attr.Value
						}
					}
					if !firstCellInRow {
						buf.WriteByte('\t')
					}
					firstCellInRow = false
				case "v":
					inV = true
				case "is":
					inIs = true
				case "t":
					if inIs {
						inT = true
					}
				}
			}
		case xml.EndElement:
			if t.Name.Space == ssNS {
				switch t.Name.Local {
				case "v":
					inV = false
					if cellType == "s" {
						// shared string index
						idx, err := strconv.Atoi(strings.TrimSpace(cellValue.String()))
						if err == nil && idx >= 0 && idx < len(sharedStrings) {
							buf.WriteString(sharedStrings[idx])
						}
					} else {
						buf.WriteString(cellValue.String())
					}
					cellValue.Reset()
				case "is":
					inIs = false
				case "t":
					if inIs {
						inT = false
					}
				}
			}
		case xml.CharData:
			if inV {
				cellValue.Write([]byte(t))
			} else if inT {
				buf.Write([]byte(t))
			}
		}
	}

	return nil
}
