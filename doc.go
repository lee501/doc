package doc

import (
	"bytes"
	"encoding/binary"
	"errors"
	"io"
	"unicode/utf8"

	"github.com/mattetti/filebuffer"
	"github.com/richardlehane/mscfb"
	"golang.org/x/text/encoding/simplifiedchinese"
)

var (
	errTable           = errors.New("cannot find table stream")
	errDocEmpty        = errors.New("WordDocument not found")
	errDocShort        = errors.New("wordDoc block too short")
	errInvalidArgument = errors.New("invalid table and/or fib")
)

type allReader interface {
	io.Closer
	io.ReaderAt
	io.ReadSeeker
}

func wrapError(e error) error {
	return errors.New("Error processing file: " + e.Error())
}

// ParseDoc converts a standard io.Reader from a Microsoft Word
// .doc binary file and returns a reader (actually a bytes.Buffer)
// which will output the plain text found in the .doc file
func ParseDoc(r io.Reader) (io.Reader, error) {
	ra, ok := r.(io.ReaderAt)
	if !ok {
		ra, _, err := toMemoryBuffer(r)
		if err != nil {
			return nil, wrapError(err)
		}
		defer ra.Close()
	}

	d, err := mscfb.New(ra)
	if err != nil {
		return nil, wrapError(err)
	}

	wordDoc, table0, table1 := getWordDocAndTables(d)
	fib, err := getFib(wordDoc)
	if err != nil {
		return nil, wrapError(err)
	}

	table := getActiveTable(table0, table1, fib)
	if table == nil {
		return nil, wrapError(errTable)
	}

	clx, err := getClx(table, fib)
	if err != nil {
		return nil, wrapError(err)
	}

	return getText(wordDoc, clx, fib)
}

func toMemoryBuffer(r io.Reader) (allReader, int64, error) {
	var b bytes.Buffer
	size, err := b.ReadFrom(r)
	if err != nil {
		return nil, 0, err
	}
	fb := filebuffer.New(b.Bytes())
	return fb, size, nil
}

func getText(wordDoc *mscfb.File, clx *clx, fib *fib) (io.Reader, error) {
	var buf bytes.Buffer
	for i := 0; i < len(clx.pcdt.PlcPcd.aPcd); i++ {
		pcd := clx.pcdt.PlcPcd.aPcd[i]
		cp := clx.pcdt.PlcPcd.aCP[i]
		cpNext := clx.pcdt.PlcPcd.aCP[i+1]

		var start, end int
		if pcd.fc.fCompressed {
			start = pcd.fc.fc / 2
			end = start + (cpNext - cp)
		} else {
			start = pcd.fc.fc
			end = start + 2*(cpNext-cp)
		}

		b := make([]byte, end-start)
		_, err := wordDoc.ReadAt(b, int64(start))
		if err != nil {
			return nil, err
		}

		err = translateText(b, &buf, pcd.fc.fCompressed, fib)
		if err != nil {
			return nil, err
		}
	}
	return &buf, nil
}

func translateText(b []byte, buf *bytes.Buffer, fCompressed bool, fib *fib) error {
	if fCompressed {
		// Handle compressed (single-byte) text
		return translateCompressedText(b, buf)
	} else {
		// Handle uncompressed (double-byte) text - typically Unicode
		return translateUncompressedText(b, buf, fib)
	}
}

func translateCompressedText(b []byte, buf *bytes.Buffer) error {
	fieldLevel := 0
	var isFieldChar bool

	for cIndex := range b {
		// Handle special field characters (section 2.8.25)
		if b[cIndex] == 0x13 {
			isFieldChar = true
			fieldLevel++
			continue
		} else if b[cIndex] == 0x14 {
			isFieldChar = false
			continue
		} else if b[cIndex] == 0x15 {
			isFieldChar = false
			fieldLevel--
			continue
		} else if isFieldChar {
			continue
		}

		if b[cIndex] == 7 { // table column separator
			buf.WriteByte(' ')
			continue
		} else if b[cIndex] < 32 && b[cIndex] != 9 && b[cIndex] != 10 && b[cIndex] != 13 {
			// skip non-printable ASCII characters
			continue
		}

		// Handle compressed characters with special mappings
		converted := replaceCompressed(b[cIndex])
		buf.Write(converted)
	}
	return nil
}

func translateUncompressedText(b []byte, buf *bytes.Buffer, fib *fib) error {
	fieldLevel := 0
	var isFieldChar bool

	// Process bytes in pairs for Unicode characters
	for i := 0; i < len(b)-1; i += 2 {
		// Read as little-endian uint16
		char := binary.LittleEndian.Uint16(b[i : i+2])

		// Handle special field characters
		if char == 0x13 {
			isFieldChar = true
			fieldLevel++
			continue
		} else if char == 0x14 {
			isFieldChar = false
			continue
		} else if char == 0x15 {
			isFieldChar = false
			fieldLevel--
			continue
		} else if isFieldChar {
			continue
		}

		if char == 7 { // table column separator
			buf.WriteByte(' ')
			continue
		} else if char < 32 && char != 9 && char != 10 && char != 13 {
			// skip non-printable characters
			continue
		}

		// Convert Unicode code point to UTF-8
		if char <= 0x7F {
			// ASCII range
			buf.WriteByte(byte(char))
		} else {
			// Unicode character - convert to UTF-8
			rune := rune(char)
			if utf8.ValidRune(rune) {
				utf8Bytes := make([]byte, 4)
				n := utf8.EncodeRune(utf8Bytes, rune)
				buf.Write(utf8Bytes[:n])
			}
		}
	}
	return nil
}

// Enhanced character replacement for compressed text
func replaceCompressed(char byte) []byte {
	var v uint16
	switch char {
	case 0x82:
		v = 0x201A // Single Low-9 Quotation Mark
	case 0x83:
		v = 0x0192 // Latin Small Letter F With Hook
	case 0x84:
		v = 0x201E // Double Low-9 Quotation Mark
	case 0x85:
		v = 0x2026 // Horizontal Ellipsis
	case 0x86:
		v = 0x2020 // Dagger
	case 0x87:
		v = 0x2021 // Double Dagger
	case 0x88:
		v = 0x02C6 // Modifier Letter Circumflex Accent
	case 0x89:
		v = 0x2030 // Per Mille Sign
	case 0x8A:
		v = 0x0160 // Latin Capital Letter S With Caron
	case 0x8B:
		v = 0x2039 // Single Left-Pointing Angle Quotation Mark
	case 0x8C:
		v = 0x0152 // Latin Capital Ligature OE
	case 0x91:
		v = 0x2018 // Left Single Quotation Mark
	case 0x92:
		v = 0x2019 // Right Single Quotation Mark
	case 0x93:
		v = 0x201C // Left Double Quotation Mark
	case 0x94:
		v = 0x201D // Right Double Quotation Mark
	case 0x95:
		v = 0x2022 // Bullet
	case 0x96:
		v = 0x2013 // En Dash
	case 0x97:
		v = 0x2014 // Em Dash
	case 0x98:
		v = 0x02DC // Small Tilde
	case 0x99:
		v = 0x2122 // Trade Mark Sign
	case 0x9A:
		v = 0x0161 // Latin Small Letter S With Caron
	case 0x9B:
		v = 0x203A // Single Right-Pointing Angle Quotation Mark
	case 0x9C:
		v = 0x0153 // Latin Small Ligature OE
	case 0x9F:
		v = 0x0178 // Latin Capital Letter Y With Diaeresis
	default:
		// For characters in the range 0x80-0xFF that don't have special mappings,
		// treat as potential ANSI/CP1252 encoding
		if char >= 0x80 {
			return handleANSICharacter(char)
		}
		return []byte{char}
	}

	// Convert Unicode code point to UTF-8
	rune := rune(v)
	utf8Bytes := make([]byte, 4)
	n := utf8.EncodeRune(utf8Bytes, rune)
	return utf8Bytes[:n]
}

// Handle ANSI/CP1252 characters that might be Chinese characters in some encodings
func handleANSICharacter(char byte) []byte {
	// Try to decode as GB2312/GBK for Chinese characters
	if char >= 0xA1 {
		// This is a simplified approach - in practice you might need more context
		// to determine the correct encoding
		decoder := simplifiedchinese.GBK.NewDecoder()

		// For single-byte character, we need to check if it's part of a multi-byte sequence
		// This is a basic implementation - you might need to buffer characters
		input := []byte{char}
		output := make([]byte, 10)

		nDst, nSrc, err := decoder.Transform(output, input, false)
		if err == nil && nSrc > 0 && nDst > 0 {
			return output[:nDst]
		}
	}

	// Fallback to original character
	return []byte{char}
}

// Helper function to detect potential Chinese text encoding
func detectChineseEncoding(data []byte, fib *fib) bool {
	// Check FIB for language information
	// This is a simplified check - you might want to examine fib.lid (language ID)

	// Look for high-byte characters that might indicate Chinese text
	highByteCount := 0
	for _, b := range data {
		if b >= 0x80 {
			highByteCount++
		}
	}

	// If more than 50% are high-byte characters, likely Chinese or other multibyte encoding
	return float64(highByteCount)/float64(len(data)) > 0.5
}

func getWordDocAndTables(r *mscfb.Reader) (*mscfb.File, *mscfb.File, *mscfb.File) {
	var wordDoc, table0, table1 *mscfb.File
	for i := 0; i < len(r.File); i++ {
		stream := r.File[i]

		switch stream.Name {
		case "WordDocument":
			wordDoc = stream
		case "0Table":
			table0 = stream
		case "1Table":
			table1 = stream
		}
	}
	return wordDoc, table0, table1
}

func getActiveTable(table0 *mscfb.File, table1 *mscfb.File, f *fib) *mscfb.File {
	if f.base.fWhichTblStm == 0 {
		return table0
	}
	return table1
}
