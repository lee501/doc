package doc

import (
	"io"
	"os"
	"strings"
	"testing"
)

func TestParseDocx(t *testing.T) {
	f, err := os.Open(`testData/simple.docx`)
	if err != nil {
		t.Fatal("expected to open test file", err)
	}
	defer f.Close()

	buf, err := ParseDocx(f)
	if err != nil {
		t.Fatal("expected successful parse", err)
	}

	result_bytes, err := io.ReadAll(buf)
	if err != nil {
		t.Fatal("expected to read result", err)
	}
	result := string(result_bytes)
	if !strings.Contains(result, "test content") {
		t.Errorf("expected result to contain 'test content', got: %q", result)
	}
	if !strings.Contains(result, "Hello World") {
		t.Errorf("expected result to contain 'Hello World', got: %q", result)
	}
}

func TestParseDocxInvalidInput(t *testing.T) {
	_, err := ParseDocx(strings.NewReader("not a docx file"))
	if err == nil {
		t.Error("expected error for invalid input")
	}
}
