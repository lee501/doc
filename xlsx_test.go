package doc

import (
	"io"
	"os"
	"strings"
	"testing"
)

func TestParseXlsx(t *testing.T) {
	f, err := os.Open(`testData/simple.xlsx`)
	if err != nil {
		t.Fatal("expected to open test file", err)
	}
	defer f.Close()

	buf, err := ParseXlsx(f)
	if err != nil {
		t.Fatal("expected successful parse", err)
	}

	result_bytes, err := io.ReadAll(buf)
	if err != nil {
		t.Fatal("expected to read result", err)
	}
	result := string(result_bytes)
	if !strings.Contains(result, "Name") {
		t.Errorf("expected result to contain 'Name', got: %q", result)
	}
	if !strings.Contains(result, "Alice") {
		t.Errorf("expected result to contain 'Alice', got: %q", result)
	}
	if !strings.Contains(result, "Beijing") {
		t.Errorf("expected result to contain 'Beijing', got: %q", result)
	}
}

func TestParseXlsxInvalidInput(t *testing.T) {
	_, err := ParseXlsx(strings.NewReader("not an xlsx file"))
	if err == nil {
		t.Error("expected error for invalid input")
	}
}
