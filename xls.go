package xls

import (
	"fmt"
	"io"
	"os"

	"github.com/morpon/extrame-xls/ole2"
)

const maxWorkbookStreamToFileRatio int64 = 4

// Open one xls file
func Open(file string, charset string) (*WorkBook, error) {
	if fi, err := os.Open(file); err == nil {
		return OpenReader(fi, charset)
	} else {
		return nil, err
	}
}

// Open one xls file and return the closer
func OpenWithCloser(file string, charset string) (*WorkBook, io.Closer, error) {
	if fi, err := os.Open(file); err == nil {
		wb, err := OpenReader(fi, charset)
		return wb, fi, err
	} else {
		return nil, nil, err
	}
}

// Open xls file from reader
func OpenReader(reader io.ReadSeeker, charset string) (wb *WorkBook, err error) {
	fileSize, err := seekReaderSize(reader)
	if err != nil {
		return nil, err
	}

	var ole *ole2.Ole
	if ole, err = ole2.Open(reader, charset); err == nil {
		var dir []*ole2.File
		if dir, err = ole.ListDir(); err == nil {
			var book *ole2.File
			var root *ole2.File
			for _, file := range dir {
				name := file.Name()
				if name == "Workbook" {
					if book == nil {
						book = file
					}
					//book = file
					// break
				}
				if name == "Book" {
					book = file
					// break
				}
				if name == "Root Entry" {
					root = file
				}
			}
			if book != nil {
				if root == nil {
					return nil, fmt.Errorf("invalid xls: root entry not found")
				}
				if err = validateWorkbookStreamSize(book.Size, fileSize); err != nil {
					return nil, err
				}
				wb, err = newWorkBookFromOle2(ole.OpenFile(book, root), fileSize)
				return wb, err
			}
			return nil, fmt.Errorf("invalid xls: workbook stream not found")
		}
	}
	return
}

func seekReaderSize(reader io.ReadSeeker) (int64, error) {
	current, err := reader.Seek(0, io.SeekCurrent)
	if err != nil {
		return 0, err
	}
	end, err := reader.Seek(0, io.SeekEnd)
	if err != nil {
		return 0, err
	}
	if _, err := reader.Seek(current, io.SeekStart); err != nil {
		return 0, err
	}
	if end <= 0 {
		return 0, fmt.Errorf("invalid file size: %d", end)
	}
	return end, nil
}

func validateWorkbookStreamSize(streamSize uint32, fileSize int64) error {
	declared := int64(streamSize)
	maxAllowed := fileSize * maxWorkbookStreamToFileRatio
	if declared > maxAllowed {
		return fmt.Errorf("invalid workbook stream size: %d exceeds %dx file size (%d)", declared, maxWorkbookStreamToFileRatio, fileSize)
	}
	return nil
}
