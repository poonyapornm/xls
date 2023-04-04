package xls

import (
	"fmt"
	"io"
	"os"

	"github.com/extrame/ole2"
)

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
func OpenReader(reader io.ReadSeeker, charset string) (*WorkBook, error) {
	ole, err := ole2.Open(reader, charset)
	if err != nil {
		return nil, err
	}

	dir, err := ole.ListDir()
	if err != nil {
		return nil, err
	}

	var book, root *ole2.File
	for _, file := range dir {
		name := file.Name()
		switch name {
		case "Workbook":
			if book == nil {
				book = file
			}
		case "Book":
			book = file
		case "Root Entry":
			root = file
		}
	}

	if book == nil || root == nil {
		return nil, fmt.Errorf("book=%v, root=%v", book, root)
	}

	return newWorkBookFromOle2(ole.OpenFile(book, root)), nil
}
