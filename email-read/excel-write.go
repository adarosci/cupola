package main

import (
	"fmt"
	"os"

	"github.com/tealeg/xlsx"
)

// ExcelFile struct file
type ExcelFile struct {
	File  *xlsx.File
	Sheet map[string]*xlsx.Sheet
}

// CreateFile excel
func CreateFile() *ExcelFile {
	file := xlsx.NewFile()
	return &ExcelFile{
		File:  file,
		Sheet: make(map[string]*xlsx.Sheet),
	}
}

// AddRow nova linha excel
func (f *ExcelFile) AddRow(sheet string, itens []string) {
	s := f.Sheet[sheet]
	if s == nil {
		s, _ = f.File.AddSheet(sheet)
		f.Sheet[sheet] = s
	}
	row := s.AddRow()
	for _, item := range itens {
		cell := row.AddCell()
		cell.Value = item
	}
}

// Save salvar excel
func (f *ExcelFile) Save(filename string) error {
	if exists := f.fileExists(filename); exists {
		os.Remove(filename)
	}
	err := f.File.Save(filename)
	if err != nil {
		fmt.Printf(err.Error())
	}
	return err
}

func (f *ExcelFile) fileExists(filename string) bool {
	info, err := os.Stat(filename)
	if os.IsNotExist(err) {
		return false
	}
	return !info.IsDir()
}
