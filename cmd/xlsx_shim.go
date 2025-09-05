package main

import (
	excelize "github.com/xuri/excelize/v2"
)

type excelizeFile struct{ *excelize.File }

func (f *excelizeFile) NewSheet(name string) int { idx, _ := f.File.NewSheet(name); return idx }
func (f *excelizeFile) SetCellValue(sheet, axis string, value interface{}) error {
	return f.File.SetCellValue(sheet, axis, value)
}
func (f *excelizeFile) SetActiveSheet(index int) { f.File.SetActiveSheet(index) }
func (f *excelizeFile) DeleteSheet(name string)  { f.File.DeleteSheet(name) }
func (f *excelizeFile) SaveAs(name string) error { return f.File.SaveAs(name) }
func (f *excelizeFile) AddDataValidation(sheet string, dv DataValidation) error {
	if edv, ok := dv.(*excelizeDataValidation); ok {
		return f.File.AddDataValidation(sheet, edv.dv)
	}
	return nil
}

func excelizeNew() excelFile { return &excelizeFile{excelize.NewFile()} }

// Data validation wrapper
type excelizeDataValidation struct {
	dv *excelize.DataValidation
}

func (dv *excelizeDataValidation) SetRange(rangeAddr string) error {
	dv.dv.Sqref = rangeAddr
	return nil
}

func (dv *excelizeDataValidation) SetDropList(options []string) error {
	return dv.dv.SetDropList(options)
}

func newDataValidation() DataValidation {
	dv := excelize.NewDataValidation(true)
	dv.ShowDropDown = true // 启用下拉箭头显示
	return &excelizeDataValidation{
		dv: dv,
	}
}
