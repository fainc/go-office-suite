package spreadsheet

import (
	"errors"
	"fmt"
	"io"
	"reflect"

	"github.com/xuri/excelize/v2"

	"github.com/fainc/go-office-suite/spreadsheet/style"
	"github.com/fainc/go-office-suite/spreadsheet/value"
)

type mapExport struct {
	Sheet []*value.Sheet
}

func MapExport(sheet []*value.Sheet) *mapExport {
	return &mapExport{Sheet: sheet}
}

func (rec *mapExport) SaveFile(savePath string, activeSheet int) (err error) {
	f := excelize.NewFile()
	defer func() {
		_ = f.Close()
	}()
	err = rec.mapWriter(f)
	if err != nil {
		return err
	}
	f.SetActiveSheet(activeSheet)
	if err = f.SaveAs(savePath); err != nil {
		return
	}
	return
}

func (rec *mapExport) WriteIO(w io.Writer, activeSheet int) (err error) {
	f := excelize.NewFile()
	defer func() {
		_ = f.Close()
	}()
	err = rec.mapWriter(f)
	if err != nil {
		return err
	}
	f.SetActiveSheet(activeSheet)
	if err = f.Write(w); err != nil {
		return
	}
	return
}

func (rec *mapExport) mapWriter(f *excelize.File) (err error) {
	sheets := rec.Sheet
	for i, sheet := range sheets {
		if i == 0 {
			if sheet.SheetName != "" { // 重命名Sheet1
				err = f.SetSheetName("Sheet1", sheet.SheetName)
				if err != nil {
					return err
				}
			} else {
				sheet.SheetName = "Sheet1"
			}
		}
		if i != 0 { // 新增表
			_, err = f.NewSheet(sheet.SheetName)
		}
	}
	for _, sheet := range sheets {
		writeX := 0 // 列起始位
		writeY := 1 // 行起始位
		// 头部描述
		for _, title := range sheet.Desc {
			titleStart, titleEnd := GetDescCoordinate(writeY, title.Column, sheet.Field)
			err = f.SetCellValue(sheet.SheetName, titleStart, title.Text)
			if err != nil {
				return
			}
			// 合并标题单元格
			if titleStart != titleEnd {
				_ = f.MergeCell(sheet.SheetName, titleStart, titleEnd)
			}
			// 定义样式
			s, _ := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: title.Align, Vertical: "center", WrapText: true}, Font: &excelize.Font{Size: title.FontSize}})
			if err != nil {
				return
			}
			_ = f.SetCellStyle(sheet.SheetName, titleStart, titleStart, s)
			// 更新行号
			if title.Column <= 1 { // 默认占一行
				writeY++
			} else {
				writeY += title.Column
			}
		}
		// 数据键写入
		writeX = 0
		w := &FieldWalker{Deep: 0}
		FieldDeep(sheet.Field, w, 0) // 计算键深
		if w.Deep >= 3 {
			err = errors.New("不支持生成导出超过二级key的数据")
			break
		}
		for _, key := range sheet.Field {
			keyStart, keyEnd := GetFieldCoordinate(writeX, writeY, w.Deep, key.Child)
			err = f.SetCellValue(sheet.SheetName, keyStart, key.Name)
			if err != nil {
				return
			}
			style.AlignCenter(f, sheet.SheetName, keyStart, keyStart)
			if keyStart != keyEnd {
				_ = f.MergeCell(sheet.SheetName, keyStart, keyEnd)
			}
			if len(key.Child) != 0 {
				writeY++
				for c, child := range key.Child {
					childStart, _ := GetFieldCoordinate(writeX+c, writeY, 1, key.Child)
					err = f.SetCellValue(sheet.SheetName, childStart, child.Name)
					if err != nil {
						return
					}
					style.AlignCenter(f, sheet.SheetName, childStart, childStart)
				}
				writeY--
			}
			if len(key.Child) > 0 {
				writeX += len(key.Child)
			} else {
				writeX++
			}
		}
		writeX = 0
		writeY += w.Deep
		// 数据行写入
		for _, data := range sheet.Rows { // 行
			jumpY := 0
			var mergeCell [][]string
			for _, key := range sheet.Field { // 列
				if len(key.Child) == 0 { // 非子键关联数据
					// 数据类型判断
					if d, ok := data[key.Index]; !ok || d == nil {
						cell := GetCoordinate(writeX, writeY)
						mergeCell = append(mergeCell, []string{cell, GetXCoordinate(writeX)})
						writeX++ // 空数据忽略
						continue
					}
					t := reflect.TypeOf(data[key.Index])
					if t == nil { // 反射错误
						err = errors.New("reflect.TypeOf " + key.Index + " 发生错误")
						break
					}
					if t.Kind() != reflect.Slice { // 常规数据写入
						cell := GetCoordinate(writeX, writeY)
						if key.RenderImage {
							err = RenderImage(f, sheet.SheetName, cell, data[key.Index])
							if err != nil {
								break
							}
						} else {
							err = f.SetCellValue(sheet.SheetName, cell, data[key.Index])
							if err != nil {
								break
							}
						}
						mergeCell = append(mergeCell, []string{cell, GetXCoordinate(writeX)})
					}
					if t.Kind() == reflect.Slice { // 切片写入
						mSlice, ok := data[key.Index].([]interface{})
						if !ok {
							err = errors.New(key.Index + " must be []interface{}")
							break
						}
						for i, slice := range mSlice {
							cell := GetCoordinate(writeX, writeY+i)
							if key.RenderImage {
								err = RenderImage(f, sheet.SheetName, cell, slice)
								if err != nil {
									break
								}
							} else {
								err = f.SetCellValue(sheet.SheetName, cell, slice)
								if err != nil {
									break
								}
							}
						}
						if len(mSlice) >= jumpY {
							jumpY = len(mSlice)
						}
					}
					writeX++
				}
				if err != nil {
					break
				}
				if len(key.Child) != 0 { // 子键关联数据
					if d, ok := data[key.Index]; !ok || d == nil { // 无数据 忽略
						writeX += len(key.Child)
						continue
					}
					mMap, ok := data[key.Index].([]map[string]interface{})
					if !ok {
						err = errors.New(key.Index + " must be []map[string]interface{}")
						break
					}
					for _, childData := range mMap {
						for c, ckey := range key.Child {
							childCell := GetCoordinate(writeX+c, writeY)
							if ckey.RenderImage {
								err = RenderImage(f, sheet.SheetName, childCell, childData[ckey.Index])
								if err != nil {
									break
								}
							} else {
								err = f.SetCellValue(sheet.SheetName, childCell, childData[ckey.Index])
								if err != nil {
									break
								}
							}
						}
						writeY++
					}
					writeY -= len(mMap)
					writeX += len(key.Child)
					if len(mMap) >= jumpY {
						jumpY = len(mMap)
					}
				}
			}
			if err != nil {
				break
			}
			if jumpY > 0 {
				writeY += jumpY
				if len(mergeCell) > 0 { // 处理主数据合并
					for _, ms := range mergeCell {
						msEnd := fmt.Sprintf("%v%v", ms[1], writeY-1)
						if ms[0] != msEnd {
							_ = f.MergeCell(sheet.SheetName, ms[0], msEnd)
							style.AlignCenter(f, sheet.SheetName, ms[0], ms[0])
						}
					}
				}
			} else {
				writeY++
			}
			writeX = 0
		}
	}
	return
}
