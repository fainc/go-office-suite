package test

import (
	"fmt"
	"testing"
	"time"

	"github.com/fainc/go-office-suite/spreadsheet"
	"github.com/fainc/go-office-suite/spreadsheet/value"
)

func makeSpecs(n int) []map[string]interface{} {
	specs := make([]map[string]interface{}, 0)
	for i := 1; i <= n; i++ {
		var spec = make(map[string]interface{}, 0)
		spec["specID"] = fmt.Sprintf("%v%v", n, i)
		spec["specName"] = fmt.Sprintf("%v%v", i*32, "G")
		// spec["image"] = "https://www.apple.com.cn/v/iphone/home/bo/images/overview/compare/icon_a15__gde9u4vunqqa_large_2x.png"
		specs = append(specs, spec)
	}
	return specs
}
func TestSpreadSheetExport(t *testing.T) {
	// Make Rows
	var dataRows []map[string]interface{}
	var dataRow = make(map[string]interface{}, 1)

	// support slice
	var slices []interface{}
	slices = append(slices, "Slice 1")
	slices = append(slices, "Slice 2")

	dataRow["id"] = "1"
	dataRow["name"] = "Apple Mac"
	dataRow["specs"] = makeSpecs(5) // child rows
	dataRow["date"] = "2023-01-01"
	dataRow["slices"] = slices

	// support net image,but need filed RenderImage is true
	// dataRow["image"] = "https://www.apple.com.cn/iphone/home/images/overview/hero/hero_iphone_14__de41900yuggi_medium_2x.jpg"
	dataRows = append(dataRows, dataRow)
	for i := 0; i < 20000; i++ {
		dataRows = append(dataRows, dataRow) //  repeat row
	}

	var sheets []*value.Sheet
	// Index accout rows key,child
	fields := []value.Field{{Name: "ID", Index: "id"}, {Name: "Name", Index: "name"}, {Name: "Specs", Index: "specs", Child: []value.Field{{Name: "ID", Index: "specID"}, {Name: "Name", Index: "specName"}, {Name: "Chip", Index: "image", RenderImage: true}}}, {Name: "Slice", Index: "slices"}, {Name: "Net Image", Index: "image", RenderImage: true}}
	// 标题
	var desc []value.Desc
	desc = append(desc, value.Desc{Text: "Test Spreadsheet", Column: 4, Align: "center", FontSize: 24})
	desc = append(desc, value.Desc{Text: time.Now().String(), Column: 2, Align: "center", FontSize: 10})
	sheets = append(sheets, &value.Sheet{SheetName: "Test", Desc: desc, Field: fields, Rows: dataRows})
	err := spreadsheet.MapExport().SaveFile("test_specs.xlsx", 0, sheets)
	if err != nil {
		fmt.Println(err)
	}
}
func TestSpreadSheetExportMini(t *testing.T) {
	// 表Test数据
	var dataRows []map[string]interface{}
	var dataRow = make(map[string]interface{}, 1)
	dataRow["id"] = "1"
	// dataRow["specs"] = makeSpecs(5)
	dataRow["image"] = "https://www.apple.com.cn/iphone/home/images/overview/hero/hero_iphone_14__de41900yuggi_medium_2x.jpg"
	dataRows = append(dataRows, dataRow)
	// Excel表数据
	var sheets []*value.Sheet
	// 键
	keys := []value.Field{{Name: "ID", Index: "id"}, {Name: "Name", Index: "name"}, {Name: "Slice", Index: "coupon", RenderImage: false}, {Name: "Amount", Index: "amount"}, {Name: "Nil", Index: "id2"}, {Name: "Net Image", Index: "image", RenderImage: true}, {Name: "Local Image", Index: "local_image", RenderImage: true}, {Name: "Image Failed", Index: "failed_image", RenderImage: true}}
	// 标题
	var title []value.Desc
	title = append(title, value.Desc{Text: "Test Spreadsheet", Column: 4, Align: "center", FontSize: 24})
	title = append(title, value.Desc{Text: time.Now().String(), Column: 2, Align: "center", FontSize: 10})
	sheets = append(sheets, &value.Sheet{SheetName: "Test", Desc: title, Field: keys, Rows: dataRows})
	err := spreadsheet.MapExport().SaveFile("test_specs.xlsx", 0, sheets)
	if err != nil {
		fmt.Println(err)
	}
}
