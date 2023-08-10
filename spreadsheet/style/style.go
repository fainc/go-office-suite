package style

import (
	"github.com/xuri/excelize/v2"
)

func AlignCenter(f *excelize.File, sheet, start, end string) {
	style, _ := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"}})
	_ = f.SetCellStyle(sheet, start, end, style)
}
