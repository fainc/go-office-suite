package spreadsheet

import (
	"errors"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"net/http"

	"github.com/xuri/excelize/v2"
)

func RenderImage(f *excelize.File, sheet, cell string, url interface{}) (err error) {
	str, ok := url.(string)
	if !ok {
		err = errors.New("image must be string")
		return
	}
	if str == "" {
		return
	}
	if isNetImage(str) {
		var file []byte
		file, err = getNetImage(str)
		if err != nil {
			return
		}
		if err = f.AddPictureFromBytes(sheet, cell, &excelize.Picture{
			Extension: ".jpg",
			File:      file,
			Format: &excelize.GraphicOptions{
				AutoFit:       true,
				Hyperlink:     str,
				HyperlinkType: "External",
				Positioning:   "twoCell",
			},
		}); err != nil {
			err = f.SetCellValue(sheet, cell, "image error:"+err.Error()+",src:"+str)
			return
		}
		return
	}
	if err = f.AddPicture(sheet, cell, str, &excelize.GraphicOptions{
		AutoFit:     true,
		Positioning: "twoCell",
	}); err != nil {
		err = f.SetCellValue(sheet, cell, "image error:"+err.Error()+",src:"+str)
		return
	}
	return
}

func isNetImage(url string) bool {
	if len(url) < 4 {
		return false
	}
	if url[0:4] == "http" {
		return true
	}
	return false
}

func getNetImage(url string) (file []byte, err error) {
	v, err := http.Get(url) //nolint:gosec
	if err != nil {
		return
	}
	defer v.Body.Close()
	file, err = io.ReadAll(v.Body)
	return
}
