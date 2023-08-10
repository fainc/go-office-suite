package spreadsheet

import (
	"github.com/fainc/go-office-suite/spreadsheet/value"
)

type FieldWalker struct {
	Deep int
}

func FieldDeep(value []value.Field, w *FieldWalker, p int) {
	for _, keysValue := range value {
		if len(keysValue.Child) != 0 {
			FieldDeep(keysValue.Child, w, p+1)
		} else if p+1 > w.Deep {
			w.Deep = p + 1
		}
	}
}
