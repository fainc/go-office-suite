package spreadsheet

import (
	"fmt"

	"github.com/fainc/go-office-suite/spreadsheet/value"
)

// GetCoordinate 获取座标
func GetCoordinate(x, y int) string {
	return fmt.Sprintf("%v%v", GetXCoordinate(x), y)
}

// GetXCoordinate 获取列座标表
func GetXCoordinate(x int) string {
	tables := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	for i := 0; i < x/26; i++ {
		for n := 0; n < 26; n++ {
			tables = append(tables, tables[i]+tables[x])
		}
	}
	return tables[x]
}

// GetDescCoordinate 计算描述座标
func GetDescCoordinate(writeY, columns int, value []value.Field) (start, end string) {
	start = GetCoordinate(0, writeY)
	var endY int
	if columns > 1 {
		endY = writeY + columns - 1
	} else {
		endY = writeY
	}
	baseLength := len(value) - 1
	childLength := 0
	for _, key := range value {
		if len(key.Child) != 0 {
			childLength = childLength + len(key.Child) - 1
		}
	}
	end = GetCoordinate(baseLength+childLength, endY)
	return
}

// GetFieldCoordinate 计算键座标
func GetFieldCoordinate(writeX, writeY int, deep int, child []value.Field) (start, end string) {
	start = GetCoordinate(writeX, writeY)
	endX := writeX
	if len(child) > 0 {
		endX = writeX + len(child) - 1
	}
	endY := writeY
	if len(child) == 0 && deep != 1 {
		endY = writeY + deep - 1
	}
	end = GetCoordinate(endX, endY)
	return
}
