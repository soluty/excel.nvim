package main

import (
	"errors"
	"fmt"

	"github.com/neovim/go-client/nvim"
	"github.com/neovim/go-client/nvim/plugin"
	"github.com/tealeg/xlsx/v3"
)

type Tables struct {
	Tables []*Table
}

type Table struct {
	Name   string
	Cells  [][]string
	MaxCol int
}

func print(v *nvim.Nvim, str ...string) {
	pstr := fmt.Sprintf("%v", str)
	v.Echo([]nvim.TextChunk{{Text: pstr}}, true, nil)
}

func openExcel(vim *nvim.Nvim, args []string) (*Tables, error) {
	if len(args) == 0 {
		return nil, errors.New("insert excel file name")
	}
	filename := args[0]
	wb, err := xlsx.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("cant open excel: %w", err)
	}
	print(vim, "打开文件", args[0])
	tbs := &Tables{}
	for _, sheet := range wb.Sheets {
		tb := &Table{
			Name:   sheet.Name,
			MaxCol: sheet.MaxCol,
		}
		for i := 0; i < sheet.MaxRow; i++ {
			tb.Cells = append(tb.Cells, nil)
			for j := 0; j < sheet.MaxCol; j++ {
				cell, err := sheet.Cell(i, j)
				if err != nil {
					return nil, fmt.Errorf("获取数据错误: %w", err)
				}
				tb.Cells[i] = append(tb.Cells[i], cell.Value)
			}
		}
		tbs.Tables = append(tbs.Tables, tb)
	}
	return tbs, nil
}

func main() {
	plugin.Main(func(p *plugin.Plugin) error {
		p.HandleFunction(&plugin.FunctionOptions{Name: "ExcelNvimOpenFile"}, openExcel)
		return nil
	})
}
