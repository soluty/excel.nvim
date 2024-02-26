package main

import (
	"fmt"

	"github.com/neovim/go-client/nvim"
	"github.com/neovim/go-client/nvim/plugin"
)

type MyStruct struct {
	Name  string
	Value string
}

func openExcel(v *nvim.Nvim, args []string) (*MyStruct, error) {
	fmt.Println("pfrom golang")
	return &MyStruct{
		Name:  "name",
		Value: "value",
	}, nil
	// if len(args) == 0 {
	// 	return "", errors.New("insert excel file name")
	// }
	// filename := args[0]
	// wb, err := xlsx.OpenFile(filename)
	// if err != nil {
	// 	return "", fmt.Errorf("cant open excel: %w", err)
	// }
	// if len(wb.Sheets) == 1 {
	// 	// open direct新开一个buffer, 在新的buffer里面打开excel文件.
	// 	sheet := wb.Sheets[0]
	// 	return sheet.Name, nil
	// }
	// // wb.Sheets
	// return "Hello " + strings.Join(args, " "), nil
}

func main() {
	plugin.Main(func(p *plugin.Plugin) error {
		p.HandleFunction(&plugin.FunctionOptions{Name: "ExcelNvimOpenFile"}, openExcel)
		return nil
	})
}
