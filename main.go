package main

import (
	"fmt"
	"github.com/lxn/walk"
	//. "github.com/lxn/walk/declarative"
	"github.com/xuri/excelize/v2"
	"log"
)

var mw *walk.MainWindow
var outTE *walk.TextEdit
var out string

func init() {
	mw, _ = walk.NewMainWindow()
	outTE, _ = walk.NewTextEdit(mw)
}
func main() {

	//Children := []Widget{
	//	HSplitter{
	//		Children: []Widget{},
	//	},
	//}
	mw.SetTitle("数据加工")
	mw.SetSize(walk.Size{300, 200})
	mw.SetLayout(walk.NewVBoxLayout())

	//hsplitter, _ := walk.NewHSplitter(mw)
	//hsplitter.Children([]Widget{
	//	TextEdit{AssignTo: &outTE, ReadOnly: true},
	//})
	button, _ := walk.NewPushButton(mw)
	button.SetText("Click me!")
	button.SetBounds(walk.Rectangle{20, 50, 80, 30})
	n := button.Clicked().Attach(loadBP)
	log.Println("button : n..", n)
	mw.Show()

	mw.Run()
}

func loadBP() {
	// 创建文件打开对话框
	dlg := new(walk.FileDialog)
	dlg.FilePath = ""
	dlg.Title = "选择要处理的excel文件"
	dlg.Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls"
	log.Println(dlg)
	if ok, err := dlg.ShowOpen(mw); err != nil {
		fmt.Println("文件选择对话框出错:", err.Error())
		return
	} else if !ok {
		return
	}

	// 读取 Excel 文件
	f, err := excelize.OpenFile(dlg.FilePath)
	if err != nil {
		fmt.Println("读取 Excel 文件出错:", err)
		return
	}
	defer f.Close()
	// 获取第一个工作表的名称
	sheetName := f.GetSheetName(0)
	// 获取工作表名称里的所有行
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println("获取行数据出错:", err)
		return
	}
	for _, row := range rows {
		for _, cell := range row {
			out += (cell + "\r\n")
		}
	}
	log.Println(len(out))
	outTE.SetText(out)

	log.Println(rows)
}
