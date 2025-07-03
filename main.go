package main

import (
	"fmt"
	"github.com/lxn/walk"
	"strconv"

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

var VolumeData = map[string]float64{}

func main() {

	//Children := []Widget{
	//	HSplitter{
	//		Children: []Widget{},
	//	},
	//}

	mw.SetTitle("来抓娃娃啊，10块钱抓10个")
	mw.SetSize(walk.Size{600, 400})
	mw.SetLayout(walk.NewVBoxLayout())

	buttonVolume, _ := walk.NewPushButton(mw)
	buttonVolume.SetText("选择体积excel文件")
	buttonVolume.SetBounds(walk.Rectangle{20, 50, 40, 30})
	n := buttonVolume.Clicked().Attach(loadVolume)
	log.Println("button : n..", n)

	//hsplitter, _ := walk.NewHSplitter(mw)
	//hsplitter.Children([]Widget{
	//	TextEdit{AssignTo: &outTE, ReadOnly: true},
	//})
	buttonData, _ := walk.NewPushButton(mw)
	buttonData.SetText("选择数据excel文件")
	buttonData.SetBounds(walk.Rectangle{20, 50, 40, 30})
	l := buttonData.Clicked().Attach(loadBP)
	log.Println("button : l..", l)
	mw.Show()

	mw.Run()
}

func loadVolume() {
	// 创建文件打开对话框
	dlg := new(walk.FileDialog)
	dlg.FilePath = ""
	dlg.Title = "选择体积excel文件,只会加载第一个sheet中的数据"
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
	for line, row := range rows {
		if line == 0 {
			continue
		} else {
			log.Println(line, row)
			if len(row) == 2 {
				VolumeData[row[0]], _ = strconv.ParseFloat(row[1], 64)
			}
		}

	}
	log.Println("volumeData", VolumeData)
	outTE.SetText("加载体积数据成功!")
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
