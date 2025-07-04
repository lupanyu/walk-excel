package main

import (
	"fmt"
	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
	"github.com/xuri/excelize/v2"
	"log"
	"strconv"
	"time"
)

var out = "提示信息...\r\n1.按按顺序输入sheet1-sheet5对应的费用\r\n2.点击选择体积excel文件\r\n3.点击选择数据excel文件\r\n4.点击生成excel文件\r\n会自动在当前目录下生成一个excel文件"
var result *walk.TextEdit
var mainWindow *walk.MainWindow
var inEdits [5]*walk.LineEdit

var VolumeData = map[string]float64{}

func main() {

	//var outTE *walk.TextEdit
	var buttonSize = Size{Width: 60, Height: 30}
	MainWindow{
		AssignTo: &mainWindow,
		Title:    "来抓娃娃啊，10块钱抓10个",
		Size:     Size{600, 400},
		Layout:   VBox{MarginsZero: true, Spacing: 10},
		Children: []Widget{
			Label{Text: "请输入5个sheet对应的费用",
				Font:      Font{PointSize: 14, Bold: true},
				Alignment: AlignHCenterVNear},
			Composite{
				Layout: HBox{MarginsZero: false, Spacing: 10},
				Children: []Widget{
					LineEdit{AssignTo: &inEdits[0], MinSize: Size{Width: 60}},
					LineEdit{AssignTo: &inEdits[1], MinSize: Size{Width: 60}},
					LineEdit{AssignTo: &inEdits[2], MinSize: Size{Width: 60}},
					LineEdit{AssignTo: &inEdits[3], MinSize: Size{Width: 60}},
					LineEdit{AssignTo: &inEdits[4], MinSize: Size{Width: 60}},
				},
			},
			TextEdit{
				AssignTo:  &result,
				ReadOnly:  true,
				Text:      out,
				Font:      Font{PointSize: 12},
				MinSize:   Size{Height: 150},
				Alignment: AlignHNearVNear,
			},

			PushButton{
				Text:      "选择体积excel文件",
				OnClicked: loadVolume,
				MaxSize:   buttonSize,
			},
			PushButton{
				Text:      "选择数据excel文件",
				OnClicked: loadBP,
				MaxSize:   buttonSize,
			},
			PushButton{
				Text:      "生成excel",
				OnClicked: loadBP,
				MaxSize:   buttonSize,
			},
		},
	}.Run()
}

// 加载体积数据
func loadVolume() {
	for i := 0; i < 5; i++ {
		log.Println(inEdits[i].Text())
	}
	// 创建文件打开对话框
	dlg := new(walk.FileDialog)
	dlg.FilePath = ""
	dlg.Title = "选择体积excel文件,只会加载第一个sheet中的数据"
	dlg.Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls"
	log.Println(dlg)
	if ok, err := dlg.ShowOpen(mainWindow); err != nil {
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
}

// 加载数据
func loadBP() {
	// 创建文件打开对话框
	dlg := new(walk.FileDialog)
	dlg.FilePath = ""
	dlg.Title = "选择要处理的excel文件"
	dlg.Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls"
	log.Println(dlg)
	if ok, err := dlg.ShowOpen(mainWindow); err != nil {
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
	out = fmt.Sprintf("加载文件%s成功\r\n开始统计\r\n", dlg.FilePath)

	// 获取所有表名
	sheetNames := f.GetSheetList()
	fmt.Println("工作表数量:", len(sheetNames))
	type sheetData struct {
		allVolume float64 // 求和算出总体积
		price     float64 // 计算单位体积的费用
		totalCost float64 // 输入总费用
		sheetName string
		rows      []struct {
			sku       string
			boxNum    int     // 箱数
			volume    float64 // 每箱体积
			sumVolume float64 // 总体积
			cost      float64 // 当前sku的总费用
		}
	}
	allSheetData := []sheetData{}
	// 遍历每个工作表
	for sheetIndex, sheetName := range sheetNames {
		fmt.Println("\n工作表名称:", sheetName)
		newsheetData := sheetData{
			sheetName: sheetName,
		}
		newsheetData.allVolume = 0
		// 获取工作表名称里的所有行
		rows, err := f.GetRows(sheetName)
		if err != nil {
			fmt.Println("获取行数据出错:", err)
			return
		}
		log.Println("------\n", rows)
		for _, row := range rows[5:] {
			if row[4] == "" {
				continue
			}
			boxNum, _ := strconv.ParseFloat(row[7], 64)
			newsheetData.rows = append(newsheetData.rows, struct {
				sku       string
				boxNum    int
				volume    float64
				sumVolume float64
				cost      float64
			}{
				sku:       row[4],
				boxNum:    int(boxNum),
				volume:    VolumeData[row[4]],
				sumVolume: VolumeData[row[4]] * boxNum,
			})
			// log.Printf("sku: %s ,箱数：%s,总数量: %s,单位体积:%f，总体积: %f", row[4], row[7], row[9], VolumeData[row[4]], VolumeData[row[4]]*boxNum)
			newsheetData.allVolume += VolumeData[row[4]] * boxNum
		}
		logData := fmt.Sprintf("统计出sheet %s: 总方数是:%f\r\n", sheetName, newsheetData.allVolume)
		out += logData
		log.Println(logData)
		var price float64 // 单位体积的费用
		// 根据输入的数字计算出单位体积的费用
		if inEdits[sheetIndex].Text() != "" {
			allPrice, _ := strconv.ParseFloat(inEdits[0].Text(), 64)
			log.Println("allPrice", allPrice, "输入的数据是：", inEdits[0].Text())
			price = allPrice / newsheetData.allVolume
			newsheetData.price = price
			logData = fmt.Sprintf("sheet %s: 单位体积费用是:%f\r\n", sheetName, newsheetData.price)
			out += logData
			log.Println(logData)
		}
		// 计算出每个sku花费
		for n, row := range newsheetData.rows {
			row.cost = row.volume * newsheetData.price * float64(row.boxNum)
			log.Printf("sku: %s,箱数：%d,总数量: %f,单位体积:%f，总体积: %f,总费用: %f\r\n", row.sku, row.boxNum, row.sumVolume, row.volume, row.sumVolume, row.cost)
			newsheetData.rows[n] = row
		}
		result.SetText(out)
		allSheetData = append(allSheetData, newsheetData)
	}
	log.Println("-----------------allSheetData", allSheetData)
	// 创建一个新的 Excel 文件
	newFile := excelize.NewFile()
	defer newFile.Close()
	for _, newFileData := range allSheetData {
		// 创建一个新的工作表
		newFile.NewSheet(newFileData.sheetName)
		newSheetName := newFileData.sheetName
		newFile.SetCellValue(newSheetName, "A1", "sku")
		newFile.SetCellValue(newSheetName, "B1", "箱数")
		newFile.SetCellValue(newSheetName, "C1", "单位体积")
		newFile.SetCellValue(newSheetName, "D1", "总体积")
		newFile.SetCellValue(newSheetName, "E1", "总费用")
		for i, row := range newFileData.rows {
			newFile.SetCellValue(newSheetName, fmt.Sprintf("A%d", i+2), row.sku)
			newFile.SetCellValue(newSheetName, fmt.Sprintf("B%d", i+2), row.boxNum)
			newFile.SetCellValue(newSheetName, fmt.Sprintf("C%d", i+2), row.volume)
			newFile.SetCellValue(newSheetName, fmt.Sprintf("D%d", i+2), row.sumVolume)
			newFile.SetCellValue(newSheetName, fmt.Sprintf("E%d", i+2), row.cost)
		}
	}
	newFile.SetActiveSheet(0)
	fileName := time.Now().Format("20060102-150405") + ".xlsx"
	if err := newFile.SaveAs(fileName); err != nil {
		log.Println(err)
	}
	out += fmt.Sprintf("生成文件%s成功\r\n", fileName)
	result.SetText(out)
	log.Println(allSheetData)
}
