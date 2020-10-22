package main

import (
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"

	"github.com/gookit/color"
)

func main() {

	filepath, err := os.Getwd()
	if err != nil {
		color.Red.Printf("获取当前运行目录失败：%v\n", err)
		return
	}

	color.Green.Println("----------------------------------")
	color.Green.Printf("是否以 ")
	color.BgBlue.Printf("[ %v ]", filepath)
	color.Green.Printf(" 作为运行为目录\n1. 若是请直接回车\n2. 反之则输入对应目录的全路径\n")
	color.Green.Println("----------------------------------")

	var input string
	fmt.Scanln(&input)
	fmt.Println("")
	if input == "" {
		start(filepath)
	} else {
		if !IsDir(input) {
			color.Red.Printf("[ %v ]不是一个有效的目录，请重新输入：\n", input)
			whileInput()
		} else {
			start(input)
		}
	}

	//
}

func whileInput() {
	var input string
	fmt.Scanln(&input)
	if !IsDir(input) {
		color.Red.Printf("[ %v ]不是一个有效的目录，请重新输入：\n", input)
		whileInput()
	} else {
		start(input)
	}
}

var errCount int = 0

func start(filepath string) {
	files, err := getAllExcel(filepath)
	if err != nil {
		color.Red.Printf("读取目录出现错误：%v\n", err)
	}

	for _, file := range files {
		color.Blue.Printf("开始处理：%s\n", file)
		readExcel(filepath, file)
		color.Green.Printf("处理完成：%s\n--------------------------\n", file)
	}

	finalMsg := fmt.Sprintf("Excel文件都已转换完成，共计包含%d个失败处理", errCount)

	if errCount != 0 {
		color.Danger.Println(finalMsg + "，请查看控制台日志记录")
	} else {
		color.Primary.Println(finalMsg)
	}
	var input string
	color.Primary.Println("请键入任意键回车进行退出...")
	fmt.Scanln(&input)
}

func IsDir(path string) bool {
	s, err := os.Stat(path)
	if err != nil {
		return false
	}
	return s.IsDir()
}

func getAllExcel(path string) ([]string, error) {
	files, err := ioutil.ReadDir(path)
	if err != nil {
		return nil, err
	}
	res := make([]string, 0, 100)
	for _, filename := range files {
		allPath := path + "/" + filename.Name()
		if IsDir(allPath) {
			nextFiles, _ := getAllExcel(allPath)
			res = append(res, nextFiles...)
		} else {
			if strings.HasSuffix(filename.Name(), ".xlsx") {
				res = append(res, allPath)
			}
		}
	}
	return res, nil
}

func checkFileIsExist(filename string) bool {
	var exist = true
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		exist = false
	}
	return exist
}

//读取excel
func readExcel(basePath string, file string) {
	outFile := strings.Replace(file, basePath, basePath+"/out-json", 1)
	outFile = strings.Replace(outFile, ".xlsx", ".json", 1)
	var rerr error
	var wf *os.File
	outPaths, _ := filepath.Split(outFile)
	if checkFileIsExist(outFile) { //如果文件存在
		os.Remove(outFile)
		wf, rerr = os.Create(outFile) //创建文件
	} else {
		os.MkdirAll(outPaths, os.ModePerm)
		wf, rerr = os.Create(outFile) //创建文件
	}
	if rerr != nil {
		errCount++
		color.Red.Printf("创建%s文件的写入流失败 %v\n", outFile, rerr)
		return
	}
	defer wf.Close()
	f, err := excelize.OpenFile(file)
	if err != nil {
		errCount++
		color.Red.Printf("读取Execl失败：%s, %v\n", file, err)
		return
	}
	firstSheet := f.GetSheetList()[0]
	rows, _ := f.GetRows(firstSheet)
	dataDict := make([]interface{}, 0, 2000)
	keys := make([]string, 0, 50)
	for i, row := range rows {
		if i == 0 {
			for _, colCell := range row {
				keys = append(keys, colCell)
			}
		}
		if i < 3 || len(row) == 0 {
			continue
		}
		cells := make(map[string]string)
		for k, colCell := range row {
			if k >= len(keys) {
				break
			}
			cells[keys[k]] = colCell
		}
		//检测字段是否全部为空
		isAppend := false
		for _, v := range cells {
			if v != "" {
				isAppend = true
				break
			}
		}
		if isAppend {
			dataDict = append(dataDict, cells)
		}
	}
	marshal, err := json.MarshalIndent(dataDict, "", "    ")
	if err != nil {
		errCount++
		color.Red.Printf("转换JSON失败：%s, %v\n", file, err)
		return
	}
	_, werr := io.WriteString(wf, string(marshal))
	if werr != nil {
		errCount++
		color.Red.Printf("写入文件失败失败：%s, %v\n", outFile, werr)
	}
}
