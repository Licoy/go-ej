package main

import (
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/AlecAivazis/survey/v2"
	"github.com/gookit/color"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
)

type EjAnswers struct {
	UsePwdPath    string
	CustomPath    string
	UseOutPwdPath string
	OutPath       string
}

var errCount = 0

var banner = `
   ____             _____    _  
  / ___| ___       | ____|  | | 
 | |  _ / _ \ _____|  _| _  | | 
 | |_| | (_) |_____| |__| |_| | 
  \____|\___/      |_____\___/  
                                

`

func main() {

	args := os.Args

	if len(args) == 3 {
		//命令行参数执行
		argInPath := args[1]
		argOutPath := args[2]
		if !IsDir(argInPath) {
			color.Red.Printf("错误：[ %s ]不是一个有效的执行目录\n", argInPath)
			return
		}
		if !IsDir(argOutPath) {
			color.Red.Printf("错误：[ %s ]不是一个有效的输出目录\n", argOutPath)
			return
		}
		start(argInPath, argOutPath)

	} else {
		//选择执行
		color.BgCyan.Println(banner)
		pwdPath, err := os.Getwd()
		if err != nil {
			color.Red.Printf("获取当前运行目录失败：%v\n", err)
			return
		}

		answers := &EjAnswers{}

		outPath := pwdPath + "/out-json"

		configInput(&answers.UsePwdPath, &answers.CustomPath, pwdPath, fmt.Sprintf("是否对[ %s ]目录下的Excel文件进行转换处理？", pwdPath), "请输入完整的处理目录：")

		configInput(&answers.UseOutPwdPath, &answers.OutPath, outPath, fmt.Sprintf("是否使用[ %s ]作为输出目录", outPath), "请输入完整的输出目录：")

		fmt.Println(answers)

		start(answers.CustomPath, answers.OutPath)
	}
}

func configInput(resp interface{}, finalTo interface{}, pwdPath string, msg string, outMsg string) {

	sErr := survey.AskOne(&survey.Select{
		Message: msg,
		Options: []string{"是", "否"},
		Default: "是",
	}, resp)
	if sErr != nil {
		color.Red.Println(sErr.Error())
		return
	}

	rp := resp.(*string)
	if *rp == "否" {
		whileInput(finalTo, false, outMsg)
	} else {
		ft := finalTo.(*string)
		*ft = pwdPath
	}

}

func whileInput(resp interface{}, notDir bool, outMsg string) {
	msg := outMsg
	if notDir {
		msg = "[错误/无效路径]" + outMsg
	}
	sErr := survey.AskOne(&survey.Input{
		Message: msg,
	}, resp)

	if sErr != nil {
		color.Red.Println(sErr.Error())
		return
	}

	if !IsDir(*resp.(*string)) {
		color.Red.Printf("错误：[ %s ]不是一个有效的目录", resp)
		whileInput(resp, true, outMsg)
	}
}

func start(filepath string, outPath string) {
	files, err := getAllExcel(filepath)
	if err != nil {
		color.Red.Printf("读取目录出现错误：%v\n", err)
	}

	for _, file := range files {
		fileCurPath := strings.Replace(file, filepath, "", 1)
		color.Blue.Printf("开始处理：%s\n", fileCurPath)
		readExcel(filepath, file, outPath)
		color.Green.Printf("处理完成：%s\n--------------------------\n", fileCurPath)
	}

	finalMsg := fmt.Sprintf("Excel文件都已转换完成，共计包含%d个失败处理", errCount)

	if errCount != 0 {
		color.Danger.Println(finalMsg + "，请查看控制台日志记录")
	} else {
		color.Primary.Println(finalMsg)
	}
	var input string
	color.Primary.Println("请键入任意键回车进行退出...")
	_, _ = fmt.Scanln(&input)
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
func readExcel(basePath string, file string, outPath string) {
	outFile := strings.Replace(file, basePath, outPath, 1)
	outFile = strings.Replace(outFile, ".xlsx", ".json", 1)
	var readErr error
	var wf *os.File
	outPaths, _ := filepath.Split(outFile)
	if checkFileIsExist(outFile) { //如果文件存在
		_ = os.Remove(outFile)
		wf, readErr = os.Create(outFile) //创建文件
	} else {
		_ = os.MkdirAll(outPaths, os.ModePerm)
		wf, readErr = os.Create(outFile) //创建文件
	}
	if readErr != nil {
		errCount++
		color.Red.Printf("创建%s文件的写入流失败 %v\n", outFile, readErr)
		return
	}
	defer wf.Close()
	f, err := excelize.OpenFile(file)
	if err != nil {
		errCount++
		color.Red.Printf("读取Excel失败：%s, %v\n", file, err)
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
	_, writeErr := io.WriteString(wf, string(marshal))
	if writeErr != nil {
		errCount++
		color.Red.Printf("写入文件失败失败：%s, %v\n", outFile, writeErr)
	}
}
