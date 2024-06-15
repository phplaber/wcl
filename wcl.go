package main

import (
	"bytes"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func exists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}

func lineCounter(r io.Reader) (int, error) {
	buf := make([]byte, 32*1024)
	count := 0
	lineSep := []byte{'\n'}

	for {
		c, err := r.Read(buf)
		count += bytes.Count(buf[:c], lineSep)

		switch {
		case err == io.EOF:
			return count, nil

		case err != nil:
			return count, err
		}
	}
}

func main() {
	// 传入文件或目录路径
	file := os.Args[1]
	if ok, _ := exists(file); !ok {
		log.Fatalf("文件或目录路径不存在或无效路径: %s", file)
	}

	files := []string{}
	fi, _ := os.Stat(file)
	switch mode := fi.Mode(); {
	case mode.IsDir():
		entries, err := os.ReadDir(file)
		if err != nil {
			log.Fatalf("无法打开目录: %v", err)
		}

		for _, e := range entries {
			if e.Type().IsRegular() && !strings.HasPrefix(e.Name(), ".") {
				files = append(files, filepath.Join(file, e.Name()))
			}
		}
	case mode.IsRegular():
		files = append(files, file)
	}

	var totalCount int
	for _, f := range files {
		var rowCount int
		ext := strings.ToLower(filepath.Ext(f))
		if ext != ".xlsx" {
			ff, _ := os.Open(f)
			defer ff.Close()
			rowCount, _ = lineCounter(ff)
		} else {
			ff, err := excelize.OpenFile(f)
			if err != nil {
				log.Fatalf("无法打开文件: %v", err)
			}

			sheetNames := ff.GetSheetList()
			for _, sheetName := range sheetNames {
				rows, err := ff.GetRows(sheetName)
				if err != nil {
					log.Fatalf("无法获取工作表 %s 行: %v", sheetName, err)
				}
				rowCount += len(rows)
			}
		}
		totalCount += rowCount
		fmt.Printf("%8d %s\n", rowCount, f)
	}
	fmt.Printf("%8d total\n", totalCount)
}
