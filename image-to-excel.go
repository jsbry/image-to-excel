package main

import (
	"errors"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"io/ioutil"
	"math"
	"os"
	"path"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize"
)

var image_w = 412.0
var output = 0
var cell_start = 3
var ss = 0
var sheet_slice = 30

var (
	border_top    = `{"type":"top","color":"000000","style":1}`
	border_left   = `{"type":"left","color":"000000","style":1}`
	border_right  = `{"type":"right","color":"000000","style":1}`
	border_bottom = `{"type":"bottom","color":"000000","style":1}`
)

func main() {
	code, err := Run()
	if err != nil {
		fmt.Println(err)
	}
	os.Exit(code)
}

func Run() (int, error) {
	var now string = time.Now().Format("20060102 150405")

	xlsx, err := excelize.OpenFile("tmp.xlsx")
	if err != nil {
		fmt.Println(err)
		return 1, errors.New("テンプレートファイルの読み込みに失敗しました")
	}

	addpicture_format := `{"x_scale": %s, "y_scale": %s, "print_obj": true, "lock_aspect_ratio": false, "locked": false}`

	paths, err := ImgFileList()
	if err != nil {
		fmt.Println(err)
		return 1, errors.New("画像リスト取得に失敗しました")
	}

	// 4枚ごとにSheet追加 #debug
	sheetCount := float64(len(paths)) / float64(sheet_slice)
	for idx := 1; idx <= int(math.Ceil(sheetCount)); idx++ {
		copy_sheetname := fmt.Sprintf("Sheet%d", idx)
		if copy_sheetname == "Sheet1" {
			continue
		}
		to_idx := xlsx.NewSheet(fmt.Sprintf("Sheet%d", idx))
		err = xlsx.CopySheet(1, to_idx)
	}

	SheetCell := make(map[string]int)
	for i, path := range paths {
		slice_num := i / sheet_slice
		Sheet_name := fmt.Sprintf("Sheet%d", (slice_num + 1))
		_, sheet_exists := SheetCell[Sheet_name]
		if sheet_exists == false {
			SheetCell[Sheet_name] = cell_start
		}

		pos := strings.LastIndex(path, ".")
		ext := path[pos:]
		lowerEXT := strings.ToLower(ext)

		if lowerEXT == ".png" || lowerEXT == ".jpg" || lowerEXT == ".jpeg" || lowerEXT == ".gif" {
			fmt.Printf("%s\n", path)
			var dstName string
			switch ext {
			case ".PNG":
				dstName = now + strings.Replace(path, ext, ".png", -1)
			case ".JPG":
				dstName = now + strings.Replace(path, ext, ".jpg", -1)
			case ".JPEG":
				dstName = now + strings.Replace(path, ext, ".jpeg", -1)
			case ".GIF":
				dstName = now + strings.Replace(path, ext, ".gif", -1)
			default:
				dstName = now + strings.Replace(path, ext, ext, -1)
			}
			src, err := os.Open(path)
			if err != nil {
				fmt.Println(err)
			}

			dst, err := os.Create(dstName)
			if err != nil {
				fmt.Println(err)
			}

			_, err = io.Copy(dst, src)
			if err != nil {
				fmt.Println(err)
			}

			err = src.Close()
			if err != nil {
				fmt.Println(err)
			}
			err = dst.Close()
			if err != nil {
				fmt.Println(err)
			}

			// 画像読み込み
			file, err := os.Open(dstName)
			if err != nil {
				fmt.Println("os.Open :", err)
			}

			imgConfig, _, err := image.DecodeConfig(file)
			if err != nil {
				fmt.Println("image.DecodeConfig :", err)
			}
			fmt.Printf("画像 幅: %dpx, 高さ: %dpx \n\n", imgConfig.Width, imgConfig.Height)

			err = file.Close()
			if err != nil {
				fmt.Println(err)
			}

			widthRatio := float64(image_w) / float64(imgConfig.Width)
			//fmt.Println(image_w, " / ", imgConfig.Width, " = ", widthRatio)

			image_h := float64(imgConfig.Height) * widthRatio
			heightRatio := float64(image_h) / float64(imgConfig.Height) * 1.111246943765281 // * 1.17454663212435 // * 1.28875236294896
			//fmt.Println(image_h, " / ", imgConfig.Height, " = ", heightRatio)

			string_image_w := strconv.FormatFloat(widthRatio, 'f', 10, 64)
			string_image_h := strconv.FormatFloat(heightRatio, 'f', 10, 64)

			image_format := fmt.Sprintf(addpicture_format, string_image_w, string_image_h)

			SheetCell[Sheet_name]++
			err = xlsx.AddPicture(Sheet_name, fmt.Sprintf("B%d", SheetCell[Sheet_name]), dstName, image_format)
			if err != nil {
				fmt.Println("貼り付けエラー :", err)
				continue
			}
			SheetCell[Sheet_name] += 18

			defer os.Remove(dstName)

			output++
		}
	}

	// 保存
	if output > 0 {
		err := xlsx.SaveAs("写真貼付#" + now + ".xlsx")
		if err != nil {
			fmt.Println(err)
			return 1, errors.New("生成したExcelの保存に失敗しました")
		}
	}
	return 0, nil
}

/**
 * 画像一覧
 */
func ImgFileList() ([]string, error) {
	dir, _ := os.Getwd()
	files, err := ioutil.ReadDir(dir)
	if err != nil {
		return []string{}, err
	}

	var paths []string
	for _, file := range files {
		if !file.IsDir() {
			lowerEXT := strings.ToLower(path.Ext(file.Name()))
			if lowerEXT == ".png" || lowerEXT == ".jpg" || lowerEXT == ".jpeg" || lowerEXT == ".gif" {
				paths = append(paths, file.Name())
			}
		}
	}

	return paths, nil
}
