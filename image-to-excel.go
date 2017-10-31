package main

import (
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"io/ioutil"
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

var (
	border_top    = `{"type":"top","color":"000000","style":1}`
	border_left   = `{"type":"left","color":"000000","style":1}`
	border_right  = `{"type":"right","color":"000000","style":1}`
	border_bottom = `{"type":"bottom","color":"000000","style":1}`
)

type FileInfos []os.FileInfo
type ByName struct {
	FileInfos
}

func main() {
	var now string = time.Now().Format("20060102 150405")

	xlsx, err := excelize.OpenFile("tmp.xlsx")
	if err != nil {
		fmt.Println("excelize.OpenFile :", err)
	}

	addpicture_format := `{"x_scale": %s, "y_scale": %s, "print_obj": true, "lock_aspect_ratio": false, "locked": false}`

	files, filePattern := FileList()
	for _, fileInfo := range files {
		// *FileInfo型
		var findName = (fileInfo).Name()
		var matched = true
		// lsのようなワイルドカード検索を行うため、path.Matchを呼び出す
		if filePattern != "" {
			matched, _ = path.Match(filePattern, findName)
		}
		// path.Matchでマッチした場合、ファイル名を表示
		if matched != true {
			continue
		}

		// フォルダ判定
		var isDir, _ = IsDirectory(findName)
		if isDir == true {
			continue
		}
		pos := strings.LastIndex(findName, ".")
		ext := findName[pos:]
		lowerEXT := strings.ToLower(ext)

		if lowerEXT == ".png" || lowerEXT == ".jpg" || lowerEXT == ".jpeg" || lowerEXT == ".gif" {
			fmt.Printf("%s\n", findName)
			var dstName string
			switch ext {
			case ".PNG":
				dstName = now + strings.Replace(findName, ext, ".png", -1)
			case ".JPG":
				dstName = now + strings.Replace(findName, ext, ".jpg", -1)
			case ".JPEG":
				dstName = now + strings.Replace(findName, ext, ".jpeg", -1)
			case ".GIF":
				dstName = now + strings.Replace(findName, ext, ".gif", -1)
			default:
				dstName = now + strings.Replace(findName, ext, ext, -1)
			}
			src, err := os.Open(findName)
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
			//fmt.Println(image_format)

			cell_start++
			err = xlsx.AddPicture("Sheet1", fmt.Sprintf("B%d", cell_start), dstName, image_format)
			if err != nil {
				fmt.Println("貼り付けエラー :", err)
				continue
			}
			cell_start += 18

			defer os.Remove(dstName)

			output++
		}
	}

	// 保存
	if output > 0 {
		err := xlsx.SaveAs("写真貼付#" + now + ".xlsx")
		if err != nil {
			fmt.Println(err)
			os.Exit(1)
		}
	}
}

// 指定されたファイル名がディレクトリかどうか調べる
func IsDirectory(name string) (isDir bool, err error) {
	fInfo, err := os.Stat(name)
	if err != nil {
		return false, err
	}
	return fInfo.IsDir(), nil
}

func FileList() (FileInfos, string) {
	var arg string

	// カレントディレクトリの取得
	var curDir, _ = os.Getwd()
	curDir += "/"

	// 引数が取得できなければ、カレントディレクトリを使用
	if arg == "" {
		arg = curDir
	}

	// ディレクトリとファイル名に分割して格納
	var dirName, filePattern = path.Split(arg)

	// ディレクトリが無いならばカレントディレクトリを使用
	if dirName == "" {
		dirName = curDir
	}

	// 取得しようとしているパスがディレクトリかチェック
	var isDir, _ = IsDirectory(dirName + filePattern)

	// ディレクトリならば、そのディレクトリ配下のファイルを調べる。
	if isDir == true {
		dirName = dirName + filePattern
		filePattern = ""
	}

	// ディレクトリ内のファイル情報の読み込み[] *os.FileInfoが返る。
	fileInfos, err := ioutil.ReadDir(dirName)

	// ディレクトリの読み込みに失敗したらエラーで終了
	if err != nil {
		fmt.Errorf("ディレクトリの読み込みに失敗しました。 %s\n", err)
		os.Exit(1)
	}

	return fileInfos, filePattern
}

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
