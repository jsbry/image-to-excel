package main

import (
	"fmt"
	"os"
	"testing"
)

func Test_ImgFileList(t *testing.T) {
	paths, err := ImgFileList()
	if err != nil {
		t.Error(err)
	} else {
		for _, path := range paths {
			fmt.Println(path)
		}
	}
}

func Test_Run(t *testing.T) {
	code, err := Run(default_image_w)
	if err != nil {
		t.Error(err)
	}
	os.Exit(code)
}
