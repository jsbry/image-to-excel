package main

import (
	"fmt"
	"testing"
)

func TestPASS(t *testing.T) {
	paths, err := ImgFileList()
	if err != nil {
		t.Error(err)
	} else {
		for _, path := range paths {
			fmt.Println(path)
		}
	}
}
