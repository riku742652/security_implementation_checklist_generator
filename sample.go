package main

import (
    "fmt"

    "github.com/xuri/excelize/v2"
)

func main() {
    f := excelize.NewFile()
    // ワークシートを作成する
    index := f.NewSheet("Sheet2")
    // セルの値を設定
    f.SetCellValue("Sheet2", "A2", "Hello world.")
    f.SetCellValue("Sheet1", "B2", 100)
    // ワークブックのデフォルトワークシートを設定します
    f.SetActiveSheet(index)
    // 指定されたパスに従ってファイルを保存します
    if err := f.SaveAs("Book1.xlsx"); err != nil {
        fmt.Println(err)
    }
}