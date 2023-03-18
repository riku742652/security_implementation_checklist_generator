package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	file, err := excelize.OpenFile("checklist.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	sheet_name := "チェックリスト"
	rows, err := file.Rows(sheet_name)
	if err != nil {
		fmt.Println(err)
		return
	}

	rows_num := 1
	for rows.Next() {
		_, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
		}
		question_cell := "G" + strconv.Itoa(rows_num)
		answer_cell := "E" + strconv.Itoa(rows_num)
		question, err := file.GetCellValue(sheet_name, question_cell)
		if question != "" && strings.Index(question, "※") == -1 && strings.Index(question, "。") != -1 {
			answer, err := file.GetCellValue(sheet_name, answer_cell)
			// 複数項目のケア
			// すでに対応済みの場合はスキップ
			if strings.Index(answer, "■ 対応済") != -1 {
				rows_num++
				continue
			}
			if strings.Index(answer, "■ 未対策") != -1 {
				answer = strings.Replace(answer, "■ 未対策", "□ 未対策", -1)
			}
			if strings.Index(answer, "■ 対応不要") != -1 {
				answer = strings.Replace(answer, "■ 対応不要", "□ 対応不要", -1)
			}
			fmt.Println(question)
			answer = strings.Replace(answer, "※\n", "", -1)
			answer_list := strings.Split(answer, "\n")

			for num, ans := range answer_list {
				fmt.Println(num+1, strings.Replace(ans, "□", "", -1))
			}
			if err = rows.Close(); err != nil {
				fmt.Println(err)
			}
			var str string
			for {
				fmt.Print("Answer: ")
				fmt.Scan(&str)
				number, err := strconv.Atoi(str)

				if err != nil || number > len(answer_list) {
					fmt.Println("Error: 値が不正です")
				} else {
					test_str := strings.Replace(answer_list[number-1], "□", "■", -1)
					answer = strings.Replace(answer, answer_list[number-1], test_str, -1)
					file.SetCellValue(sheet_name, answer_cell, answer)
					break
				}
			}
		}
		rows_num++
	}
	file.SaveAs("output.xlsx")
	fmt.Println("Output Completed.")
}
