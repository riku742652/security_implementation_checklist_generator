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
		if question != "" && strings.Index(question, "※") == -1 {
			answer, err := file.GetCellValue(sheet_name, answer_cell)
			fmt.Println(question)
			// fmt.Println(answer)
			answer = strings.Replace(answer, "※\n", "", -1)
			answer_list := strings.Split(answer, "\n")

			for num, ans := range answer_list {
				fmt.Println(num+1, ans)
			}
			// fmt.Println(ans)
			if err = rows.Close(); err != nil {
				fmt.Println(err)
			}
			var str string
			fmt.Scan(&str)
			number, _ := strconv.Atoi(str)
			answer = strings.Replace(answer, "□", "■", number)
			file.SetCellValue(sheet_name, answer_cell, answer)

		}
		rows_num++
	}
	file.SaveAs("output.xlsx")
}
