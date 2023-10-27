package main

import (
	"encoding/json"
	"github.com/dataismo/xlsx"
	"io"
	"net/http"
)

type Character struct {
	Id      int    `json:"id"`
	Name    string `json:"name"`
	Status  string `json:"status"`
	Species string `json:"species"`
	Type    string `json:"type"`
	Gender  string `json:"gender"`
}

type Response struct {
	Results []Character `json:"results"`
}

func main() {

	var response Response

	excel := xlsx.New("General", 2)

	sheet, err := excel.NewSheet("Demo", 2)

	if err != nil {
		panic(err)
	}

	res, err := http.Get("https://rickandmortyapi.com/api/character")

	if err != nil {
		panic(err)
	}

	jsonData, err := io.ReadAll(res.Body)

	if err != nil {
		panic(err)
	}

	if err = json.Unmarshal(jsonData, &response); err != nil {
		panic(err)
	}

	if err != nil {
		panic(err)
	}

	sheet.AddRowHeader(
		"ID",
		"Genero",
		"Especie",
		"Tipo",
		"Estatus",
		"Nombre",
	)

	for _, c := range response.Results {
		err = sheet.AddRow(
			c.Id,
			c.Gender,
			c.Species,
			c.Type,
			c.Status,
			c.Name,
		)

		if err != nil {
			panic(err)
		}
	}

	if err = sheet.SetColumnStyle(excel.Styles.MoneyFormat, "A"); err != nil {
		panic(err)
	}

	if err = sheet.Subtotal("A1", "A"); err != nil {
		panic(err)
	}

	sheet.SetCellValue("B1", "total")

	if err = sheet.Freeze(2, 0); err != nil {
		panic(err)
	}

	if err = excel.SaveFile("test.xlsx"); err != nil {
		panic(err)
	}

}
