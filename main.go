package main

import (
	"fmt"
	"log"
	"strings"

	"github.com/mtslzr/pokeapi-go"
	"github.com/xuri/excelize/v2"
)

type move struct {
	nombre, tipo, clase string
	potencia, precision int
}

// func main() {

// 	r, err := pokeapi.Resource("move")

// 	if err != nil {
// 		log.Printf("error obteniendo: %V", err)
// 		return
// 	}

// 	var nombres []string

// 	for _, mov := range r.Results {

// 		nombres = append(nombres, mov.Name)

// 	}

// 	for i, nombre := range nombres {

// 		log.Println(nombre)

// 		if i == 30 {
// 			break
// 		}

// 		if strings.Contains(nombre, "_") {
// 			fmt.Printf("El nombre '%s' contiene un guión bajo.\n", nombre)
// 			break
// 		}
// 	}

// }

func main() {

	//Aca va el nombre de tu excel con la lista de movimientos en columna, en ingles y si tiene mas de una palabra entonces
	//separalas con un guion.
	f, err := excelize.OpenFile("movimientos.xlsx")

	if err != nil {
		log.Printf("error abriendo: %V", err)
		return
	}

	cols, err := f.GetCols("Sheet1")
	if err != nil {
		log.Printf("error obteniendo: %V", err)
		return
	}

	log.Println(cols)

	var errors []error
	var moves []move
	var move move

	for _, col := range cols {

		for _, moveName := range col {

			data, err := pokeapi.Move(moveName)

			if err != nil {

				log.Printf("err with %v: %v", move.nombre, err)

				errors = append(errors, err)

				continue

			}

			move.clase = data.DamageClass.Name
			move.nombre = data.Name
			move.potencia = data.Power
			move.precision = data.Accuracy
			move.tipo = data.Type.Name

			moves = append(moves, move)

		}

		if len(errors) > 0 {
			log.Printf("errors getting moves: %v", errors)
		}
	}

	movesFile := excelize.NewFile()

	_, err = movesFile.NewSheet("Sheet1")

	if err != nil {
		log.Printf("err guardando: %v", err)
	}

	for i, move := range moves {

		var valor string

		for j := 0; j < 5; j++ {
			letter, _ := getLetter(j)
			coord := strings.ToUpper(letter) + fmt.Sprint(i+1)

			switch j {
			case 0:
				valor = move.nombre
			case 1:
				valor = move.tipo
			case 2:
				valor = move.clase
			case 3:
				valor = fmt.Sprint(move.potencia)
			case 4:
				valor = fmt.Sprint(move.precision)
			}

			err := movesFile.SetCellValue("Sheet1", coord, valor)

			if err != nil {
				log.Printf("err setteando a %v: %v", valor, err)
				continue
			}
		}

	}

	//Aca le asignas el nombre al excel
	err = movesFile.SaveAs("monosexy.xlsx")

	if err != nil {
		log.Printf("err guardando: %v", err)
	}

}

func getLetter(index int) (string, error) {
	if index < 0 || index > 26 {
		return "", fmt.Errorf("index fuera de rango")
	}
	// El alfabeto español tiene 27 letras, incluyendo la "ñ"
	alfabeto := "abcdefghijklmnñopqrstuvwxyz"
	return string(alfabeto[index]), nil
}
