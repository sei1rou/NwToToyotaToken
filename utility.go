package main

import (
	"fmt"
	"strconv"
	"time"
)

func WaToSeireki(nen string) string {

	if len(nen) != 9 {
		return nen
	} else {
		w := nen[0:1]
		y := nen[1 : 1+2]
		yi, _ := strconv.Atoi(y)
		m := nen[4 : 4+2]
		d := nen[7 : 7+2]

		switch w {
		case "M":
			yi = yi + 1867
		case "T":
			yi = yi + 1911
		case "S":
			yi = yi + 1925
		case "H":
			yi = yi + 1988
		default:
			yi = 0
		}

		if yi == 0 {
			return "err"
		} else {
			return fmt.Sprint(yi) + "/" + m + "/" + d
		}
	}
}

func nendo(JDay string) string {
	var nen int
	t, _ := time.Parse("2006-01-02", JDay)
	if t.Month() > 3 {
		nen = t.Year()
	} else {
		nen = t.Year() - 1
	}

	return strconv.Itoa(nen)
}

func Hantei(s string) string {

	switch s {
	case "":
		s = ""
	case "Ａ":
		s = "Ａ"
	case "Ｂ":
		s = "Ｂ"
	case "Ｃ":
		s = "Ｃ"
	case "Ｃ(1)":
		s = "Ｃ"
	case "Ｃ(2)":
		s = "Ｃ"
	case "Ｄ":
		s = "Ｄ"
	case "Ｅ":
		s = "Ｅ"
	case "Ｅ(1)":
		s = "Ｅ"
	case "Ｆ":
		s = "Ｆ"
	case "Ｇ":
		s = "Ｇ"
	default:
		s = "err"
	}
	return s
}

func HanteiCode(s string) string {

	switch s {
	case "":
		s = ""
	case "Ａ":
		s = "1"
	case "Ｂ":
		s = "2"
	case "Ｃ":
		s = "3"
	case "Ｄ":
		s = "4"
	case "Ｅ":
		s = "5"
	case "Ｆ":
		s = "6"
	case "Ｇ":
		s = "7"
	default:
		s = "err"
	}
	return s
}

func HanteiWeight(s string) int {
	weight := 0

	switch s {
	case "":
		weight = 1
	case "Ａ":
		weight = 1
	case "１":
		weight = 1
	case "Ｂ":
		weight = 2
	case "Ｃ":
		weight = 3
	case "Ｃ(1)":
		weight = 3
	case "Ｃ(2)":
		weight = 3
	case "２":
		weight = 3
	case "Ｇ":
		weight = 4
	case "Ｄ":
		weight = 5
	case "Ｅ":
		weight = 6
	case "３":
		weight = 6
	case "Ｆ":
		weight = 7
	default:
		weight = 0
	}
	return weight
}

func WeightToHantei(s int) string {
	Hantei := ""

	switch s {
	case 1:
		Hantei = "Ａ"
	case 2:
		Hantei = "Ｂ"
	case 3:
		Hantei = "Ｃ"
	case 4:
		Hantei = "Ｇ"
	case 5:
		Hantei = "Ｄ"
	case 6:
		Hantei = "Ｅ"
	case 7:
		Hantei = "Ｆ"
	default:
		Hantei = "err"
	}
	return Hantei
}

func WeightToHanteiCode(s int) string {
	Hantei := ""

	switch s {
	case 1:
		Hantei = "1"
	case 2:
		Hantei = "2"
	case 3:
		Hantei = "3"
	case 4:
		Hantei = "7"
	case 5:
		Hantei = "4"
	case 6:
		Hantei = "5"
	case 7:
		Hantei = "6"
	default:
		Hantei = "err"
	}
	return Hantei
}
