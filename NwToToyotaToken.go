package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"time"

	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
)

/*
func failOnError(err error) {
	if err != nil {
		log.Fatal("Error:", err)
	}
}

func doError() error {
	return errors.New("エラーが発生しました。")
}
*/

func main() {
	// ログファイル準備
	logfile, err := os.OpenFile("./log.txt", os.O_APPEND|os.O_CREATE|os.O_WRONLY, os.ModePerm)
	failOnError(err)
	defer logfile.Close()

	log.SetOutput(logfile)

	log.Print("Start\r\n")

	// ファイルの確認
	flag.Parse()
	filesu := flag.NArg()
	if filesu != 2 {
		log.Print("健診データを取扱い終了年月日、２つのファイルをドロップしてください。")
		failOnError(doError())
	}

	// 健診データと取扱い終了年月日のファイルチェックとファイル名の確認
	fkenshin, fenddate := filecheck(flag.Arg(0), flag.Arg(1))

	// ファイルを読み込んで二次元配列に入れる
	records := readfile(fkenshin)

	recordsEnddate := readfile(fenddate)

	// 出力する会社を調査
	coRecods := coSurvey(records)

	//出力するフォルダを作成
	outDir := dirCreate(flag.Arg(0))

	// 有機溶剤データの変換
	ConversionYuki(outDir, records, coRecods)

	// クロロホルム他物質データの変換（ジクロロ他）
	ConversionChloroform(outDir, records, coRecods, recordsEnddate)

	// コバルトデータの変換
	ConversionCobalt(outDir, records, coRecods, recordsEnddate)

	// エチルベンゼンデータの変換
	ConversionEthylbenzene(outDir, records, coRecods, recordsEnddate)

	// 溶接ヒュームデータの変換
	ConversionWeldingfume(outDir, records, coRecods, recordsEnddate)

	log.Print("Finish !\r\n")

}

func readfile(filename string) [][]string {
	//入力ファイル準備
	infile, err := os.Open(filename)
	failOnError(err)
	defer infile.Close()

	reader := csv.NewReader(transform.NewReader(infile, japanese.ShiftJIS.NewDecoder()))
	reader.Comma = '\t'

	//CSVファイルを２次元配列に展開
	readrecords := make([][]string, 0)
	for {
		record, err := reader.Read() // 1行読み出す
		if err == io.EOF {
			break
		} else {
			failOnError(err)
		}

		readrecords = append(readrecords, record)
	}

	return readrecords
}

func coSurvey(records [][]string) [][]string {

	/*
		companys := [][]string{{"2000100100000001", "トヨタモビリティ東京（株）", "0"},
			{"2000100100000011", "東京トヨタカーライフサービス（株）", "0"},
			{"2000100100000026", "ティーシーサービス（株）", "0"},
			{"2000100100009002", "トヨタ東京カローラ（株）", "0"},
			{"2000100100009004", "（株）センチュリーサービス", "0"},
		}
	*/

	companys := [][]string{{"2000100100000001", "トヨタ東京販売ホールディングス（株）", "0"},
		{"2000100100000002", "東京トヨタ自動車（株）", "0"},
		{"2000100100000003", "東京トヨペット（株）", "0"},
		{"2000100100000006", "ネッツトヨタ東京（株）", "0"},
		{"2000100100000020", "ＴＭプロサービス（株）", "0"},
	}

	coRecMax := len(records)
	for i := 1; i < coRecMax; i++ {
		for _, com := range companys {
			if com[0] == records[i][0] {
				count, _ := strconv.Atoi(com[2])
				com[2] = fmt.Sprint(count + 1)
				break
			}
		}
	}

	outCompanys := make([][]string, 0)
	for j, _ := range companys {
		if companys[j][2] != "0" {
			outCompanys = append(outCompanys, companys[j])
			// log.Print(companys[j][2] + " " + companys[j][0] + ":" + companys[j][1] + "\r\n")
		}
	}

	return outCompanys

}

func dirCreate(path string) string {
	day := time.Now()
	outDir, _ := filepath.Split(path)
	outDirPlus := outDir + "/トヨタモビリティ東京" + day.Format("20060102")

	if err := os.Mkdir(outDirPlus, 0777); err != nil {
		log.Print(outDirPlus + "\r\n")
		log.Print("出力先のディレクトリを作成できませんでした\r\n")
		return outDir
	} else {
		return outDirPlus + "/"
	}
}

func filecheck(f1 string, f2 string) (string, string) {
	// ファイルの中身を確認

	// １番目にドロップされたファイルの読み込み
	filecheck, err := os.Open(f1)
	failOnError(err)
	defer filecheck.Close()

	readerFilecheck := csv.NewReader(transform.NewReader(filecheck, japanese.ShiftJIS.NewDecoder()))
	readerFilecheck.Comma = '\t'

	readFilecheck, err := readerFilecheck.Read()
	failOnError(err)

	//ファイルの先頭が「社員番号」なら取扱い終了年月日ファイル
	fData := ""
	fEnddate := ""
	if readFilecheck[0] == "社員番号" {
		fData = f2
		fEnddate = f1
	} else {
		fData = f1
		fEnddate = f2
	}

	// 取扱い終了年月日のファイルの確認
	fileEnddate, err := os.Open(fEnddate)
	failOnError(err)
	defer fileEnddate.Close()

	readerEnddate := csv.NewReader(transform.NewReader(fileEnddate, japanese.ShiftJIS.NewDecoder()))
	readerEnddate.Comma = '\t'

	readEnddate, err := readerEnddate.Read()
	failOnError(err)

	//タイトルチェック
	fcheck := false
	fhead := []string{"社員番号", "フリガナ", "氏名", "生年月日", "エチルベンゼン取扱い終了年月日", "コバルト取扱い終了年月日", "ジクロロメタン取扱い終了年月日", "スチレン取扱い終了年月日", "MIBK取扱い終了年月日", "アーク溶接取扱い終了年月日"}
	for i, _ := range fhead {
		if readEnddate[i] != fhead[i] {
			fcheck = true
		}
	}

	if fcheck {
		log.Print("ファイルのタイトルが違います\r\n")
		log.Print("社員番号, フリガナ, 氏名, 生年月日, エチルベンゼン取扱い終了年月日, コバルト取扱い終了年月日, ジクロロメタン取扱い終了年月日, スチレン取扱い終了年月日, MIBK取扱い終了年月日,アーク溶接取扱い終了年月日\r\n")
		failOnError(doError())
	}

	return fData, fEnddate
}
