package main

import (
	"log"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

func ConversionCobalt(filename string, inRecs [][]string, coRecs [][]string, dateRecs [][]string) {
	// コバルトデータ変換
	var vcell *xlsx.Cell
	var r int
	var cell string

	recLen := 64 //出力するレコードの項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	//会社毎に健診データファイルを作成する
	for _, coRec := range coRecs {

		excelName := filename + coRec[1] + "コバルト健診データ" + day.Format("20060102") + ".xlsx"
		excelFile := xlsx.NewFile()
		xlsx.SetDefaultFont(11, "ＭＳ Ｐゴシック")
		sheet, err := excelFile.AddSheet("データ")
		failOnError(err)

		// 1行目（タイトル）
		for I, _ = range cRec {
			cRec[I] = ""
		}
		cRec[4] = "idou.sya_bg"
		cRec[10] = "knk_kenkork.jushin_date"
		cRec[13] = "knk_kenkork_kensa.kensa_val_03"
		cRec[20] = "knk_kenkork_kensa.kensa_val_026"
		cRec[21] = "knk_kenkork_kensa.kensa_val_027"
		cRec[22] = "knk_kenkork_kensa.kensa_val_028"
		cRec[23] = "knk_kenkork_kensa.kensa_val_029"
		cRec[24] = "knk_kenkork_kensa.kensa_val_030"
		cRec[25] = "knk_kenkork_kensa.kensa_val_031"
		cRec[26] = "knk_kenkork_kensa.kensa_val_032"
		cRec[27] = "knk_kenkork_kensa.kensa_val_033"
		cRec[28] = "knk_kenkork_kensa.kensa_val_034"
		cRec[29] = "knk_kenkork_kensa.kensa_val_035"
		cRec[35] = "knk_kenkork_kensa.kensa_val_08"
		cRec[36] = "knk_kenkork_kensa.kensa_val_09"
		cRec[38] = "knk_kenkork_kensa.kensa_val_07"
		cRec[41] = "knk_kenkork_kensa.kensa_val_011"
		cRec[44] = "knk_kenkork_kensa.kensa_val_012"
		cRec[47] = "knk_kenkork_kensa.kensa_val_013"
		cRec[50] = "knk_kenkork_kensa.kensa_val_010"
		cRec[52] = "knk_kenkork_kensa.kensa_val_014"
		cRec[53] = "knk_kenkork_kensa.kensa_val_015"
		cRec[54] = "knk_kenkork_kensa.kensa_val_016"
		cRec[55] = "knk_kenkork_kensa.kensa_val_017"
		cRec[56] = "knk_kenkork_kensa.kensa_val_01"
		cRec[57] = "knk_kenkork_kensa.kensa_val_02"
		cRec[58] = "knk_kenkork_kensa.kensa_val_05"
		cRec[59] = "knk_kenkork_kensa.kensa_val_06"
		cRec[60] = "knk_kenkork_kensa.kensa_val_024"
		cRec[61] = "knk_kenkork_kensa.kensa_val_025"
		cRec[62] = "knk_kenkork_kensa.hantei_val_010"
		cRec[63] = "knk_kenkork_kensa.hantei_val_014"

		//writer.Write(cRec)
		row := sheet.AddRow()
		for _, cell = range cRec {
			//sheet.Cell(0, c).Value = cell
			vcell = row.AddCell()
			vcell.Value = cell
		}

		// 2行目（タイトル）
		for I, _ = range cRec {
			cRec[I] = ""
		}
		cRec[4] = "#社員番号"
		cRec[10] = "受診日付"
		cRec[13] = "検査コード03_医療機関側検査値"
		cRec[20] = "検査コード026_医療機関側検査値"
		cRec[21] = "検査コード027_医療機関側検査値"
		cRec[22] = "検査コード028_医療機関側検査値"
		cRec[23] = "検査コード029_医療機関側検査値"
		cRec[24] = "検査コード030_医療機関側検査値"
		cRec[25] = "検査コード031_医療機関側検査値"
		cRec[26] = "検査コード032_医療機関側検査値"
		cRec[27] = "検査コード033_医療機関側検査値"
		cRec[28] = "検査コード034_医療機関側検査値"
		cRec[29] = "検査コード035_医療機関側検査値"
		cRec[35] = "検査コード08_医療機関側検査値"
		cRec[36] = "検査コード09_医療機関側検査値"
		cRec[38] = "検査コード07_医療機関側検査値"
		cRec[41] = "検査コード011_医療機関側検査値"
		cRec[44] = "検査コード012_医療機関側検査値"
		cRec[47] = "検査コード013_医療機関側検査値"
		cRec[50] = "検査コード010_医療機関側検査値"
		cRec[52] = "検査コード014_医療機関側検査値"
		cRec[53] = "検査コード015_医療機関側検査値"
		cRec[54] = "検査コード016_医療機関側検査値"
		cRec[55] = "検査コード017_医療機関側検査値"
		cRec[56] = "検査コード01_医療機関側検査値"
		cRec[57] = "検査コード02_医療機関側検査値"
		cRec[58] = "検査コード05_医療機関側検査値"
		cRec[59] = "検査コード06_医療機関側検査値"
		cRec[60] = "検査コード024_医療機関側検査値"
		cRec[61] = "検査コード025_医療機関側検査値"
		cRec[62] = "検査コード010_医療機関側判定結果"
		cRec[63] = "検査コード014_医療機関側判定結果"

		//writer.Write(cRec)
		row = sheet.AddRow()
		for _, cell = range cRec {
			vcell = row.AddCell()
			vcell.Value = cell
		}

		// 3行目（タイトル）
		for I, _ = range cRec {
			cRec[I] = ""
		}
		cRec[0] = "所属cd１"
		cRec[1] = "所属名１"
		cRec[2] = "所属cd２"
		cRec[3] = "所属名２"
		cRec[4] = "社員No"
		cRec[5] = "ﾌﾘｶﾞﾅ"
		cRec[6] = "受診者名"
		cRec[7] = "性別"
		cRec[8] = "生年月日"
		cRec[9] = "年齢"
		cRec[10] = "受診日"
		cRec[11] = "受診番号"
		cRec[12] = "◆特化コバルト及び無機化合物"
		cRec[13] = "ｺﾊﾞﾙﾄ_作業名"
		cRec[14] = "ｺﾊﾞﾙﾄ_従事年数"
		cRec[15] = "ｺﾊﾞﾙﾄ_従事年数_月"
		cRec[16] = "ｺﾊﾞﾙﾄ_作業時間"
		cRec[17] = "ｺﾊﾞﾙﾄ_作業時間_分"
		cRec[18] = "ｺﾊﾞﾙﾄ_従事日数"
		cRec[19] = "ｺﾊﾞﾙﾄ_従事日数_週月"
		cRec[20] = "ｺﾊﾞﾙﾄ_作業工程に変化"
		cRec[21] = "ｺﾊﾞﾙﾄ_局所排気装置"
		cRec[22] = "ｺﾊﾞﾙﾄ_全体換気装置"
		cRec[23] = "ｺﾊﾞﾙﾄ_防毒マスク"
		cRec[24] = "ｺﾊﾞﾙﾄ_保護手袋"
		cRec[25] = "ｺﾊﾞﾙﾄ_保護メガネ"
		cRec[26] = "ｺﾊﾞﾙﾄ_保護衣"
		cRec[27] = "ｺﾊﾞﾙﾄ_取扱量・使用頻度"
		cRec[28] = "ｺﾊﾞﾙﾄ_大量のばく露"
		cRec[29] = "ｺﾊﾞﾙﾄ_直接触れる作業"
		cRec[30] = "ｺﾊﾞﾙﾄ_せき"
		cRec[31] = "ｺﾊﾞﾙﾄ_息苦しさ"
		cRec[32] = "ｺﾊﾞﾙﾄ_息切れ"
		cRec[33] = "ｺﾊﾞﾙﾄ_喘鳴"
		cRec[34] = "ｺﾊﾞﾙﾄ_皮膚炎"
		cRec[35] = "ｺﾊﾞﾙﾄ_自覚症状１"
		cRec[36] = "ｺﾊﾞﾙﾄ_自覚症状２"
		cRec[37] = "ｺﾊﾞﾙﾄ_自覚症状３"
		cRec[38] = "ｺﾊﾞﾙﾄ_既往歴１"
		cRec[39] = "ｺﾊﾞﾙﾄ_既往歴２"
		cRec[40] = "ｺﾊﾞﾙﾄ_既往歴３"
		cRec[41] = "ｺﾊﾞﾙﾄ_診察_呼吸器所見１"
		cRec[42] = "ｺﾊﾞﾙﾄ_診察_呼吸器所見２"
		cRec[43] = "ｺﾊﾞﾙﾄ_診察_呼吸器所見３"
		cRec[44] = "ｺﾊﾞﾙﾄ_診察_皮膚所見１"
		cRec[45] = "ｺﾊﾞﾙﾄ_診察_皮膚所見２"
		cRec[46] = "ｺﾊﾞﾙﾄ_診察_皮膚所見３"
		cRec[47] = "ｺﾊﾞﾙﾄ_診察_その他所見１"
		cRec[48] = "ｺﾊﾞﾙﾄ_診察_その他所見２"
		cRec[49] = "ｺﾊﾞﾙﾄ_診察_その他所見３"
		cRec[50] = "ｺﾊﾞﾙﾄ_診察判定"
		cRec[51] = "ｺﾊﾞﾙﾄ_診察判定コメント"
		cRec[52] = "ｺﾊﾞﾙﾄ_管理区分"
		cRec[53] = "医療機関判定（ｺﾊﾞﾙﾄ）"
		cRec[54] = "医療機関名称"
		cRec[55] = "健康診断を実施した医師の氏名"
		cRec[56] = "特定化学物質業務名"
		cRec[57] = "健診種別"
		cRec[58] = "ｺﾊﾞﾙﾄ_従事年数"
		cRec[59] = "ｺﾊﾞﾙﾄ_取扱い終了時期"
		cRec[60] = "ｺﾊﾞﾙﾄ_作業時間"
		cRec[61] = "ｺﾊﾞﾙﾄ_従事日数_週月"
		cRec[62] = "ｺﾊﾞﾙﾄ_診察判定"
		cRec[63] = "ｺﾊﾞﾙﾄ_管理区分"

		//writer.Write(cRec)
		row = sheet.AddRow()
		for _, cell = range cRec {
			//sheet.Cell(0, c).Value = cell
			vcell = row.AddCell()
			vcell.Value = cell
		}

		// 4行目以降（データ）
		inRecsMax := len(inRecs)
		for J := 1; J < inRecsMax; J++ {
			for I, _ = range cRec {
				cRec[I] = ""
			}

			if inRecs[J][0] == coRec[0] && inRecs[J][161] == "●" { //コバルトを受診しているか確認
				// 社員番号の桁数確認
				if len(inRecs[J][4]) != 10 {
					log.Printf("社員番号が10桁ではありません:%v\r\n", inRecs[J][4])
				}
				// 0.所属cd１
				cRec[0] = inRecs[J][0]

				// 1.所属名１
				cRec[1] = inRecs[J][1]

				// 2.所属cd２
				cRec[2] = inRecs[J][2]

				// 3.所属名２
				cRec[3] = inRecs[J][3]

				// 4.社員No
				cRec[4] = inRecs[J][4]

				// 5.ﾌﾘｶﾞﾅ
				cRec[5] = inRecs[J][5]

				// 6.受診者名
				cRec[6] = inRecs[J][6]

				// 7.性別
				cRec[7] = inRecs[J][7]

				// 8.生年月日
				cRec[8] = WaToSeireki(inRecs[J][8])

				// 9.年齢
				cRec[9] = inRecs[J][9]

				// 10.受診日
				cRec[10] = strings.Replace(inRecs[J][10], "-", "/", -1)

				// 11.受診番号
				cRec[11] = inRecs[J][11]

				// 12.◆特化コバルト及び無機化合物
				cRec[12] = inRecs[J][161]

				// 13.ｺﾊﾞﾙﾄ_作業名
				cRec[13] = inRecs[J][162]

				// 14.ｺﾊﾞﾙﾄ_従事年数
				cRec[14] = inRecs[J][163]

				// 15.ｺﾊﾞﾙﾄ_従事年数_月
				cRec[15] = inRecs[J][164]

				// 16.ｺﾊﾞﾙﾄ_作業時間
				cRec[16] = inRecs[J][165]

				// 17.ｺﾊﾞﾙﾄ_作業時間_分
				cRec[17] = inRecs[J][166]

				// 18.ｺﾊﾞﾙﾄ_従事日数
				cRec[18] = inRecs[J][167]

				// 19.ｺﾊﾞﾙﾄ_従事日数_週月
				cRec[19] = inRecs[J][168]

				// 20.ｺﾊﾞﾙﾄ_作業工程に変化
				cRec[20] = inRecs[J][169]

				// 21.ｺﾊﾞﾙﾄ_局所排気装置
				cRec[21] = inRecs[J][170]

				// 22.ｺﾊﾞﾙﾄ_全体換気装置
				cRec[22] = inRecs[J][171]

				// 23.ｺﾊﾞﾙﾄ_防毒マスク
				cRec[23] = inRecs[J][172]

				// 24.ｺﾊﾞﾙﾄ_保護手袋
				cRec[24] = inRecs[J][173]

				// 25.ｺﾊﾞﾙﾄ_保護メガネ
				cRec[25] = inRecs[J][174]

				// 26.ｺﾊﾞﾙﾄ_保護衣
				cRec[26] = inRecs[J][175]

				// 27.ｺﾊﾞﾙﾄ_取扱量・使用頻度
				cRec[27] = inRecs[J][176]

				// 28.ｺﾊﾞﾙﾄ_大量のばく露
				cRec[28] = inRecs[J][177]

				// 29.ｺﾊﾞﾙﾄ_直接触れる作業
				cRec[29] = inRecs[J][178]

				// 30.ｺﾊﾞﾙﾄ_せき
				cRec[30] = inRecs[J][179]

				// 31.ｺﾊﾞﾙﾄ_息苦しさ
				cRec[31] = inRecs[J][180]

				// 32.ｺﾊﾞﾙﾄ_息切れ
				cRec[32] = inRecs[J][181]

				// 33.ｺﾊﾞﾙﾄ_喘鳴
				cRec[33] = inRecs[J][182]

				// 34.ｺﾊﾞﾙﾄ_皮膚炎
				cRec[34] = inRecs[J][183]

				// 35.ｺﾊﾞﾙﾄ_自覚症状１
				cRec[35] = inRecs[J][184]

				// 36.ｺﾊﾞﾙﾄ_自覚症状２
				cRec[36] = inRecs[J][185]

				// 37.ｺﾊﾞﾙﾄ_自覚症状３
				cRec[37] = inRecs[J][186]

				// 38.ｺﾊﾞﾙﾄ_既往歴１
				cRec[38] = inRecs[J][187]

				// 39.ｺﾊﾞﾙﾄ_既往歴２
				cRec[39] = inRecs[J][188]

				// 40.ｺﾊﾞﾙﾄ_既往歴３
				cRec[40] = inRecs[J][189]

				// 41.ｺﾊﾞﾙﾄ_診察_呼吸器所見１
				cRec[41] = inRecs[J][190]

				// 42.ｺﾊﾞﾙﾄ_診察_呼吸器所見２
				cRec[42] = inRecs[J][191]

				// 43.ｺﾊﾞﾙﾄ_診察_呼吸器所見３
				cRec[43] = inRecs[J][192]

				// 44.ｺﾊﾞﾙﾄ_診察_皮膚所見１
				cRec[44] = inRecs[J][193]

				// 45.ｺﾊﾞﾙﾄ_診察_皮膚所見２
				cRec[45] = inRecs[J][194]

				// 46.ｺﾊﾞﾙﾄ_診察_皮膚所見３
				cRec[46] = inRecs[J][195]

				// 47.ｺﾊﾞﾙﾄ_診察_その他所見１
				cRec[47] = inRecs[J][196]

				// 48.ｺﾊﾞﾙﾄ_診察_その他所見２
				cRec[48] = inRecs[J][197]

				// 49.ｺﾊﾞﾙﾄ_診察_その他所見３
				cRec[49] = inRecs[J][198]

				// 50.ｺﾊﾞﾙﾄ_診察判定
				cRec[50] = Hantei(inRecs[J][199])
				if Hantei(inRecs[J][199]) == "err" {
					log.Print("診察所見判定にエラーがあります。\r\n")
				}

				// 51.ｺﾊﾞﾙﾄ_診察判定コメント
				cRec[51] = inRecs[J][200]

				// 52.ｺﾊﾞﾙﾄ_管理区分
				cRec[52] = inRecs[J][201]

				// 53.医療機関判定（ｺﾊﾞﾙﾄ）
				sogo := ""
				var h [1][2]string
				h[0][0] = Hantei(inRecs[J][199]) //診察所見判定
				h[0][1] = inRecs[J][200]         //診察所見所見

				hKigo := [...]string{"Ｆ", "Ｅ", "３", "Ｄ", "２", "Ｇ", "Ｃ"}
				for k := 0; k < 7; k++ {
					for l := 0; l < 1; l++ {
						if h[l][0] == hKigo[k] {
							if sogo == "" {
								sogo = h[l][1]
							} else {
								sogo = sogo + "　" + h[l][1]
							}
						}
					}
				}

				if sogo == "" {
					sogo = "検査範囲では異常ありません"
				}

				cRec[53] = sogo

				// 54.医療機関名称
				cRec[54] = "医療法人社団　松英会"

				// 55.健康診断を実施した医師の氏名
				cRec[55] = "寺門　節雄"

				// 56.特定化学物質業務名
				cRec[56] = "ｺﾊﾞﾙﾄ取扱業務"

				// 57.健診種別
				cRec[57] = "定期"

				// 58.ｺﾊﾞﾙﾄ_従事年数
				jyujiYear := ""
				if inRecs[J][163] != "" {
					jyujiYear = inRecs[J][163] + "年"
				}

				jyujiMon := ""
				if inRecs[J][164] != "" {
					jyujiMon = inRecs[J][164] + "ヵ月"
				}

				if jyujiYear != "" && jyujiMon != "" {
					cRec[58] = jyujiYear + " " + jyujiMon
				} else {
					cRec[58] = jyujiYear + jyujiMon
				}

				// 59.ｺﾊﾞﾙﾄ_取扱い終了時期
				if dateRecs[0][5] != "コバルト取扱い終了年月日" {
					log.Print("「コバルト取扱い終了年月日」が見つかりませんでした。")
					failOnError(doError())
				}

				findflag := false
				for l, _ := range dateRecs {
					if cRec[4] == dateRecs[l][0] {
						if cRec[8] != dateRecs[l][3] {
							log.Printf("コバルト生年月日の不一致: %v %v %v != %v\r\n", cRec[4], cRec[6], cRec[8], dateRecs[l][3])
						}

						edate := dateRecs[l][5]
						if edate != "" {
							cRec[59] = edate[0:4] + "年" + edate[5:7] + "月" + edate[8:] + "日"
						} else {
							cRec[59] = edate
						}
						findflag = true
						break
					}
				}

				if findflag == false {
					log.Printf("取扱い終了名簿に対象がいません。 %v %v", cRec[4], cRec[6])
					cRec[59] = "err"
				}

				// 60.ｺﾊﾞﾙﾄ_作業時間
				sagyoHour := ""
				if inRecs[J][165] != "" {
					sagyoHour = inRecs[J][165] + "時間"
				}

				sagyoMin := ""
				if inRecs[J][166] != "" {
					sagyoMin = inRecs[J][166] + "分"
				}

				if sagyoHour != "" && sagyoMin != "" {
					cRec[60] = sagyoHour + " " + sagyoMin
				} else {
					cRec[60] = sagyoHour + sagyoMin
				}

				// 61.ｺﾊﾞﾙﾄ_従事日数_週月
				jyuji := ""
				if inRecs[J][167] != "" {
					jyuji = inRecs[J][167] + "日"
				}

				jyujiWM := ""
				if inRecs[J][168] != "" {
					jyujiWM = "/" + inRecs[J][168]
				}

				cRec[61] = jyuji + jyujiWM

				// 62.ｺﾊﾞﾙﾄ_診察判定
				cRec[62] = HanteiCode(inRecs[J][199])
				if HanteiCode(inRecs[J][199]) == "err" {
					log.Print("診察判定にエラーがあります\r\n")
				}

				// 63.ｺﾊﾞﾙﾄ_管理区分
				cRec[63] = HanteiCode(inRecs[J][201])
				if HanteiCode(inRecs[J][201]) == "err" {
					log.Print("管理区分にエラーがあります\r\n")
				}

				//writer.Write(cRec)
				row = sheet.AddRow()
				for _, cell = range cRec {
					// sheet.Cell(r, c).Value = cell
					vcell = row.AddCell()
					vcell.Value = cell
				}
				r++
			}
		}

		//writer.Flush()
		err = excelFile.Save(excelName)
		failOnError(err)
	}

}
