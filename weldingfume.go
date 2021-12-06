package main

import (
	"log"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

func ConversionWeldingfume(filename string, inRecs [][]string, coRecs [][]string, dateRecs [][]string) {
	// 溶接ヒュームデータ変換
	var vcell *xlsx.Cell
	var r int
	var cell string

	recLen := 61 //出力するレコードの項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	//会社毎に健診データファイルを作成する
	for _, coRec := range coRecs {

		excelName := filename + coRec[1] + "溶接ヒューム健診データ" + day.Format("20060102") + ".xlsx"
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
		cRec[13] = "knk_kenkork_kensa.kensa_val_003"
		cRec[20] = "knk_kenkork_kensa.kensa_val_010"
		cRec[21] = "knk_kenkork_kensa.kensa_val_011"
		cRec[22] = "knk_kenkork_kensa.kensa_val_012"
		cRec[23] = "knk_kenkork_kensa.kensa_val_013"
		cRec[25] = "knk_kenkork_kensa.kensa_val_015"
		cRec[26] = "knk_kenkork_kensa.kensa_val_022"
		cRec[27] = "knk_kenkork_kensa.kensa_val_023"
		cRec[36] = "knk_kenkork_kensa.kensa_val_017"
		cRec[37] = "knk_kenkork_kensa.kensa_val_018"
		cRec[39] = "knk_kenkork_kensa.kensa_val_016"
		cRec[42] = "knk_kenkork_kensa.kensa_val_020"
		cRec[43] = "knk_kenkork_kensa.kensa_val_021"
		cRec[46] = "knk_kenkork_kensa.kensa_val_019"
		cRec[48] = "knk_kenkork_kensa.kensa_val_024"
		cRec[50] = "knk_kenkork_kensa.kensa_val_025"
		cRec[51] = "knk_kenkork_kensa.kensa_val_026"
		cRec[52] = "knk_kenkork_kensa.kensa_val_027"
		cRec[53] = "knk_kenkork_kensa.kensa_val_001"
		cRec[54] = "knk_kenkork_kensa.kensa_val_002"
		cRec[55] = "knk_kenkork_kensa.kensa_val_005"
		cRec[56] = "knk_kenkork_kensa.kensa_val_006"
		cRec[57] = "knk_kenkork_kensa.kensa_val_008"
		cRec[58] = "knk_kenkork_kensa.kensa_val_009"
		cRec[59] = "knk_kenkork_kensa.hantei_val_019"
		cRec[60] = "knk_kenkork_kensa.hantei_val_024"

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
		cRec[13] = "検査コード003_医療機関側検査値"
		cRec[20] = "検査コード010_医療機関側検査値"
		cRec[21] = "検査コード011_医療機関側検査値"
		cRec[22] = "検査コード012_医療機関側検査値"
		cRec[23] = "検査コード013_医療機関側検査値"
		cRec[25] = "検査コード015_医療機関側検査値"
		cRec[26] = "検査コード022_医療機関側検査値"
		cRec[27] = "検査コード023_医療機関側検査値"
		cRec[36] = "検査コード017_医療機関側検査値"
		cRec[37] = "検査コード018_医療機関側検査値"
		cRec[39] = "検査コード016_医療機関側検査値"
		cRec[42] = "検査コード020_医療機関側検査値"
		cRec[43] = "検査コード021_医療機関側検査値"
		cRec[46] = "検査コード019_医療機関側検査値"
		cRec[48] = "検査コード024_医療機関側検査値"
		cRec[50] = "検査コード025_医療機関側検査値"
		cRec[51] = "検査コード026_医療機関側検査値"
		cRec[52] = "検査コード027_医療機関側検査値"
		cRec[53] = "検査コード001_医療機関側検査値"
		cRec[54] = "検査コード002_医療機関側検査値"
		cRec[55] = "検査コード005_医療機関側検査値"
		cRec[56] = "検査コード006_医療機関側検査値"
		cRec[57] = "検査コード008_医療機関側検査値"
		cRec[58] = "検査コード009_医療機関側検査値"
		cRec[59] = "検査コード019_医療機関側判定結果"
		cRec[60] = "検査コード024_医療機関側判定結果"

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
		cRec[4] = "#社員No"
		cRec[5] = "ﾌﾘｶﾞﾅ"
		cRec[6] = "受診者名"
		cRec[7] = "性別"
		cRec[8] = "生年月日"
		cRec[9] = "年齢"
		cRec[10] = "受診日"
		cRec[11] = "受診番号"
		cRec[12] = "◆特化金属アーク溶接作業等"
		cRec[13] = "作業名"
		cRec[14] = "従事年"
		cRec[15] = "従事月"
		cRec[16] = "作業時間"
		cRec[17] = "分"
		cRec[18] = "作業日数"
		cRec[19] = "作業日数(週・月)"
		cRec[20] = "作業工程に変化"
		cRec[21] = "全体換気装置"
		cRec[22] = "保護マスク"
		cRec[23] = "取扱量・使用頻度"
		cRec[24] = "大量のばく露"
		cRec[25] = "直接触れる作業"
		cRec[26] = "握力右(１回)"
		cRec[27] = "握力左(１回)"
		cRec[28] = "せき"
		cRec[29] = "たん"
		cRec[30] = "よだれがでる"
		cRec[31] = "発汗異常"
		cRec[32] = "手指のふるえ"
		cRec[33] = "字が書きにくい"
		cRec[34] = "握力低下感"
		cRec[35] = "歩行障害"
		cRec[36] = "自覚症状１"
		cRec[37] = "自覚症状２"
		cRec[38] = "自覚症状３"
		cRec[39] = "既往歴１"
		cRec[40] = "既往歴２"
		cRec[41] = "既往歴３"
		cRec[42] = "パーキンソン症候群様症状"
		cRec[43] = "診察所見１"
		cRec[44] = "診察所見２"
		cRec[45] = "診察所見３"
		cRec[46] = "H_診察"
		cRec[47] = "H_診察 "
		cRec[48] = "管理区分"
		cRec[49] = "管理区分"
		cRec[50] = "医療機関判定（溶接ヒューム）"
		cRec[51] = "医療機関名称"
		cRec[52] = "健康診断を実施した医師の氏名"
		cRec[53] = "特定化学物質業務名"
		cRec[54] = "健診種別"
		cRec[55] = "アーク溶接_従事年数"
		cRec[56] = "アーク溶接_作業終了時期"
		cRec[57] = "アーク溶接_作業時間"
		cRec[58] = "アーク溶接_従事日数_週月"
		cRec[59] = "アーク溶接_診察判定"
		cRec[60] = "アーク溶接_管理区分"

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

			if inRecs[J][0] == coRec[0] && inRecs[J][275] == "●" { //溶接ヒュームを受診しているか確認
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

				// 12.◆特化金属アーク溶接作業等
				cRec[12] = inRecs[J][275]

				// 13.作業名
				cRec[13] = inRecs[J][276]

				// 14.従事年
				cRec[14] = inRecs[J][277]

				// 15.従事月
				cRec[15] = inRecs[J][278]

				// 16.作業時間
				cRec[16] = inRecs[J][279]

				// 17.分
				cRec[17] = inRecs[J][280]

				// 18.作業日数
				cRec[18] = inRecs[J][281]

				// 19.作業日数(週・月)
				cRec[19] = inRecs[J][282]

				// 20.作業工程に変化
				cRec[20] = inRecs[J][283]

				// 21.全体換気装置
				cRec[21] = inRecs[J][284]

				// 22.保護マスク
				cRec[22] = inRecs[J][285]

				// 23.取扱量・使用頻度
				cRec[23] = inRecs[J][286]

				// 24.大量のばく露
				cRec[24] = inRecs[J][287]

				// 25.直接触れる作業
				cRec[25] = inRecs[J][288]

				// 26.握力右
				cRec[26] = inRecs[J][289]

				// 27.握力左
				cRec[27] = inRecs[J][290]

				// 28.せき
				cRec[28] = inRecs[J][291]

				// 29.たん
				cRec[29] = inRecs[J][292]

				// 30.よだれがでる
				cRec[30] = inRecs[J][293]

				// 31.発汗異常
				cRec[31] = inRecs[J][294]

				// 32.手指のふるえ
				cRec[32] = inRecs[J][295]

				// 33.字が書きにくい
				cRec[33] = inRecs[J][296]

				// 34.握力低下感
				cRec[34] = inRecs[J][297]

				// 35.歩行障害
				cRec[35] = inRecs[J][298]

				// 36.自覚症状１
				cRec[36] = inRecs[J][299]

				// 37.自覚症状２
				cRec[37] = inRecs[J][300]

				// 38.自覚症状３
				cRec[38] = inRecs[J][301]

				// 39.既往歴１
				cRec[39] = inRecs[J][302]

				// 40.既往歴２
				cRec[40] = inRecs[J][303]

				// 41.既往歴３
				cRec[41] = inRecs[J][304]

				// 42.パーキンソン症候群様症状
				cRec[42] = inRecs[J][305]

				// 43.診察所見１
				cRec[43] = inRecs[J][306]

				// 44.診察所見２
				cRec[44] = inRecs[J][307]

				// 45.診察所見３
				cRec[45] = inRecs[J][308]

				// 46.溶接フューム_診察判定
				cRec[46] = Hantei(inRecs[J][309])
				if Hantei(inRecs[J][309]) == "err" {
					log.Print("診察所見判定にエラーがあります。\r\n")
				}

				// 47.溶接フューム_診察判定コメント
				cRec[47] = inRecs[J][310]

				// 48.管理区分
				cRec[48] = inRecs[J][311]

				// 49.管理区分コメント
				cRec[49] = inRecs[J][312]

				// 50.医療機関判定（溶接ヒューム）
				sogo := ""
				var h [1][2]string
				h[0][0] = Hantei(inRecs[J][309]) //診察所見判定
				h[0][1] = inRecs[J][310]         //診察所見所見

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

				cRec[50] = sogo

				// 51.医療機関名称
				cRec[51] = "医療法人社団　松英会"

				// 52.健康診断を実施した医師の氏名
				cRec[52] = "寺門　節雄"

				// 53.特定化学物質業務名
				cRec[53] = "アーク溶接作業（溶接ヒューム）"

				// 54.健診種別
				cRec[54] = "定期"

				// 55.アーク溶接_従事年数
				jyujiYear := ""
				if inRecs[J][277] != "" {
					jyujiYear = inRecs[J][277] + "年"
				}

				jyujiMon := ""
				if inRecs[J][278] != "" {
					jyujiMon = inRecs[J][278] + "ヵ月"
				}

				if jyujiYear != "" && jyujiMon != "" {
					cRec[55] = jyujiYear + " " + jyujiMon
				} else {
					cRec[55] = jyujiYear + jyujiMon
				}

				// 56.アーク溶接_作業終了時期
				if dateRecs[0][9] != "アーク溶接取扱い終了年月日" {
					log.Print("「アーク溶接取扱い終了年月日」が見つかりませんでした。")
					failOnError(doError())
				}

				findflag := false
				for l, _ := range dateRecs {
					if cRec[4] == dateRecs[l][0] {
						if cRec[8] != dateRecs[l][3] {
							log.Printf("アーク溶接生年月日の不一致: %v %v %v != %v\r\n", cRec[4], cRec[6], cRec[8], dateRecs[l][3])
						}

						edate := dateRecs[l][9]
						if edate != "" {
							cRec[56] = edate[0:4] + "年" + edate[5:7] + "月" + edate[8:] + "日"
						} else {
							cRec[56] = edate
						}
						findflag = true
						break
					}
				}

				if findflag == false {
					log.Printf("取扱い終了名簿に対象がいません。 %v %v", cRec[4], cRec[6])
					cRec[56] = "err"
				}

				// 57.アーク溶接_作業時間
				sagyoHour := ""
				if inRecs[J][279] != "" {
					sagyoHour = inRecs[J][279] + "時間"
				}

				sagyoMin := ""
				if inRecs[J][280] != "" {
					sagyoMin = inRecs[J][280] + "分"
				}

				if sagyoHour != "" && sagyoMin != "" {
					cRec[57] = sagyoHour + " " + sagyoMin
				} else {
					cRec[57] = sagyoHour + sagyoMin
				}

				// 58._アーク溶接従事日数_週月
				jyuji := ""
				if inRecs[J][281] != "" {
					jyuji = inRecs[J][281] + "日"
				}

				jyujiWM := ""
				if inRecs[J][282] != "" {
					jyujiWM = "/" + inRecs[J][282]
				}

				cRec[58] = jyuji + jyujiWM

				// 59.アーク溶接_診察判定
				cRec[59] = HanteiCode(inRecs[J][309])
				if HanteiCode(inRecs[J][309]) == "err" {
					log.Print("診察判定にエラーがあります\r\n")
				}

				// 60.アーク溶接_管理区分
				cRec[60] = HanteiCode(inRecs[J][311])
				if HanteiCode(inRecs[J][311]) == "err" {
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
