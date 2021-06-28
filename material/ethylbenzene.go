package material

import (
	"log"
	"strings"
	"time"

	"../utility"
	"github.com/tealeg/xlsx"
)

func ConversionEthylbenzene(filename string, inRecs [][]string, coRecs [][]string, dateRecs [][]string) {
	// エチルベンゼンデータ変換
	var vcell *xlsx.Cell
	var r int
	var cell string

	recLen := 62 //出力するレコードの項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	//会社毎に健診データファイルを作成する
	for _, coRec := range coRecs {

		excelName := filename + coRec[1] + "エチルベンゼン健診データ" + day.Format("20060102") + ".xlsx"
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
		cRec[20] = "knk_kenkork_kensa.kensa_val_027"
		cRec[21] = "knk_kenkork_kensa.kensa_val_028"
		cRec[22] = "knk_kenkork_kensa.kensa_val_029"
		cRec[23] = "knk_kenkork_kensa.kensa_val_030"
		cRec[24] = "knk_kenkork_kensa.kensa_val_031"
		cRec[25] = "knk_kenkork_kensa.kensa_val_032"
		cRec[26] = "knk_kenkork_kensa.kensa_val_033"
		cRec[27] = "knk_kenkork_kensa.kensa_val_034"
		cRec[28] = "knk_kenkork_kensa.kensa_val_013"
		cRec[29] = "knk_kenkork_kensa.kensa_val_014"
		cRec[37] = "knk_kenkork_kensa.kensa_val_008"
		cRec[38] = "knk_kenkork_kensa.kensa_val_009"
		cRec[40] = "knk_kenkork_kensa.kensa_val_007"
		cRec[43] = "knk_kenkork_kensa.kensa_val_011"
		cRec[46] = "knk_kenkork_kensa.kensa_val_010"
		cRec[50] = "knk_kenkork_kensa.kensa_val_014"
		cRec[51] = "knk_kenkork_kensa.kensa_val_016"
		cRec[52] = "knk_kenkork_kensa.kensa_val_017"
		cRec[53] = "knk_kenkork_kensa.kensa_val_018"
		cRec[54] = "knk_kenkork_kensa.kensa_val_001"
		cRec[55] = "knk_kenkork_kensa.kensa_val_002"
		cRec[56] = "knk_kenkork_kensa.kensa_val_005"
		cRec[57] = "knk_kenkork_kensa.kensa_val_006"
		cRec[58] = "knk_kenkork_kensa.kensa_val_025"
		cRec[59] = "knk_kenkork_kensa.kensa_val_026"
		cRec[60] = "knk_kenkork_kensa.hantei_val_010"
		cRec[61] = "knk_kenkork_kensa.hantei_val_015"

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
		cRec[20] = "検査コード027_医療機関側検査値"
		cRec[21] = "検査コード028_医療機関側検査値"
		cRec[22] = "検査コード029_医療機関側検査値"
		cRec[23] = "検査コード030_医療機関側検査値"
		cRec[24] = "検査コード031_医療機関側検査値"
		cRec[25] = "検査コード032_医療機関側検査値"
		cRec[26] = "検査コード033_医療機関側検査値"
		cRec[27] = "検査コード034_医療機関側検査値"
		cRec[28] = "検査コード013_医療機関側検査値"
		cRec[29] = "検査コード014_医療機関側検査値"
		cRec[37] = "検査コード008_医療機関側検査値"
		cRec[38] = "検査コード009_医療機関側検査値"
		cRec[40] = "検査コード007_医療機関側検査値"
		cRec[43] = "検査コード011_医療機関側検査値"
		cRec[46] = "検査コード010_医療機関側検査値"
		cRec[50] = "検査コード014_医療機関側検査値"
		cRec[51] = "検査コード016_医療機関側検査値"
		cRec[52] = "検査コード017_医療機関側検査値"
		cRec[53] = "検査コード018_医療機関側検査値"
		cRec[54] = "検査コード001_医療機関側検査値"
		cRec[55] = "検査コード002_医療機関側検査値"
		cRec[56] = "検査コード005_医療機関側検査値"
		cRec[57] = "検査コード006_医療機関側検査値"
		cRec[58] = "検査コード025_医療機関側検査値"
		cRec[59] = "検査コード026_医療機関側検査値"
		cRec[60] = "検査コード010_医療機関側判定結果"
		cRec[61] = "検査コード015_医療機関側判定結果"

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
		cRec[12] = "◆特化エチルベンゼン"
		cRec[13] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業名"
		cRec[14] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事年数"
		cRec[15] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事年数_月"
		cRec[16] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業時間"
		cRec[17] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業時間_分"
		cRec[18] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事日数"
		cRec[19] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事日数_週月"
		cRec[20] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業工程に変化"
		cRec[21] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_局所排気装置"
		cRec[22] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_全体換気装置"
		cRec[23] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_防毒マスク"
		cRec[24] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_送気マスク"
		cRec[25] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_取扱量・使用頻度"
		cRec[26] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_大量のばく露"
		cRec[27] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_直接触れる作業"
		cRec[28] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_マンデル酸"
		cRec[29] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_マンデル酸_分布"
		cRec[30] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_眼の痛み"
		cRec[31] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_発赤"
		cRec[32] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_せき"
		cRec[33] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_咽頭痛"
		cRec[34] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_鼻腔刺激症状"
		cRec[35] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_頭痛"
		cRec[36] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_倦怠感"
		cRec[37] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_自覚症状１"
		cRec[38] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_自覚症状２"
		cRec[39] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_自覚症状３"
		cRec[40] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_既往歴１"
		cRec[41] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_既往歴２"
		cRec[42] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_既往歴３"
		cRec[43] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察所見１"
		cRec[44] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察所見２"
		cRec[45] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察所見３"
		cRec[46] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察判定"
		cRec[47] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察判定コメント"
		cRec[48] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_マンデル酸判定"
		cRec[49] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_マンデル酸判定コメント"
		cRec[50] = "管理区分             (ｴﾁﾙﾍﾞﾝｾﾞﾝ)"
		cRec[51] = "医療機関判定（ｴﾁﾙﾍﾞﾝｾﾞﾝ）"
		cRec[52] = "医療機関名称"
		cRec[53] = "健康診断を実施した医師の氏名"
		cRec[54] = "特定化学物質業務名"
		cRec[55] = "健診種別"
		cRec[56] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事年数"
		cRec[57] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_取扱い終了時期"
		cRec[58] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業時間"
		cRec[59] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事日数_週月"
		cRec[60] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察判定"
		cRec[61] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ_管理区分"

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

			if inRecs[J][0] == coRec[0] && inRecs[J][122] == "●" {
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
				cRec[8] = utility.WaToSeireki(inRecs[J][8])

				// 9.年齢
				cRec[9] = inRecs[J][9]

				// 10.受診日
				cRec[10] = strings.Replace(inRecs[J][10], "-", "/", -1)

				// 11.受診番号
				cRec[11] = inRecs[J][11]

				// 12.◆特化エチルベンゼン
				cRec[12] = inRecs[J][122]

				// 13.ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業名
				cRec[13] = inRecs[J][123]

				// 14.ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事年数
				cRec[14] = inRecs[J][124]

				// 15.ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事年数_月
				cRec[15] = inRecs[J][125]

				// 16.ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業時間
				cRec[16] = inRecs[J][126]

				// 17.ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業時間_分
				cRec[17] = inRecs[J][127]

				// 18.ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事日数
				cRec[18] = inRecs[J][128]

				// 19.ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事日数_週月
				cRec[19] = inRecs[J][129]

				// 20.ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業工程に変化
				cRec[20] = inRecs[J][130]

				// 21.ｴﾁﾙﾍﾞﾝｾﾞﾝ_局所排気装置
				cRec[21] = inRecs[J][131]

				// 22.ｴﾁﾙﾍﾞﾝｾﾞﾝ_全体換気装置
				cRec[22] = inRecs[J][132]

				// 23.ｴﾁﾙﾍﾞﾝｾﾞﾝ_防毒マスク
				cRec[23] = inRecs[J][133]

				// 24.ｴﾁﾙﾍﾞﾝｾﾞﾝ_送気マスク
				cRec[24] = inRecs[J][134]

				// 25.ｴﾁﾙﾍﾞﾝｾﾞﾝ_取扱量・使用頻度
				cRec[25] = inRecs[J][135]

				// 26.ｴﾁﾙﾍﾞﾝｾﾞﾝ_大量のばく露
				cRec[26] = inRecs[J][136]

				// 27.ｴﾁﾙﾍﾞﾝｾﾞﾝ_直接触れる作業
				cRec[27] = inRecs[J][137]

				// 28.ｴﾁﾙﾍﾞﾝｾﾞﾝ_マンデル酸
				cRec[28] = inRecs[J][138]

				// 29.ｴﾁﾙﾍﾞﾝｾﾞﾝ_マンデル酸_分布
				cRec[29] = inRecs[J][139]

				// 30.ｴﾁﾙﾍﾞﾝｾﾞﾝ_眼の痛み
				cRec[30] = inRecs[J][140]

				// 31.ｴﾁﾙﾍﾞﾝｾﾞﾝ_発赤
				cRec[31] = inRecs[J][141]

				// 32.ｴﾁﾙﾍﾞﾝｾﾞﾝ_せき
				cRec[32] = inRecs[J][142]

				// 33.ｴﾁﾙﾍﾞﾝｾﾞﾝ_咽頭痛
				cRec[33] = inRecs[J][143]

				// 34.ｴﾁﾙﾍﾞﾝｾﾞﾝ_鼻腔刺激症状
				cRec[34] = inRecs[J][144]

				// 35.ｴﾁﾙﾍﾞﾝｾﾞﾝ_頭痛
				cRec[35] = inRecs[J][145]

				// 36.ｴﾁﾙﾍﾞﾝｾﾞﾝ_倦怠感
				cRec[36] = inRecs[J][146]

				// 37.ｴﾁﾙﾍﾞﾝｾﾞﾝ_自覚症状１
				cRec[37] = inRecs[J][147]

				// 38.ｴﾁﾙﾍﾞﾝｾﾞﾝ_自覚症状２
				cRec[38] = inRecs[J][148]

				// 39.ｴﾁﾙﾍﾞﾝｾﾞﾝ_自覚症状３
				cRec[39] = inRecs[J][149]

				// 40.ｴﾁﾙﾍﾞﾝｾﾞﾝ_既往歴１
				cRec[40] = inRecs[J][150]

				// 41.ｴﾁﾙﾍﾞﾝｾﾞﾝ_既往歴２
				cRec[41] = inRecs[J][151]

				// 42.ｴﾁﾙﾍﾞﾝｾﾞﾝ_既往歴３
				cRec[42] = inRecs[J][152]

				// 43.ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察所見１
				cRec[43] = inRecs[J][153]

				// 44.ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察所見２
				cRec[44] = inRecs[J][154]

				// 45.ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察所見３
				cRec[45] = inRecs[J][155]

				// 46.ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察判定
				cRec[46] = utility.Hantei(inRecs[J][156])
				if utility.Hantei(inRecs[J][156]) == "err" {
					log.Print("診察所見判定にエラーがあります。\r\n")
				}

				// 47.ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察判定コメント
				cRec[47] = inRecs[J][157]

				// 48.ｴﾁﾙﾍﾞﾝｾﾞﾝ_マンデル酸判定
				cRec[48] = inRecs[J][158]

				// 49.ｴﾁﾙﾍﾞﾝｾﾞﾝ_マンデル酸判定コメント
				cRec[49] = inRecs[J][159]

				// 50.管理区分             (ｴﾁﾙﾍﾞﾝｾﾞﾝ)
				cRec[50] = inRecs[J][160]

				// 51.医療機関判定（ｴﾁﾙﾍﾞﾝｾﾞﾝ）
				sogo := ""
				var h [2][2]string
				h[0][0] = utility.Hantei(inRecs[J][156]) //診察所見判定
				h[0][1] = inRecs[J][157]                 //診察所見所見
				h[1][0] = utility.Hantei(inRecs[J][158]) //マンデル酸判定
				h[1][1] = inRecs[J][159]                 //マンデル酸所見

				hKigo := [...]string{"Ｆ", "Ｅ", "３", "Ｄ", "２", "Ｇ", "Ｃ"}
				for k := 0; k < 7; k++ {
					for l := 0; l < 2; l++ {
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

				cRec[51] = sogo

				// 52.医療機関名称
				cRec[52] = "医療法人社団　松英会"

				// 53.健康診断を実施した医師の氏名
				cRec[53] = "寺門　節雄"

				// 54.特定化学物質業務名
				cRec[54] = "ｴﾁﾙﾍﾞﾝｾﾞﾝ取扱業務"

				// 55.健診種別
				cRec[55] = "定期"

				// 56.ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事年数
				jyujiYear := ""
				if inRecs[J][124] != "" {
					jyujiYear = inRecs[J][124] + "年"
				}

				jyujiMon := ""
				if inRecs[J][125] != "" {
					jyujiMon = inRecs[J][125] + "ヵ月"
				}

				if jyujiYear != "" && jyujiMon != "" {
					cRec[56] = jyujiYear + " " + jyujiMon
				} else {
					cRec[56] = jyujiYear + jyujiMon
				}

				// 57.ｴﾁﾙﾍﾞﾝｾﾞﾝ_取扱い終了時期
				if dateRecs[0][4] != "エチルベンゼン取扱い終了年月日" {
					log.Print("「エチルベンゼン取扱い終了年月日」が見つかりませんでした。")
					failOnError(doError())
				}

				findflag := false
				for l, _ := range dateRecs {
					if cRec[4] == dateRecs[l][0] {
						if cRec[8] != dateRecs[l][3] {
							log.Printf("エチルベンゼン生年月日の不一致: %v %v %v != %v\r\n", cRec[4], cRec[6], cRec[8], dateRecs[l][3])
						}

						edate := dateRecs[l][4]
						if edate != "" {
							cRec[57] = edate[0:4] + "年" + edate[5:7] + "月" + edate[8:] + "日"
						} else {
							cRec[57] = edate
						}
						findflag = true
						break
					}
				}

				if findflag == false {
					log.Printf("取扱い終了名簿に対象がいません。 %v %v", cRec[4], cRec[6])
					cRec[57] = "err"
				}

				// 58.ｴﾁﾙﾍﾞﾝｾﾞﾝ_作業時間
				sagyoHour := ""
				if inRecs[J][126] != "" {
					sagyoHour = inRecs[J][126] + "時間"
				}

				sagyoMin := ""
				if inRecs[J][127] != "" {
					sagyoMin = inRecs[J][127] + "分"
				}

				if sagyoHour != "" && sagyoMin != "" {
					cRec[58] = sagyoHour + " " + sagyoMin
				} else {
					cRec[58] = sagyoHour + sagyoMin
				}

				// 59.ｴﾁﾙﾍﾞﾝｾﾞﾝ_従事日数_週月
				jyuji := ""
				if inRecs[J][128] != "" {
					jyuji = inRecs[J][128] + "日"
				}

				jyujiWM := ""
				if inRecs[J][129] != "" {
					jyujiWM = "/" + inRecs[J][129]
				}

				cRec[59] = jyuji + jyujiWM

				// 60.ｴﾁﾙﾍﾞﾝｾﾞﾝ_診察判定
				cRec[60] = utility.HanteiCode(inRecs[J][156])
				if utility.HanteiCode(inRecs[J][156]) == "err" {
					log.Print("診察判定にエラーがあります\r\n")
				}

				// 61.ｴﾁﾙﾍﾞﾝｾﾞﾝ_管理区分
				cRec[61] = utility.HanteiCode(inRecs[J][160])
				if utility.HanteiCode(inRecs[J][160]) == "err" {
					log.Print("診察判定にエラーがあります\r\n")
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
