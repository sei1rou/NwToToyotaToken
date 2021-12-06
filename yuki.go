package main

import (
	"log"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

func ConversionYuki(filename string, inRecs [][]string, coRecs [][]string) {
	// 有機溶剤データ変換
	var vcell *xlsx.Cell
	var r int
	var cell string

	recLen := 122 //出力するレコードの項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	//会社毎に健診データファイルを作成する
	for _, coRec := range coRecs {

		/*
			outfile, err := os.Create(filename + coRec[1] + "健診データ" + day.Format("20060102") + ".txt")
			failOnError(err)
			defer outfile.Close()

			writer := csv.NewWriter(transform.NewWriter(outfile, japanese.ShiftJIS.NewEncoder()))
			writer.Comma = '\t'
			writer.UseCRLF = true
		*/

		excelName := filename + coRec[1] + "有機溶剤健診データ" + day.Format("20060102") + ".xlsx"
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
		cRec[12] = "knk_kenkork_kensa.kensa_val_022"
		cRec[13] = "knk_kenkork_kensa.kensa_val_031"
		cRec[14] = "knk_kenkork_kensa.kensa_val_032"
		cRec[15] = "knk_kenkork_kensa.kensa_val_034"
		cRec[16] = "knk_kenkork_kensa.kensa_val_035"
		cRec[17] = "knk_kenkork_kensa.kensa_val_036"
		cRec[61] = "knk_kenkork_kensa.kensa_val_024"
		cRec[62] = "knk_kenkork_kensa.kensa_val_025"
		cRec[63] = "knk_kenkork_kensa.kensa_val_026"
		cRec[64] = "knk_kenkork_kensa.kensa_val_027"
		cRec[65] = "knk_kenkork_kensa.kensa_val_028"
		cRec[66] = "knk_kenkork_kensa.kensa_val_029"
		cRec[67] = "knk_kenkork_kensa.kensa_val_002"
		cRec[68] = "knk_kenkork_kensa.kensa_val_003"
		cRec[69] = "knk_kenkork_kensa.kensa_val_004"
		cRec[79] = "knk_kenkork_kensa.kensa_val_010"
		cRec[80] = "knk_kenkork_kensa.kensa_val_011"
		cRec[81] = "knk_kenkork_kensa.kensa_val_012"
		cRec[82] = "knk_kenkork_kensa.kensa_val_013"
		cRec[83] = "knk_kenkork_kensa.kensa_val_014"
		cRec[84] = "knk_kenkork_kensa.kensa_val_015"
		cRec[85] = "knk_kenkork_kensa.kensa_val_016"
		cRec[86] = "knk_kenkork_kensa.kensa_val_018"
		cRec[87] = "knk_kenkork_kensa.kensa_val_019"
		cRec[91] = "knk_kenkork_kensa.kensa_val_017"
		cRec[94] = "knk_kenkork_kensa.kensa_val_021"
		cRec[112] = "knk_kenkork_kensa.kensa_val_040"
		cRec[113] = "knk_kenkork_kensa.kensa_val_041"
		cRec[114] = "knk_kenkork_kensa.kensa_val_042"
		cRec[115] = "knk_kenkork_kensa.kensa_val_001"
		cRec[116] = "knk_kenkork_kensa.kensa_val_005"
		cRec[117] = "knk_kenkork_kensa.kensa_val_006"
		cRec[118] = "knk_kenkork_kensa.kensa_val_008"
		cRec[119] = "knk_kenkork_kensa.kensa_val_009"
		cRec[120] = "knk_kenkork_kensa.kensa_val_020"
		cRec[121] = "knk_kenkork_kensa.hantei_val_020"

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
		cRec[12] = "検査コード022_医療機関側検査値"
		cRec[13] = "検査コード031_医療機関側検査値"
		cRec[14] = "検査コード032_医療機関側検査値"
		cRec[15] = "検査コード034_医療機関側検査値"
		cRec[16] = "検査コード035_医療機関側検査値"
		cRec[17] = "検査コード036_医療機関側検査値"
		cRec[61] = "検査コード024_医療機関側検査値"
		cRec[62] = "検査コード025_医療機関側検査値"
		cRec[63] = "検査コード026_医療機関側検査値"
		cRec[64] = "検査コード027_医療機関側検査値"
		cRec[65] = "検査コード028_医療機関側検査値"
		cRec[66] = "検査コード029_医療機関側検査値"
		cRec[67] = "検査コード002_医療機関側検査値"
		cRec[68] = "検査コード003_医療機関側検査値"
		cRec[69] = "検査コード004_医療機関側検査値"
		cRec[79] = "検査コード010_医療機関側検査値"
		cRec[80] = "検査コード011_医療機関側検査値"
		cRec[81] = "検査コード012_医療機関側検査値"
		cRec[82] = "検査コード013_医療機関側検査値"
		cRec[83] = "検査コード014_医療機関側検査値"
		cRec[84] = "検査コード015_医療機関側検査値"
		cRec[85] = "検査コード016_医療機関側検査値"
		cRec[86] = "検査コード018_医療機関側検査値"
		cRec[87] = "検査コード019_医療機関側検査値"
		cRec[91] = "検査コード017_医療機関側検査値"
		cRec[94] = "検査コード021_医療機関側検査値"
		cRec[112] = "検査コード040_医療機関側検査値"
		cRec[113] = "検査コード041_医療機関側検査値"
		cRec[114] = "検査コード042_医療機関側検査値"
		cRec[115] = "検査コード001_医療機関側検査値"
		cRec[116] = "検査コード005_医療機関側検査値"
		cRec[117] = "検査コード006_医療機関側検査値"
		cRec[118] = "検査コード008_医療機関側検査値"
		cRec[119] = "検査コード009_医療機関側検査値"
		cRec[120] = "検査コード020_医療機関側検査値"
		cRec[121] = "検査コード020_医療機関側判定結果"

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
		cRec[12] = "尿蛋白定性"
		cRec[13] = "赤血球数"
		cRec[14] = "血色素量"
		cRec[15] = "ＧＯＴ"
		cRec[16] = "ＧＰＴ"
		cRec[17] = "γ－ＧＴＰ"
		cRec[18] = "◆有機"
		cRec[19] = "1.ｱｾﾄﾝ"
		cRec[20] = "2.ｲｿﾌﾞﾁﾙｱﾙｺｰﾙ"
		cRec[21] = "3.ｲｿﾌﾟﾛﾋﾟﾙｱﾙｺｰﾙ"
		cRec[22] = "4.ｲｿﾍﾟﾝﾁﾙｱﾙｺｰﾙ"
		cRec[23] = "5.ｴﾁﾙｴｰﾃﾙ"
		cRec[24] = "6.ｴﾁﾚﾝｸﾞﾘｺｰﾙﾓﾉｴﾁﾙｴｰﾃﾙ"
		cRec[25] = "7.ｴﾁﾚﾝｸﾞﾘｺｰﾙﾓﾉｴﾁﾙｴｰﾃﾙｱｾﾃｰﾄ"
		cRec[26] = "8.ｴﾁﾚﾝｸﾞﾘｺｰﾙﾓﾉﾌﾞﾁﾙｴｰﾃﾙ"
		cRec[27] = "9.ｴﾁﾚﾝｸﾞﾘｺｰﾙﾓﾉﾒﾁﾙｴｰﾃﾙ"
		cRec[28] = "10.ｵﾙﾄ-ｼﾞｸﾛﾙﾍﾞﾝｾﾞﾝ"
		cRec[29] = "11.ｷｼﾚﾝ"
		cRec[30] = "12.ｸﾚｿﾞｰﾙ"
		cRec[31] = "13.ｸﾛﾙﾍﾞﾝｾﾞﾝ"
		cRec[32] = "15.酢酸ｲｿﾌﾞﾁﾙ"
		cRec[33] = "16.酢酸ｲｿﾌﾟﾛﾋﾟﾙ"
		cRec[34] = "17.酢酸ｲｿﾍﾟﾝﾁﾙ"
		cRec[35] = "18.酢酸ｴﾁﾙ"
		cRec[36] = "19.酢酸ﾌﾞﾁﾙ"
		cRec[37] = "20.酢酸ﾌﾟﾛﾋﾟﾙ"
		cRec[38] = "21.酢酸ﾍﾟﾝﾁﾙ"
		cRec[39] = "22.酢酸ﾒﾁﾙ"
		cRec[40] = "24.ｼｸﾛﾍｷｻﾉｰﾙ"
		cRec[41] = "25.ｼｸﾛﾍｷｻﾉﾝ"
		cRec[42] = "30.N,N-ｼﾞﾒﾁﾙﾎﾙﾑｱﾐﾄﾞ"
		cRec[43] = "34.ﾃﾄﾗﾋﾄﾞﾛﾌﾗﾝ"
		cRec[44] = "35.1,1,1-ﾄﾘｸﾛﾙｴﾀﾝ"
		cRec[45] = "37.ﾄﾙｴﾝ"
		cRec[46] = "39.ﾉﾙﾏﾙﾍｷｻﾝ"
		cRec[47] = "40.1-ﾌﾞﾀﾉｰﾙ"
		cRec[48] = "41.2-ﾌﾞﾀﾉｰﾙ"
		cRec[49] = "42.ﾒﾀﾉｰﾙ"
		cRec[50] = "44.ﾒﾁﾙｴﾁﾙｹﾄﾝ"
		cRec[51] = "45.ﾒﾁﾙｼｸﾛﾍｷｻﾉｰﾙ"
		cRec[52] = "46.ﾒﾁﾙｼｸﾛﾍｷｻﾉﾝ"
		cRec[53] = "47.ﾒﾁﾙ-ﾉﾙﾏﾙ-ﾌﾞﾁﾙｹﾄﾝ"
		cRec[54] = "48.ｶﾞｿﾘﾝ"
		cRec[55] = "49.ｺｰﾙﾀｰﾙﾅﾌｻ"
		cRec[56] = "50.石油ｴｰﾃﾙ"
		cRec[57] = "51.石油ﾅﾌｻ"
		cRec[58] = "52.石油ﾍﾞﾝｼﾞﾝ"
		cRec[59] = "53.ﾃﾚﾋﾞﾝ油"
		cRec[60] = "54.ﾐﾈﾗﾙｽﾋﾟﾘｯﾄ"
		cRec[61] = "馬尿酸"
		cRec[62] = "馬尿酸_分布"
		cRec[63] = "メチル馬尿酸"
		cRec[64] = "メチル馬尿酸_分布"
		cRec[65] = "2.5-ﾍｷｻﾝｼﾞｵﾝ"
		cRec[66] = "2.5-ﾍｷｻﾝｼﾞｵﾝ_分布"
		cRec[67] = "有機_健診種別"
		cRec[68] = "有機_作業名１"
		cRec[69] = "有機_作業名２"
		cRec[70] = "有機_業務名１"
		cRec[71] = "有機_業務名２"
		cRec[72] = "有機_業務名３"
		cRec[73] = "有機_従事年数"
		cRec[74] = "有機_従事年数_月"
		cRec[75] = "有機_作業時間"
		cRec[76] = "有機_作業時間_分"
		cRec[77] = "有機_従事日数"
		cRec[78] = "有機_従事日数_週月"
		cRec[79] = "有機_作業工程に変化"
		cRec[80] = "有機_局所排気装置"
		cRec[81] = "有機_防毒マスク"
		cRec[82] = "有機_不透過性手袋"
		cRec[83] = "有機_取扱量・使用頻度"
		cRec[84] = "有機_大量のばく露"
		cRec[85] = "有機_直接触れる作業"
		cRec[86] = "有機_自覚症状１"
		cRec[87] = "有機_自覚症状２"
		cRec[88] = "有機_自覚症状３"
		cRec[89] = "有機_自覚症状４"
		cRec[90] = "有機_自覚症状５"
		cRec[91] = "有機_既往歴１"
		cRec[92] = "有機_既往歴２"
		cRec[93] = "有機_既往歴３"
		cRec[94] = "有機_診察所見１"
		cRec[95] = "有機_診察所見２"
		cRec[96] = "有機_診察所見３"
		cRec[97] = "有機_診察判定"
		cRec[98] = "有機_診察判定コメント"
		cRec[99] = "有機_尿蛋白判定"
		cRec[100] = "有機_尿蛋白判定コメント"
		cRec[101] = "有機_貧血判定"
		cRec[102] = "有機_貧血判定コメント"
		cRec[103] = "有機_肝機能判定"
		cRec[104] = "有機_肝機能コメント"
		cRec[105] = "有機_馬尿酸判定"
		cRec[106] = "有機_馬尿酸判定コメント"
		cRec[107] = "有機_ﾒﾁﾙ馬尿酸判定"
		cRec[108] = "有機_ﾒﾁﾙ馬尿酸判定コメント"
		cRec[109] = "有機_2.5ﾍｷｻﾝｼﾞｵﾝ判定"
		cRec[110] = "有機_2.5ﾍｷｻﾝｼﾞｵﾝ判定コメント"
		cRec[111] = "有機_管理区分"
		cRec[112] = "医療機関判定（有機）"
		cRec[113] = "医療機関名称"
		cRec[114] = "健康診断を実施した医師の氏名"
		cRec[115] = "従事年数"
		cRec[116] = "健診対象有機溶剤"
		cRec[117] = "有機溶剤業務名"
		cRec[118] = "1日の作業時間（有機）"
		cRec[119] = "従事日数_週月（有機）"
		cRec[120] = "有機_診察判定"
		cRec[121] = "有機_診察判定結果"

		//writer.Write(cRec)
		row = sheet.AddRow()
		for _, cell = range cRec {
			//sheet.Cell(0, c).Value = cell
			vcell = row.AddCell()
			vcell.Value = cell
		}

		// 4行目以降（データ）
		r = 3
		inRecsMax := len(inRecs)
		for J := 1; J < inRecsMax; J++ {
			for I, _ = range cRec {
				cRec[I] = ""
			}

			if inRecs[J][0] == coRec[0] && inRecs[J][28] == "●" {
				// 0.社員番号
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

				// 4.#社員No
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

				// 12.尿蛋白定性
				if inRecs[J][109] != "" { //判定チェック
					cRec[12] = inRecs[J][12]
				} else {
					cRec[12] = ""
				}

				// 13.赤血球数
				// 14.血色素量
				if inRecs[J][111] != "" { //判定チェック
					cRec[13] = inRecs[J][13]
					cRec[14] = inRecs[J][14]
				} else {
					cRec[13] = ""
					cRec[14] = ""
				}

				// 15.ＧＯＴ
				// 16.ＧＰＴ
				// 17.γ－ＧＴＰ
				if inRecs[J][113] != "" { // 判定チェック
					cRec[15] = inRecs[J][16]
					cRec[16] = inRecs[J][17]
					cRec[17] = inRecs[J][18]
				} else {
					cRec[15] = ""
					cRec[16] = ""
					cRec[17] = ""
				}

				// 18.◆有機
				cRec[18] = inRecs[J][28]

				// 19.1.ｱｾﾄﾝ
				cRec[19] = inRecs[J][29]

				// 20.2.ｲｿﾌﾞﾁﾙｱﾙｺｰﾙ
				cRec[20] = inRecs[J][30]

				// 21.3.ｲｿﾌﾟﾛﾋﾟﾙｱﾙｺｰﾙ
				cRec[21] = inRecs[J][31]

				// 22.4.ｲｿﾍﾟﾝﾁﾙｱﾙｺｰﾙ
				cRec[22] = inRecs[J][32]

				// 23.5.ｴﾁﾙｴｰﾃﾙ
				cRec[23] = inRecs[J][33]

				// 24.6.ｴﾁﾚﾝｸﾞﾘｺｰﾙﾓﾉｴﾁﾙｴｰﾃﾙ
				cRec[24] = inRecs[J][34]

				// 25.7.ｴﾁﾚﾝｸﾞﾘｺｰﾙﾓﾉｴﾁﾙｴｰﾃﾙｱｾﾃｰﾄ
				cRec[25] = inRecs[J][35]

				// 26.8.ｴﾁﾚﾝｸﾞﾘｺｰﾙﾓﾉﾌﾞﾁﾙｴｰﾃﾙ
				cRec[26] = inRecs[J][36]

				// 27.9.ｴﾁﾚﾝｸﾞﾘｺｰﾙﾓﾉﾒﾁﾙｴｰﾃﾙ
				cRec[27] = inRecs[J][37]

				// 28.10.ｵﾙﾄ-ｼﾞｸﾛﾙﾍﾞﾝｾﾞﾝ
				cRec[28] = inRecs[J][38]

				// 29.11.ｷｼﾚﾝ
				cRec[29] = inRecs[J][39]

				// 30.12.ｸﾚｿﾞｰﾙ
				cRec[30] = inRecs[J][40]

				// 31.13.ｸﾛﾙﾍﾞﾝｾﾞﾝ
				cRec[31] = inRecs[J][41]

				// 32.15.酢酸ｲｿﾌﾞﾁﾙ
				cRec[32] = inRecs[J][42]

				// 33.16.酢酸ｲｿﾌﾟﾛﾋﾟﾙ
				cRec[33] = inRecs[J][43]

				// 34.17.酢酸ｲｿﾍﾟﾝﾁﾙ
				cRec[34] = inRecs[J][44]

				// 35.18.酢酸ｴﾁﾙ
				cRec[35] = inRecs[J][45]

				// 36.19.酢酸ﾌﾞﾁﾙ
				cRec[36] = inRecs[J][46]

				// 37.20.酢酸ﾌﾟﾛﾋﾟﾙ
				cRec[37] = inRecs[J][47]

				// 38.21.酢酸ﾍﾟﾝﾁﾙ
				cRec[38] = inRecs[J][48]

				// 39.22.酢酸ﾒﾁﾙ
				cRec[39] = inRecs[J][49]

				// 40.24.ｼｸﾛﾍｷｻﾉｰﾙ
				cRec[40] = inRecs[J][50]

				// 41.25.ｼｸﾛﾍｷｻﾉﾝ
				cRec[41] = inRecs[J][51]

				// 42.30.N,N-ｼﾞﾒﾁﾙﾎﾙﾑｱﾐﾄﾞ
				cRec[42] = inRecs[J][52]

				// 43.34.ﾃﾄﾗﾋﾄﾞﾛﾌﾗﾝ
				cRec[43] = inRecs[J][53]

				// 44.35.1,1,1-ﾄﾘｸﾛﾙｴﾀﾝ
				cRec[44] = inRecs[J][54]

				// 45.37.ﾄﾙｴﾝ
				cRec[45] = inRecs[J][55]

				// 46.39.ﾉﾙﾏﾙﾍｷｻﾝ
				cRec[46] = inRecs[J][56]

				// 47.40.1-ﾌﾞﾀﾉｰﾙ
				cRec[47] = inRecs[J][57]

				// 48.41.2-ﾌﾞﾀﾉｰﾙ
				cRec[48] = inRecs[J][58]

				// 49.42.ﾒﾀﾉｰﾙ
				cRec[49] = inRecs[J][59]

				// 50.44.ﾒﾁﾙｴﾁﾙｹﾄﾝ
				cRec[50] = inRecs[J][60]

				// 51.45.ﾒﾁﾙｼｸﾛﾍｷｻﾉｰﾙ
				cRec[51] = inRecs[J][61]

				// 52.46.ﾒﾁﾙｼｸﾛﾍｷｻﾉﾝ
				cRec[52] = inRecs[J][62]

				// 53.47.ﾒﾁﾙ-ﾉﾙﾏﾙ-ﾌﾞﾁﾙｹﾄﾝ
				cRec[53] = inRecs[J][63]

				// 54.48.ｶﾞｿﾘﾝ
				cRec[54] = inRecs[J][64]

				// 55.49.ｺｰﾙﾀｰﾙﾅﾌｻ
				cRec[55] = inRecs[J][65]

				// 56.50.石油ｴｰﾃﾙ
				cRec[56] = inRecs[J][66]

				// 57.51.石油ﾅﾌｻ
				cRec[57] = inRecs[J][67]

				// 58.52.石油ﾍﾞﾝｼﾞﾝ
				cRec[58] = inRecs[J][68]

				// 59.53.ﾃﾚﾋﾞﾝ油
				cRec[59] = inRecs[J][69]

				// 60.54.ﾐﾈﾗﾙｽﾋﾟﾘｯﾄ
				cRec[60] = inRecs[J][70]

				// 61.馬尿酸
				cRec[61] = inRecs[J][71]

				// 62.馬尿酸_分布
				cRec[62] = inRecs[J][72]

				// 63.メチル馬尿酸
				cRec[63] = inRecs[J][73]

				// 64.メチル馬尿酸_分布
				cRec[64] = inRecs[J][74]

				// 65.2.5-ﾍｷｻﾝｼﾞｵﾝ
				cRec[65] = inRecs[J][75]

				// 66.2.5-ﾍｷｻﾝｼﾞｵﾝ_分布
				cRec[66] = inRecs[J][76]

				// 67.有機_健診種別
				cRec[67] = inRecs[J][77]

				// 68.有機_作業名１
				cRec[68] = inRecs[J][78]

				// 69.有機_作業名２
				cRec[69] = inRecs[J][79]

				// 70.有機_業務名１
				cRec[70] = inRecs[J][80]

				// 71.有機_業務名２
				cRec[71] = inRecs[J][81]

				// 72.有機_業務名３
				cRec[72] = inRecs[J][82]

				// 73.有機_従事年数
				cRec[73] = inRecs[J][83]

				// 74.有機_従事年数_月
				cRec[74] = inRecs[J][84]

				// 75.有機_作業時間
				cRec[75] = inRecs[J][85]

				// 76.有機_作業時間_分
				cRec[76] = inRecs[J][86]

				// 77.有機_従事日数
				cRec[77] = inRecs[J][87]

				// 78.有機_従事日数_週月
				cRec[78] = inRecs[J][88]

				// 79.有機_作業工程に変化
				cRec[79] = inRecs[J][89]

				// 80.有機_局所排気装置
				cRec[80] = inRecs[J][90]

				// 81.有機_防毒マスク
				cRec[81] = inRecs[J][91]

				// 82.有機_不透過性手袋
				cRec[82] = inRecs[J][92]

				// 83.有機_取扱量・使用頻度
				cRec[83] = inRecs[J][93]

				// 84.有機_大量のばく露
				cRec[84] = inRecs[J][94]

				// 85.有機_直接触れる作業
				cRec[85] = inRecs[J][95]

				// 86.有機_自覚症状１
				cRec[86] = inRecs[J][96]

				// 87.有機_自覚症状２
				cRec[87] = inRecs[J][97]

				// 88.有機_自覚症状３
				cRec[88] = inRecs[J][98]

				// 89.有機_自覚症状４
				cRec[89] = inRecs[J][99]

				// 90.有機_自覚症状５
				cRec[90] = inRecs[J][100]

				// 91.有機_既往歴１
				cRec[91] = inRecs[J][101]

				// 92.有機_既往歴２
				cRec[92] = inRecs[J][102]

				// 93.有機_既往歴３
				cRec[93] = inRecs[J][103]

				// 94.有機_診察所見１
				cRec[94] = inRecs[J][104]

				// 95.有機_診察所見２
				cRec[95] = inRecs[J][105]

				// 96.有機_診察所見３
				cRec[96] = inRecs[J][106]

				// 97.有機_診察判定
				cRec[97] = Hantei(inRecs[J][107])
				if Hantei(inRecs[J][107]) == "err" {
					log.Print("診察所見判定にエラーがあります。\r\n")
				}
				// 98.有機_診察判定コメント
				cRec[98] = inRecs[J][108]

				// 99.有機_尿蛋白判定
				cRec[99] = Hantei(inRecs[J][109])
				if Hantei(inRecs[J][109]) == "err" {
					log.Print("尿蛋白判定にエラーがあります。\r\n")
				}

				// 100.有機_尿蛋白判定コメント
				cRec[100] = inRecs[J][110]

				// 101.有機_貧血判定
				cRec[101] = Hantei(inRecs[J][111])
				if Hantei(inRecs[J][111]) == "err" {
					log.Print("貧血判定にエラーがあります。\r\n")
				}

				// 102.有機_貧血判定コメント
				cRec[102] = inRecs[J][112]

				// 103.有機_肝機能判定
				cRec[103] = Hantei(inRecs[J][113])
				if Hantei(inRecs[J][113]) == "err" {
					log.Print("肝機能判定にエラーがあります。\r\n")
				}

				// 104.有機_肝機能コメント
				cRec[104] = inRecs[J][114]

				// 105.有機_馬尿酸判定
				cRec[105] = inRecs[J][115]

				// 106.有機_馬尿酸判定コメント
				cRec[106] = inRecs[J][116]

				// 107.有機_ﾒﾁﾙ馬尿酸判定
				cRec[107] = inRecs[J][117]

				// 108.有機_ﾒﾁﾙ馬尿酸判定コメント
				cRec[108] = inRecs[J][118]

				// 109.有機_2.5ﾍｷｻﾝｼﾞｵﾝ判定
				cRec[109] = inRecs[J][119]

				// 110.有機_2.5ﾍｷｻﾝｼﾞｵﾝ判定コメント
				cRec[110] = inRecs[J][120]

				// 111.有機_管理区分
				cRec[111] = inRecs[J][121]

				// 112.医療機関判定（有機）
				sogo := ""
				var h [7][2]string
				h[0][0] = Hantei(inRecs[J][107]) //診察所見判定
				h[0][1] = inRecs[J][108]         //診察所見所見
				h[1][0] = Hantei(inRecs[J][109]) //尿蛋白判定
				h[1][1] = inRecs[J][110]         //尿蛋白所見
				h[2][0] = Hantei(inRecs[J][111]) //貧血判定
				h[2][1] = inRecs[J][112]         //貧血所見
				h[3][0] = Hantei(inRecs[J][113]) //肝機能判定
				h[3][1] = inRecs[J][114]         //肝機能所見
				h[4][0] = inRecs[J][115]         //馬尿酸判定
				h[4][1] = inRecs[J][116]         //馬尿酸所見
				h[5][0] = inRecs[J][117]         //ﾒﾁﾙ馬尿酸判定
				h[5][1] = inRecs[J][118]         //ﾒﾁﾙ馬尿酸所見
				h[6][0] = inRecs[J][119]         //2.5ﾍｷｻﾝｼﾞｵﾝ判定
				h[6][1] = inRecs[J][120]         //2.5ﾍｷｻﾝｼﾞｵﾝ所見

				hKigo := [...]string{"Ｆ", "Ｅ", "３", "Ｄ", "２", "Ｇ", "Ｃ"}
				for k := 0; k < 7; k++ {
					for l := 0; l < 7; l++ {
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

				cRec[112] = sogo

				// 113.医療機関名称
				cRec[113] = "医療法人社団　松英会"

				// 114.健康診断を実施した医師の氏名
				cRec[114] = "寺門　節雄"

				// 115.従事年数
				jyujiYear := ""
				if inRecs[J][83] != "" {
					jyujiYear = inRecs[J][83] + "年"
				}

				jyujiMon := ""
				if inRecs[J][84] != "" {
					jyujiMon = inRecs[J][84] + "ヵ月"
				}

				if jyujiYear != "" && jyujiMon != "" {
					cRec[115] = jyujiYear + " " + jyujiMon
				} else {
					cRec[115] = jyujiYear + jyujiMon
				}

				// 116.健診対象有機溶剤
				yozai := ""
				for k := 0; k < 42; k++ {
					if inRecs[J][29+k] != "" {
						if yozai == "" {
							yozai = inRecs[J][29+k]
						} else {
							yozai = yozai + "、" + inRecs[J][29+k]
						}
					}
				}

				cRec[116] = yozai

				// 117.有機溶剤業務名
				gyomu := ""
				for k := 0; k < 3; k++ {
					if inRecs[J][80+k] != "" {
						if gyomu == "" {
							gyomu = inRecs[J][80+k]
						} else {
							gyomu = gyomu + "、" + inRecs[J][80+k]
						}
					}
				}

				cRec[117] = gyomu

				// 118.1日の作業時間（有機）
				sagyoHour := ""
				if inRecs[J][85] != "" {
					sagyoHour = inRecs[J][85] + "時間"
				}

				sagyoMin := ""
				if inRecs[J][86] != "" {
					sagyoMin = inRecs[J][86] + "分"
				}

				if sagyoHour != "" && sagyoMin != "" {
					cRec[118] = sagyoHour + " " + sagyoMin
				} else {
					cRec[118] = sagyoHour + sagyoMin
				}

				// 119.従事日数_週月（有機）
				jyuji := ""
				if inRecs[J][87] != "" {
					jyuji = inRecs[J][87] + "日"
				}

				jyujiWM := ""
				if inRecs[J][88] != "" {
					jyujiWM = "/" + inRecs[J][88]
				}

				/*
					if jyuji != "" && jyujiWM != "" {
						cRec[119] = jyuji + " " + jyujiWM
					} else {
						cRec[119] = jyuji + jyujiWM
					}
				*/
				cRec[119] = jyuji + jyujiWM

				// 120.有機_診察判定
				cRec[120] = HanteiCode(inRecs[J][107])
				if HanteiCode(inRecs[J][107]) == "err" {
					log.Print("診察判定にエラーがあります\r\n")
				}

				// 120.有機_診察判定結果
				cRec[121] = HanteiCode(inRecs[J][107])
				if HanteiCode(inRecs[J][107]) == "err" {
					log.Print("診察結果にエラーがあります\r\n")
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
