package main

import (
	"log"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

func ConversionChloroform(filename string, inRecs [][]string, coRecs [][]string, dateRecs [][]string) {
	// クロロホルム他９物質データ変換
	var vcell *xlsx.Cell
	var r int
	var cell string

	recLen := 96 //出力するレコードの項目数
	cRec := make([]string, recLen)
	var I int

	day := time.Now()

	//会社毎に健診データファイルを作成する
	for _, coRec := range coRecs {

		excelName := filename + coRec[1] + "ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK健診データ" + day.Format("20060102") + ".xlsx"
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
		cRec[13] = "knk_kenkork_kensa.kensa_val_027"
		cRec[14] = "knk_kenkork_kensa.kensa_val_012"
		cRec[15] = "knk_kenkork_kensa.kensa_val_013"
		cRec[16] = "knk_kenkork_kensa.kensa_val_014"
		cRec[17] = "knk_kenkork_kensa.kensa_val_015"
		cRec[18] = "knk_kenkork_kensa.kensa_val_016"
		cRec[19] = "knk_kenkork_kensa.kensa_val_018"
		cRec[20] = "knk_kenkork_kensa.kensa_val_020"
		cRec[21] = "knk_kenkork_kensa.kensa_val_021"
		cRec[22] = "knk_kenkork_kensa.kensa_val_022"
		cRec[23] = "knk_kenkork_kensa.kensa_val_023"
		cRec[24] = "knk_kenkork_kensa.kensa_val_024"
		cRec[25] = "knk_kenkork_kensa.kensa_val_025"
		cRec[26] = "knk_kenkork_kensa.kensa_val_026"
		cRec[34] = "knk_kenkork_kensa.kensa_val_029"
		cRec[35] = "knk_kenkork_kensa.kensa_val_030"
		cRec[36] = "knk_kenkork_kensa.kensa_val_031"
		cRec[37] = "knk_kenkork_kensa.kensa_val_003"
		cRec[38] = "knk_kenkork_kensa.kensa_val_004"
		cRec[43] = "knk_kenkork_kensa.kensa_val_051"
		cRec[44] = "knk_kenkork_kensa.kensa_val_052"
		cRec[45] = "knk_kenkork_kensa.kensa_val_053"
		cRec[46] = "knk_kenkork_kensa.kensa_val_054"
		cRec[47] = "knk_kenkork_kensa.kensa_val_055"
		cRec[48] = "knk_kenkork_kensa.kensa_val_056"
		cRec[49] = "knk_kenkork_kensa.kensa_val_057"
		cRec[50] = "knk_kenkork_kensa.kensa_val_033"
		cRec[51] = "knk_kenkork_kensa.kensa_val_034"
		cRec[53] = "knk_kenkork_kensa.kensa_val_032"
		cRec[56] = "knk_kenkork_kensa.kensa_val_035"
		cRec[59] = "knk_kenkork_kensa.kensa_val_062"
		cRec[67] = "knk_kenkork_kensa.kensa_val_019"
		cRec[69] = "knk_kenkork_kensa.kensa_val_036"
		cRec[71] = "knk_kenkork_kensa.kensa_val_037"
		cRec[73] = "knk_kenkork_kensa.kensa_val_038"
		cRec[76] = "knk_kenkork_kensa.kensa_val_040"
		cRec[77] = "knk_kenkork_kensa.kensa_val_041"
		cRec[78] = "knk_kenkork_kensa.kensa_val_042"
		cRec[79] = "knk_kenkork_kensa.kensa_val_001"
		cRec[80] = "knk_kenkork_kensa.kensa_val_002"
		cRec[81] = "knk_kenkork_kensa.kensa_val_005"
		cRec[82] = "knk_kenkork_kensa.kensa_val_006"
		cRec[83] = "knk_kenkork_kensa.kensa_val_007"
		cRec[84] = "knk_kenkork_kensa.kensa_val_008"
		cRec[85] = "knk_kenkork_kensa.kensa_val_009"
		cRec[86] = "knk_kenkork_kensa.kensa_val_010"
		cRec[87] = "knk_kenkork_kensa.kensa_val_049"
		cRec[88] = "knk_kenkork_kensa.kensa_val_050"
		cRec[89] = "knk_kenkork_kensa.hantei_val_019"
		cRec[90] = "knk_kenkork_kensa.hantei_val_062"
		cRec[91] = "knk_kenkork_kensa.hantei_val_036"
		cRec[92] = "knk_kenkork_kensa.hantei_val_037"
		cRec[93] = "knk_kenkork_kensa.hantei_val_038"
		cRec[94] = "knk_kenkork_kensa.kensa_val_039"
		cRec[95] = "knk_kenkork_kensa.hantei_val_039"

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
		cRec[13] = "検査コード027_医療機関側検査値"
		cRec[14] = "検査コード012_医療機関側検査値"
		cRec[15] = "検査コード013_医療機関側検査値"
		cRec[16] = "検査コード014_医療機関側検査値"
		cRec[17] = "検査コード015_医療機関側検査値"
		cRec[18] = "検査コード016_医療機関側検査値"
		cRec[19] = "検査コード018_医療機関側検査値"
		cRec[20] = "検査コード020_医療機関側検査値"
		cRec[21] = "検査コード021_医療機関側検査値"
		cRec[22] = "検査コード022_医療機関側検査値"
		cRec[23] = "検査コード023_医療機関側検査値"
		cRec[24] = "検査コード024_医療機関側検査値"
		cRec[25] = "検査コード025_医療機関側検査値"
		cRec[26] = "検査コード026_医療機関側検査値"
		cRec[34] = "検査コード029_医療機関側検査値"
		cRec[35] = "検査コード030_医療機関側検査値"
		cRec[36] = "検査コード031_医療機関側検査値"
		cRec[37] = "検査コード003_医療機関側検査値"
		cRec[38] = "検査コード004_医療機関側検査値"
		cRec[43] = "検査コード051_医療機関側検査値"
		cRec[44] = "検査コード052_医療機関側検査値"
		cRec[45] = "検査コード053_医療機関側検査値"
		cRec[46] = "検査コード054_医療機関側検査値"
		cRec[47] = "検査コード055_医療機関側検査値"
		cRec[48] = "検査コード056_医療機関側検査値"
		cRec[49] = "検査コード057_医療機関側検査値"
		cRec[50] = "検査コード033_医療機関側検査値"
		cRec[51] = "検査コード034_医療機関側検査値"
		cRec[53] = "検査コード032_医療機関側検査値"
		cRec[56] = "検査コード035_医療機関側検査値"
		cRec[59] = "検査コード062_医療機関側検査値"
		cRec[67] = "検査コード019_医療機関側検査値"
		cRec[69] = "検査コード036_医療機関側検査値"
		cRec[71] = "検査コード037_医療機関側検査値"
		cRec[73] = "検査コード038_医療機関側検査値"
		cRec[76] = "検査コード040_医療機関側検査値"
		cRec[77] = "検査コード041_医療機関側検査値"
		cRec[78] = "検査コード042_医療機関側検査値"
		cRec[79] = "検査コード001_医療機関側検査値"
		cRec[80] = "検査コード002_医療機関側検査値"
		cRec[81] = "検査コード005_医療機関側検査値"
		cRec[82] = "検査コード006_医療機関側検査値"
		cRec[83] = "検査コード007_医療機関側検査値"
		cRec[84] = "検査コード008_医療機関側検査値"
		cRec[85] = "検査コード009_医療機関側検査値"
		cRec[86] = "検査コード010_医療機関側検査値"
		cRec[87] = "検査コード049_医療機関側検査値"
		cRec[88] = "検査コード050_医療機関側検査値"
		cRec[89] = "検査コード019_医療機関側判定結果"
		cRec[90] = "検査コード062_医療機関側判定結果"
		cRec[91] = "検査コード036_医療機関側判定結果"
		cRec[92] = "検査コード037_医療機関側判定結果"
		cRec[93] = "検査コード038_医療機関側判定結果"
		cRec[94] = "検査コード039_医療機関側判定結果"
		cRec[95] = "検査コード039_医療機関側判定結果"

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
		cRec[12] = "◆クロロホルムほか９物質"
		cRec[13] = "尿蛋白定性"
		cRec[14] = "総ビリルビン"
		cRec[15] = "ＧＯＴ"
		cRec[16] = "ＧＰＴ"
		cRec[17] = "γ－ＧＴＰ"
		cRec[18] = "ＡＬＰ"
		cRec[19] = "白血球数"
		cRec[20] = "白血球像_BASO"
		cRec[21] = "白血球像_EOSINO"
		cRec[22] = "白血球像_Stab"
		cRec[23] = "白血球像_Seg"
		cRec[24] = "白血球像_Lympho"
		cRec[25] = "白血球像_Mono"
		cRec[26] = "白血球像_その他"
		cRec[27] = "ｼﾞｸﾛﾛﾒﾀﾝ従事年"
		cRec[28] = "ｼﾞｸﾛﾛﾒﾀﾝ従事月"
		cRec[29] = "ｽﾁﾚﾝ従事年"
		cRec[30] = "ｽﾁﾚﾝ従事月"
		cRec[31] = "ﾒﾁﾙｲｿﾌﾞﾁﾙｹﾄﾝ従事年"
		cRec[32] = "ﾒﾁﾙｲｿﾌﾞﾁﾙｹﾄﾝ従事月"
		cRec[33] = "ｽﾁﾚﾝ_マンデル酸"
		cRec[34] = "ｽﾁﾚﾝ_PGA＋ＭＡ"
		cRec[35] = "ｽﾁﾚﾝ_マンデル酸_旧"
		cRec[36] = "ｽﾁﾚﾝ_マンデル酸分布_旧"
		cRec[37] = "ｸﾛﾛﾎﾙﾑ他_作業名１"
		cRec[38] = "ｸﾛﾛﾎﾙﾑ他_作業名２"
		cRec[39] = "ｸﾛﾛﾎﾙﾑ他_作業時間"
		cRec[40] = "ｸﾛﾛﾎﾙﾑ他_作業時間_分"
		cRec[41] = "ｸﾛﾛﾎﾙﾑ他_作業日数"
		cRec[42] = "ｸﾛﾛﾎﾙﾑ他_作業日数_週月"
		cRec[43] = "ｸﾛﾛﾎﾙﾑ他_作業工程に変化"
		cRec[44] = "ｸﾛﾛﾎﾙﾑ他_局所排気装置"
		cRec[45] = "ｸﾛﾛﾎﾙﾑ他_防毒マスク"
		cRec[46] = "ｸﾛﾛﾎﾙﾑ他_不透過性手袋"
		cRec[47] = "ｸﾛﾛﾎﾙﾑ他_取扱量・使用頻度"
		cRec[48] = "ｸﾛﾛﾎﾙﾑ他_大量のばく露"
		cRec[49] = "ｸﾛﾛﾎﾙﾑ他_直接触れる作業"
		cRec[50] = "ｸﾛﾛﾎﾙﾑ他_自覚症状１"
		cRec[51] = "ｸﾛﾛﾎﾙﾑ他_自覚症状２"
		cRec[52] = "ｸﾛﾛﾎﾙﾑ他_自覚症状３"
		cRec[53] = "ｸﾛﾛﾎﾙﾑ他_既往歴１"
		cRec[54] = "ｸﾛﾛﾎﾙﾑ他_既往歴２"
		cRec[55] = "ｸﾛﾛﾎﾙﾑ他_既往歴３"
		cRec[56] = "ｸﾛﾛﾎﾙﾑ他_診察所見１"
		cRec[57] = "ｸﾛﾛﾎﾙﾑ他_診察所見２"
		cRec[58] = "ｸﾛﾛﾎﾙﾑ他_診察所見３"
		cRec[59] = "ｸﾛﾛﾎﾙﾑ他_診察判定"
		cRec[60] = "ｸﾛﾛﾎﾙﾑ他_診察判定コメント"
		cRec[61] = "ｸﾛﾛﾎﾙﾑ他_尿蛋白判定"
		cRec[62] = "ｸﾛﾛﾎﾙﾑ他_尿蛋白判定コメント"
		cRec[63] = "ｸﾛﾛﾎﾙﾑ他_肝機能判定"
		cRec[64] = "ｸﾛﾛﾎﾙﾑ他_肝機能判定コメント"
		cRec[65] = "ｸﾛﾛﾎﾙﾑ他_白血球判定"
		cRec[66] = "ｸﾛﾛﾎﾙﾑ他_白血球判定コメント"
		cRec[67] = "ｸﾛﾛﾎﾙﾑ他_白血球像判定"
		cRec[68] = "ｸﾛﾛﾎﾙﾑ他_白血球像判定コメント"
		cRec[69] = "ジクロロメタン判定"
		cRec[70] = "ジクロロメタン判定コメント"
		cRec[71] = "スチレン判定"
		cRec[72] = "スチレン判定コメント"
		cRec[73] = "メチルイソブチルケトン判定"
		cRec[74] = "メチルイソブチルケトン判定コメント"
		cRec[75] = "ｸﾛﾛﾎﾙﾑ他_管理区分"
		cRec[76] = "医療機関判定（ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK）"
		cRec[77] = "医療機関名称"
		cRec[78] = "健康診断を実施した医師の氏名"
		cRec[79] = "特定化学物質業務名"
		cRec[80] = "健診種別"
		cRec[81] = "ｼﾞｸﾛﾛﾒﾀﾝ_従事年数"
		cRec[82] = "ｼﾞｸﾛﾛﾒﾀﾝ_取扱い終了時期"
		cRec[83] = "ｽﾁﾚﾝ_従事年数"
		cRec[84] = "ｽﾁﾚﾝ_取扱い終了時期"
		cRec[85] = "MIBK_従事年数"
		cRec[86] = "MIBK_取扱い終了時期"
		cRec[87] = "ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK_作業時間"
		cRec[88] = "ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK_従事日数_週月"
		cRec[89] = "ｸﾛﾛﾎﾙﾑ他_白血球像判定"
		cRec[90] = "ｸﾛﾛﾎﾙﾑ他_診察判定"
		cRec[91] = "ジクロロメタン判定"
		cRec[92] = "スチレン判定"
		cRec[93] = "メチルイソブチルケトン判定"
		cRec[94] = "医療機関判定（ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK）"
		cRec[95] = "医療機関判定（ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK）"

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

			if inRecs[J][0] == coRec[0] && inRecs[J][202] == "●" {
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

				// 12.◆クロロホルムほか９物質
				cRec[12] = inRecs[J][202]

				// 13.尿蛋白定性
				if inRecs[J][237] != "" { //判定チェック
					cRec[13] = inRecs[J][12]
				} else {
					cRec[13] = ""
				}

				// 14.総ビリルビン
				// 18.ＡＬＰ
				if inRecs[J][245] != "" { //判定チェック
					cRec[14] = inRecs[J][15]
					cRec[18] = inRecs[J][19]
				} else {
					cRec[14] = ""
					cRec[18] = ""
				}

				// 15.ＧＯＴ
				// 16.ＧＰＴ
				// 17.γ－ＧＴＰ
				// if inRecs[J][245] != "" || inRecs[J][247] != "" { //判定チェック
				if inRecs[J][245] != "" { //判定チェック
					cRec[15] = inRecs[J][16]
					cRec[16] = inRecs[J][17]
					cRec[17] = inRecs[J][18]
				} else {
					cRec[15] = ""
					cRec[16] = ""
					cRec[17] = ""
				}

				// 19.白血球数
				//if inRecs[J][247] != "" { //判定チェック
				//	cRec[19] = inRecs[J][20]
				//} else {
				cRec[19] = ""
				//}

				// 20.白血球像_BASO
				cRec[20] = inRecs[J][21]

				// 21.白血球像_EOSINO
				cRec[21] = inRecs[J][22]

				// 22.白血球像_Stab
				cRec[22] = inRecs[J][23]

				// 23.白血球像_Seg
				cRec[23] = inRecs[J][24]

				// 24.白血球像_Lympho
				cRec[24] = inRecs[J][25]

				// 25.白血球像_Mono
				cRec[25] = inRecs[J][26]

				// 26.白血球像_その他
				cRec[26] = inRecs[J][27]

				// 27.ｼﾞｸﾛﾛﾒﾀﾝ従事年
				cRec[27] = inRecs[J][203]

				// 28.ｼﾞｸﾛﾛﾒﾀﾝ従事月
				cRec[28] = inRecs[J][204]

				// 29.ｽﾁﾚﾝ従事年
				cRec[29] = inRecs[J][205]

				// 30.ｽﾁﾚﾝ従事月
				cRec[30] = inRecs[J][206]

				// 31.ﾒﾁﾙｲｿﾌﾞﾁﾙｹﾄﾝ従事年
				cRec[31] = inRecs[J][207]

				// 32.ﾒﾁﾙｲｿﾌﾞﾁﾙｹﾄﾝ従事月
				cRec[32] = inRecs[J][208]

				// 33.ｽﾁﾚﾝ_マンデル酸
				cRec[33] = inRecs[J][209]

				// 34.ｽﾁﾚﾝ_PGA＋ＭＡ
				cRec[34] = inRecs[J][210]

				// 35.ｽﾁﾚﾝ_マンデル酸_旧
				cRec[35] = inRecs[J][211]

				// 36.ｽﾁﾚﾝ_マンデル酸分布_旧
				cRec[36] = inRecs[J][212]

				// 37.ｸﾛﾛﾎﾙﾑ他_作業名１
				cRec[37] = inRecs[J][213]

				// 38.ｸﾛﾛﾎﾙﾑ他_作業名２
				cRec[38] = inRecs[J][214]

				// 39.ｸﾛﾛﾎﾙﾑ他_作業時間
				cRec[39] = inRecs[J][215]

				// 40.ｸﾛﾛﾎﾙﾑ他_作業時間_分
				cRec[40] = inRecs[J][216]

				// 41.ｸﾛﾛﾎﾙﾑ他_作業日数
				cRec[41] = inRecs[J][217]

				// 42.ｸﾛﾛﾎﾙﾑ他_作業日数_週月
				cRec[42] = inRecs[J][218]

				// 43.ｸﾛﾛﾎﾙﾑ他_作業工程に変化
				cRec[43] = inRecs[J][219]

				// 44.ｸﾛﾛﾎﾙﾑ他_局所排気装置
				cRec[44] = inRecs[J][220]

				// 45.ｸﾛﾛﾎﾙﾑ他_防毒マスク
				cRec[45] = inRecs[J][221]

				// 46.ｸﾛﾛﾎﾙﾑ他_不透過性手袋
				cRec[46] = inRecs[J][222]

				// 47.ｸﾛﾛﾎﾙﾑ他_取扱量・使用頻度
				cRec[47] = inRecs[J][223]

				// 48.ｸﾛﾛﾎﾙﾑ他_大量のばく露
				cRec[48] = inRecs[J][224]

				// 49.ｸﾛﾛﾎﾙﾑ他_直接触れる作業
				cRec[49] = inRecs[J][225]

				// 50.ｸﾛﾛﾎﾙﾑ他_自覚症状１
				cRec[50] = inRecs[J][226]

				// 51.ｸﾛﾛﾎﾙﾑ他_自覚症状２
				cRec[51] = inRecs[J][227]

				// 52.ｸﾛﾛﾎﾙﾑ他_自覚症状３
				cRec[52] = inRecs[J][228]

				// 53.ｸﾛﾛﾎﾙﾑ他_既往歴１
				cRec[53] = inRecs[J][229]

				// 54.ｸﾛﾛﾎﾙﾑ他_既往歴２
				cRec[54] = inRecs[J][230]

				// 55.ｸﾛﾛﾎﾙﾑ他_既往歴３
				cRec[55] = inRecs[J][231]

				// 56.ｸﾛﾛﾎﾙﾑ他_診察所見１
				cRec[56] = inRecs[J][232]

				// 57.ｸﾛﾛﾎﾙﾑ他_診察所見２
				cRec[57] = inRecs[J][233]

				// 58.ｸﾛﾛﾎﾙﾑ他_診察所見３
				cRec[58] = inRecs[J][234]

				// 59.ｸﾛﾛﾎﾙﾑ他_診察判定
				cRec[59] = Hantei(inRecs[J][235])
				if Hantei(inRecs[J][235]) == "err" {
					log.Print("診察所見判定にエラーがあります。\r\n")
				}

				// 60.ｸﾛﾛﾎﾙﾑ他_診察判定コメント
				cRec[60] = inRecs[J][236]

				// 61.ｸﾛﾛﾎﾙﾑ他_尿蛋白判定
				cRec[61] = Hantei(inRecs[J][237])
				if Hantei(inRecs[J][237]) == "err" {
					log.Print("尿蛋白判定にエラーがあります。\r\n")
				}

				// 62.ｸﾛﾛﾎﾙﾑ他_尿蛋白判定コメント
				cRec[62] = inRecs[J][238]

				// 63.ｸﾛﾛﾎﾙﾑ他_肝機能判定
				cRec[63] = Hantei(inRecs[J][239])
				if Hantei(inRecs[J][239]) == "err" {
					log.Print("肝機能判定にエラーがあります。\r\n")
				}

				// 64.ｸﾛﾛﾎﾙﾑ他_肝機能判定コメント
				cRec[64] = inRecs[J][240]

				// 65.ｸﾛﾛﾎﾙﾑ他_白血球判定
				cRec[65] = Hantei(inRecs[J][241])
				if Hantei(inRecs[J][241]) == "err" {
					log.Print("白血球判定にエラーがあります。\r\n")
				}

				// 66.ｸﾛﾛﾎﾙﾑ他_白血球判定コメント
				cRec[66] = inRecs[J][242]

				// 67.ｸﾛﾛﾎﾙﾑ他_白血球像判定
				cRec[67] = Hantei(inRecs[J][243])
				if Hantei(inRecs[J][243]) == "err" {
					log.Print("白血球像判定にエラーがあります。\r\n")
				}

				// 68.ｸﾛﾛﾎﾙﾑ他_白血球像判定コメント
				cRec[68] = inRecs[J][244]

				// 69.ジクロロメタン判定
				cRec[69] = inRecs[J][245]

				// 70.ジクロロメタン判定コメント
				cRec[70] = inRecs[J][246]

				// 71.スチレン判定
				cRec[71] = inRecs[J][247]

				// 72.スチレン判定コメント
				cRec[72] = inRecs[J][248]

				// 73.メチルイソブチルケトン判定
				cRec[73] = inRecs[J][249]

				// 74.メチルイソブチルケトン判定コメント
				cRec[74] = inRecs[J][250]

				// 75.ｸﾛﾛﾎﾙﾑ他_管理区分
				cRec[75] = inRecs[J][251]

				// 76.医療機関判定（ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK）
				sogo := ""
				var h [8][2]string
				h[0][0] = Hantei(inRecs[J][235]) //診察所見判定
				h[0][1] = inRecs[J][236]         //診察所見所見
				h[1][0] = Hantei(inRecs[J][237]) //尿蛋白判定
				h[1][1] = inRecs[J][238]         //尿蛋白所見
				h[2][0] = Hantei(inRecs[J][239]) //肝機能判定
				h[2][1] = inRecs[J][240]         //肝機能所見
				h[3][0] = Hantei(inRecs[J][241]) //白血球数判定
				h[3][1] = inRecs[J][242]         //白血球数所見
				h[4][0] = Hantei(inRecs[J][243]) //白血球像判定
				h[4][1] = inRecs[J][244]         //白血球像所見
				h[5][0] = inRecs[J][245]         //ジクロロメタン判定
				h[5][1] = inRecs[J][246]         //ジクロロメタン所見
				h[6][0] = inRecs[J][247]         //スチレン判定
				h[6][1] = inRecs[J][248]         //スチレン所見
				h[7][0] = inRecs[J][249]         //ﾒﾁﾙｲｿﾌﾞﾁﾙｹﾄﾝ判定
				h[7][1] = inRecs[J][250]         //ﾒﾁﾙｲｿﾌﾞﾁﾙｹﾄﾝ所見

				hKigo := [...]string{"Ｆ", "Ｅ", "３", "Ｄ", "２", "Ｇ", "Ｃ"}
				for k := 0; k < 7; k++ {
					for l := 0; l < 8; l++ {
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

				cRec[76] = sogo

				// 77.医療機関名称
				cRec[77] = "医療法人社団　松英会"

				// 78.健康診断を実施した医師の氏名
				cRec[78] = "寺門　節雄"

				// 79.特定化学物質業務名
				gName := "ｸﾛﾛﾎﾙﾑほか9物質取扱い業務"
				/*
					if inRecs[J][245] != "" {
						gName = "ｼﾞｸﾛﾛﾒﾀﾝ取扱業務"
					}

					if inRecs[J][247] != "" {
						if gName != "" {
							gName = gName + "、ｽﾁﾚﾝ取扱業務"
						} else {
							gName = "ｽﾁﾚﾝ取扱業務"
						}
					}

					if inRecs[J][249] != "" {
						if gName != "" {
							gName = gName + "、MIBK取扱業務"
						} else {
							gName = "MIBK取扱業務"
						}
					}
				*/

				cRec[79] = gName

				// 80.健診種別
				cRec[80] = "定期"

				// 81.ｼﾞｸﾛﾛﾒﾀﾝ_従事年数
				jyujiYearJ := ""
				if inRecs[J][203] != "" {
					jyujiYearJ = inRecs[J][203] + "年"
				}

				jyujiMonJ := ""
				if inRecs[J][204] != "" {
					jyujiMonJ = inRecs[J][204] + "ヵ月"
				}

				if jyujiYearJ != "" && jyujiMonJ != "" {
					cRec[81] = jyujiYearJ + " " + jyujiMonJ
				} else {
					cRec[81] = jyujiYearJ + jyujiMonJ
				}

				// 82.ｼﾞｸﾛﾛﾒﾀﾝ_取扱い終了時期
				if dateRecs[0][6] != "ジクロロメタン取扱い終了年月日" {
					log.Print("「ジクロロメタン取扱い終了年月日」が見つかりませんでした。")
					failOnError(doError())
				}

				findflag := false
				for l, _ := range dateRecs {
					if cRec[4] == dateRecs[l][0] {
						if cRec[8] != dateRecs[l][3] {
							log.Printf("ジクロロメタン生年月日の不一致: %v %v %v != %v\r\n", cRec[4], cRec[6], cRec[8], dateRecs[l][3])
						}

						edate := dateRecs[l][6]
						if edate != "" {
							cRec[82] = edate[0:4] + "年" + edate[5:7] + "月" + edate[8:] + "日"
						} else {
							cRec[82] = edate
						}
						findflag = true
						break
					}
				}

				if findflag == false {
					log.Printf("取扱い終了名簿に対象がいません。 %v %v", cRec[4], cRec[6])
					cRec[82] = "err"
				}

				// 83.ｽﾁﾚﾝ_従事年数
				jyujiYearS := ""
				if inRecs[J][205] != "" {
					jyujiYearS = inRecs[J][205] + "年"
				}

				jyujiMonS := ""
				if inRecs[J][206] != "" {
					jyujiMonS = inRecs[J][206] + "ヵ月"
				}

				if jyujiYearS != "" && jyujiMonS != "" {
					cRec[83] = jyujiYearS + " " + jyujiMonS
				} else {
					cRec[83] = jyujiYearS + jyujiMonS
				}

				// 84.ｽﾁﾚﾝ_取扱い終了時期
				if dateRecs[0][7] != "スチレン取扱い終了年月日" {
					log.Print("「スチレン取扱い終了年月日」が見つかりませんでした。")
					failOnError(doError())
				}

				findflag = false
				for l, _ := range dateRecs {
					if cRec[4] == dateRecs[l][0] {
						if cRec[8] != dateRecs[l][3] {
							log.Printf("スチレン生年月日の不一致: %v %v %v != %v\r\n", cRec[4], cRec[6], cRec[8], dateRecs[l][3])
						}

						edate := dateRecs[l][7]
						if edate != "" {
							cRec[84] = edate[0:4] + "年" + edate[5:7] + "月" + edate[8:] + "日"
						} else {
							cRec[84] = edate
						}
						findflag = true
						break
					}
				}

				if findflag == false {
					log.Printf("取扱い終了名簿に対象がいません。 %v %v", cRec[4], cRec[6])
					cRec[84] = "err"
				}

				// 85.MIBK_従事年数
				jyujiYearM := ""
				if inRecs[J][207] != "" {
					jyujiYearM = inRecs[J][207] + "年"
				}

				jyujiMonM := ""
				if inRecs[J][208] != "" {
					jyujiMonM = inRecs[J][208] + "ヵ月"
				}

				if jyujiYearM != "" && jyujiMonM != "" {
					cRec[85] = jyujiYearM + " " + jyujiMonM
				} else {
					cRec[85] = jyujiYearM + jyujiMonM
				}

				// 86.MIBK_取扱い終了時期
				if dateRecs[0][8] != "MIBK取扱い終了年月日" {
					log.Print("「MIBK取扱い終了年月日」が見つかりませんでした。")
					failOnError(doError())
				}

				findflag = false
				for l, _ := range dateRecs {
					if cRec[4] == dateRecs[l][0] {
						if cRec[8] != dateRecs[l][3] {
							log.Printf("MIBK生年月日の不一致: %v %v %v != %v\r\n", cRec[4], cRec[6], cRec[8], dateRecs[l][3])
						}

						edate := dateRecs[l][8]
						if edate != "" {
							cRec[86] = edate[0:4] + "年" + edate[5:7] + "月" + edate[8:] + "日"
						} else {
							cRec[86] = edate
						}
						findflag = true
						break
					}
				}

				if findflag == false {
					log.Printf("取扱い終了名簿に対象がいません。 %v %v", cRec[4], cRec[6])
					cRec[86] = "err"
				}

				// 87.ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK_作業時間
				sagyoHour := ""
				if inRecs[J][215] != "" {
					sagyoHour = inRecs[J][215] + "時間"
				}

				sagyoMin := ""
				if inRecs[J][216] != "" {
					sagyoMin = inRecs[J][216] + "分"
				}

				if sagyoHour != "" && sagyoMin != "" {
					cRec[87] = sagyoHour + " " + sagyoMin
				} else {
					cRec[87] = sagyoHour + sagyoMin
				}

				// 88.ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK_従事日数_週月
				jyuji := ""
				if inRecs[J][217] != "" {
					jyuji = inRecs[J][217] + "日"
				}

				jyujiWM := ""
				if inRecs[J][218] != "" {
					jyujiWM = "/" + inRecs[J][218]
				}

				cRec[88] = jyuji + jyujiWM

				// 89.ｸﾛﾛﾎﾙﾑ他_白血球像判定
				cRec[89] = HanteiCode(inRecs[J][243])
				if HanteiCode(inRecs[J][243]) == "err" {
					log.Print("白血球像判定にエラーがあります\r\n")
				}

				// 90.ｸﾛﾛﾎﾙﾑ他_診察判定
				cRec[90] = HanteiCode(inRecs[J][235])
				if HanteiCode(inRecs[J][235]) == "err" {
					log.Print("診察所見判定にエラーがあります\r\n")
				}

				// 91.ジクロロメタン判定
				cRec[91] = HanteiCode(inRecs[J][245])
				if HanteiCode(inRecs[J][245]) == "err" {
					log.Print("ジクロロメタン判定にエラーがあります\r\n")
				}

				// 92.スチレン判定
				cRec[92] = HanteiCode(inRecs[J][247])
				if HanteiCode(inRecs[J][247]) == "err" {
					log.Print("スチレン判定にエラーがあります\r\n")
				}

				// 93.メチルイソブチルケトン判定
				cRec[93] = HanteiCode(inRecs[J][249])
				if HanteiCode(inRecs[J][249]) == "err" {
					log.Print("メチルイソブチルケトン判定にエラーがあります\r\n")
				}

				// 94.医療機関判定（ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK）
				// 95.医療機関判定（ｼﾞｸﾛﾛﾒﾀﾝｽﾁﾚﾝMIBK）
				sogow := 0
				var hw [8]int
				hw[0] = HanteiWeight(inRecs[J][235]) //診察所見判定
				hw[1] = HanteiWeight(inRecs[J][237]) //尿蛋白判定
				hw[2] = HanteiWeight(inRecs[J][239]) //肝機能判定
				hw[3] = HanteiWeight(inRecs[J][241]) //白血球数判定
				hw[4] = HanteiWeight(inRecs[J][243]) //白血球像判定
				hw[5] = HanteiWeight(inRecs[J][245]) //ジクロロメタン判定
				hw[6] = HanteiWeight(inRecs[J][247]) //スチレン判定
				hw[7] = HanteiWeight(inRecs[J][249]) //ﾒﾁﾙｲｿﾌﾞﾁﾙｹﾄﾝ判定

				for k := 0; k < 8; k++ {
					if sogow < hw[k] {
						sogow = hw[k]
					}
				}

				cRec[94] = WeightToHantei(sogow)
				cRec[95] = WeightToHanteiCode(sogow)

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
