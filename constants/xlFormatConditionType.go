package constants

type (
	// XlFormatConditionType は、セル値または演算式のどちらを基に条件付き書式を設定するかを表します。
	XlFormatConditionType int
)

// XlFormatConditionType -- セル値または演算式のどちらを基に条件付き書式を設定するかを指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlAboveAverageCondition XlFormatConditionType = 12 //平均以上の条件
	XlBlanksCondition       XlFormatConditionType = 10 //空白の条件
	XlCellValue             XlFormatConditionType = 1  //セルの値
	XlColorScale            XlFormatConditionType = 3  //カラー スケール
	XlDatabar               XlFormatConditionType = 4  //データバー
	XlErrorsCondition       XlFormatConditionType = 16 //エラー条件
	XlExpression            XlFormatConditionType = 2  //演算
	XlIconSet               XlFormatConditionType = 6  //アイコン セット
	XlNoBlanksCondition     XlFormatConditionType = 13 //空白の条件なし
	XlNoErrorsCondition     XlFormatConditionType = 17 //エラー条件なし
	XlTextString            XlFormatConditionType = 9  //テキスト文字列
	XlTimePeriod            XlFormatConditionType = 11 //期間
	XlTop10                 XlFormatConditionType = 5  //上位の 10 の値
	XlUniqueValues          XlFormatConditionType = 8  //一意の値
)
