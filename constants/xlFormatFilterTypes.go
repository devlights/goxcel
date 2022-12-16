package constants

type (
	// XlFormatFilterTypes は、書式フィルターの種類を表します。
	XlFormatFilterTypes int
)

// XlFormatFilterTypes -- 書式フィルターの種類を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	FilterBottom        XlFormatFilterTypes = 0 //下
	FilterBottomPercent XlFormatFilterTypes = 2 //最低パーセント
	FilterTop           XlFormatFilterTypes = 1 //上
	FilterTopPercent    XlFormatFilterTypes = 3 //最高パーセント
)
