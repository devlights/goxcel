package constants

type (
	// XlUnderlineStyle は、文字の下線の種類を表します。
	XlUnderlineStyle int
)

// XlUnderlineStyle -- 文字の下線の種類を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlUnderlineStyleDouble           XlUnderlineStyle = -4119 //太い二重下線
	XlUnderlineStyleDoubleAccounting XlUnderlineStyle = 5     //並んだ 2 本の細い下線
	XlUnderlineStyleNone             XlUnderlineStyle = -4142 //下線なし
	XlUnderlineStyleSingle           XlUnderlineStyle = 2     //一重下線
	XlUnderlineStyleSingleAccounting XlUnderlineStyle = 4     //サポートされていません。
)
