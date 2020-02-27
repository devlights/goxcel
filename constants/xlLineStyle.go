package constants

type (
	// XlLineStyleは、線の種類を表します。
	XlLineStyle int
)

// XlLineStyle -- 線の種類を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlContinuous    XlLineStyle = 1     //実線
	XlDash          XlLineStyle = -4115 //破線
	XlDashDot       XlLineStyle = 4     //一点鎖線
	XlDashDotDot    XlLineStyle = 5     //ニ点鎖線
	XlDot           XlLineStyle = -4118 //点線
	XlDouble        XlLineStyle = -4119 //2 本線
	XlLineStyleNone XlLineStyle = -4142 //線なし
	XlSlantDashDot  XlLineStyle = 13    //斜破線
)
