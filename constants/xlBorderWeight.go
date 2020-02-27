package constants

type (
	// XlBorderWeightは、罫線の太さを表します。
	XlBorderWeight int
)

// XlBorderWeight -- 罫線の太さを指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlHairline XlBorderWeight = 1     //細線 (最も細い罫線)
	XlMedium   XlBorderWeight = -4138 //普通
	XlThick    XlBorderWeight = 4     //太線 (最も太い罫線)
	XlThin     XlBorderWeight = 2     //極細
)
