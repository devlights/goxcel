package constants

type (
	// XlDirection は、移動する方向を表します。
	XlDirection int
)

// XlDirection -- 移動する方向を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlDown    XlDirection = -4121 //下へ
	XlToLeft  XlDirection = -4159 //左へ
	XlToRight XlDirection = -4161 //右へ
	XlUp      XlDirection = -4162 //上へ
)
