package constants

type (
	// XlBordersIndexは、設定する罫線を表します。
	XlBordersIndex int
)

// XlBordersIndex -- 設定する罫線を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlDiagonalDown     XlBordersIndex = 5  // 範囲内の各セルの左上隅から右下への罫線
	XlDiagonalUp       XlBordersIndex = 6  // 範囲内の各セルの左下隅から右上への罫線
	XlEdgeBottom       XlBordersIndex = 9  // 範囲内の下側の罫線
	XlEdgeLeft         XlBordersIndex = 7  // 範囲内の左端の罫線
	XlEdgeRight        XlBordersIndex = 10 // 範囲内の右端の罫線
	XlEdgeTop          XlBordersIndex = 8  // 範囲内の上側の罫線
	XlInsideHorizontal XlBordersIndex = 12 // 範囲外の罫線を除く、範囲内のすべてのセルの水平罫線
	XlInsideVertical   XlBordersIndex = 11 // 範囲外の罫線を除く、範囲内のすべてのセルの垂直罫線
)
