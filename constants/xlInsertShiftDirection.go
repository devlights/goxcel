package constants

type (
	// XlInsertShiftDirection は、挿入時にセルをシフトする方向を表します。
	XlInsertShiftDirection int
)

// XlInsertShiftDirection -- 挿入時にセルをシフトする方向を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlShiftDown    XlInsertShiftDirection = -4121 //セルを挿入後、下に伸ばす
	XlShiftToRight XlInsertShiftDirection = -4161 //セルを挿入後、右に伸ばす
)
