package constants

type (
	// XlDeleteShiftDirectionは、セルを削除した後でセルをどのようにシフトするかを表します。
	XlDeleteShiftDirection int
)

// XlDeleteShiftDirection -- セルを削除した後でセルをどのようにシフトするかを指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlShiftToLeft XlDeleteShiftDirection = -4159 //セルは左にシフトします。
	XlShiftUp     XlDeleteShiftDirection = -4162 //セルは上にシフトします。
)
