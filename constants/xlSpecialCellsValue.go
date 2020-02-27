package constants

type (
	// XlSpecialCellsValueは、特定の種類の値を持つセルを表します。
	XlSpecialCellsValue int
)

// XlSpecialCellsValue -- 特定の種類の値を持つセルを指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlErrors     XlSpecialCellsValue = 16 //エラーのあるセル
	XlLogical    XlSpecialCellsValue = 4  //論理値のあるセル
	XlNumbers    XlSpecialCellsValue = 1  //数値のあるセル
	XlTextValues XlSpecialCellsValue = 2  //テキストのあるセル
)
