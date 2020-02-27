package constants

type (
	// XlSortOnは、データを並べ替える基準となるパラメーターを表します。
	XlSortOn int
)

// XlSortMethod -- データを並べ替える基準となるパラメーターを指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	SortOnCellColor XlSortOn = 1 //セルの色
	SortOnFontColor XlSortOn = 2 //フォントの色
	SortOnIcon      XlSortOn = 3 //アイコン
	SortOnValues    XlSortOn = 0 //値
)
