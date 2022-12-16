package constants

type (
	// XlSortOrientation は、並べ替え方向を表します。
	XlSortOrientation int
)

// XlSortOrientation -- 並べ替え方向を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlSortColumns XlSortOrientation = 1 //列単位で並べ替えます。
	XlSortRows    XlSortOrientation = 2 //行単位で並べ替えます。これは既定値です。
)
