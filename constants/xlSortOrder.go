package constants

type (
	// XlSortOrder は、指定したフィールドまたは範囲の並べ替え順序を表します。
	XlSortOrder int
)

// XlSortMethod -- 指定したフィールドまたは範囲の並べ替え順序を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlAscending  XlSortOrder = 1 //指定したフィールドを昇順で並べ替えます。これは既定値です。
	XlDescending XlSortOrder = 2 //指定したフィールドを降順で並べ替えます。
)
