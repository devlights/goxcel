package constants

type (
	// XlSortMethod は、並べ替えの種類を表します。
	XlSortMethod int
)

// XlSortMethod -- 並べ替えの種類を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlPinYin XlSortMethod = 1 //中国語の発音表記の順で並べ替えます。これは既定値です。
	XlStroke XlSortMethod = 2 //各文字の総画数で並べ替えます。
)
