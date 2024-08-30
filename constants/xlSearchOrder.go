package constants

type (
	// XlSearchOrder は、検索する順序を表します。
	XlSearchOrder int
)

// XlSearchOrder -- 検索する順序を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlSearchOrderByColumns XlSearchOrder = 2 //列を下方向に検索してから、次の列に移動します。
	XlSearchOrderByRows    XlSearchOrder = 1 //行を横方向に検索してから、次の行に移動します。
)
