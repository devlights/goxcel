package constants

type (
	// XlSortDataOption は、テキストを並べ替える方法を表します。
	XlSortDataOption int
)

// XlSortDataOption -- テキストを並べ替える方法を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlSortNormal        XlSortDataOption = 0 //既定値。数値データとテキスト データを別々に並べ替えます。
	XlSortTextAsNumbers XlSortDataOption = 1 //テキストを数値データとして並べ替えます。
)
