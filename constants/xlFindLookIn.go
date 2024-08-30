package constants

type (
	// XlFindLookIn は、検索するデータの種類を表します。
	XlFindLookIn int
)

// XlFindLookIn -- 検索するデータの種類を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlFindLookInComments         XlFindLookIn = -4144 //コメント
	XlFindLookInFormulas         XlFindLookIn = -4123 //数式
	XlFindLookInValues           XlFindLookIn = -4163 //値
	XlFindLookInCommentsThreaded XlFindLookIn = -4184 //スレッド化されたコメント
)
