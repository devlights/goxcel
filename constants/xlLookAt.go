package constants

type (
	// XlLookAtは、検索テキスト全体または検索テキストの一部を検索するかどうかを表します。
	XlLookAt int
)

// XlLookAt -- 検索テキスト全体または検索テキストの一部を検索するかどうかを指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlPart  XlLookAt = 2 //検索テキストの一部を検索します。
	XlWhole XlLookAt = 1 //検索テキスト全体を検索します。
)
