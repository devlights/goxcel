package constants

type (
	// XlFixedFormatType -- ファイル形式の種類を指定します。
	XlFixedFormatType int
)

// XlFixedFormatType -- ファイル形式の種類を指定します。
//
// REFERENCES:
//   - https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfixedformattype
//
//noinspection GoUnusedConst
const (
	XlTypePDF XlFixedFormatType = 0 // PDF
	XlTypeXPS XlFixedFormatType = 1 // XPS
)
