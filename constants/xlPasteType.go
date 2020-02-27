package constants

type (
	// XlPasteTypeは、貼り付ける部分を表します。
	XlPasteType int
)

// XlPasteType -- 貼り付ける部分を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlPasteAll                          XlPasteType = -4104 //すべてを貼り付けます。
	XlPasteAllExceptBorders             XlPasteType = 7     //輪郭以外のすべてを貼り付けます。
	XlPasteAllMergingConditionalFormats XlPasteType = 14    //すべてを貼り付け、条件付き書式をマージします。
	XlPasteAllUsingSourceTheme          XlPasteType = 13    //ソースのテーマを使用してすべてを貼り付けます。
	XlPasteColumnWidths                 XlPasteType = 8     //コピーした列の幅を貼り付けます。
	XlPasteComments                     XlPasteType = -4144 //コメントを貼り付けます。
	XlPasteFormats                      XlPasteType = -4122 //コピーしたソースの形式を貼り付けます。
	XlPasteFormulas                     XlPasteType = -4123 //数式を貼り付けます。
	XlPasteFormulasAndNumberFormats     XlPasteType = 11    //数式と数値の書式を貼り付けます。
	XlPasteValidation                   XlPasteType = 6     //入力規則を貼り付けます。
	XlPasteValues                       XlPasteType = -4163 //値を貼り付けます。
	XlPasteValuesAndNumberFormats       XlPasteType = 12    //値と数値の書式を貼り付けます。
)
