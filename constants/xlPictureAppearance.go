package constants

type (
	// XlPictureAppearance は、図をコピーする方法を指定します。
	XlPictureAppearance int
)

// XlPictureAppearance -- 図をコピーする方法を指定します。
//
// REFERENCES::
//   - https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xlpictureappearance?view=excel-pia
//
// noinspection GoUnusedConst
const (
	XlPrinter XlPictureAppearance = 2 // 画像は、印刷時に表示されるとおりにコピーされます。
	XlScreen  XlPictureAppearance = 1 // 画像は、画面の表示にできるだけ近いようにコピーされます。
)
