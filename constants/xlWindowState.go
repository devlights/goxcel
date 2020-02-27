package constants

type (
	// XlWindowStateは、ウィンドウの状態を表します。
	XlWindowState int
)

// XlWindowState -- ウィンドウの状態を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlMaximized XlWindowState = -4137 //最大化
	XlMinimized XlWindowState = -4140 //最小化
	XlNormal    XlWindowState = -4143 //標準
)
