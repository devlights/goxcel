package constants

type (
	// XlCVErrorは、セルのエラー番号と値を表します。
	XlCVError int
)

// XlCVError -- セルのエラー番号と値を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlErrDiv0  XlCVError = 2007 //エラー番号 : 2007
	XlErrNA    XlCVError = 2042 //エラー番号 : 2042
	XlErrName  XlCVError = 2029 //エラー番号 : 2029
	XlErrNull  XlCVError = 2000 //エラー番号 : 2000
	XlErrNum   XlCVError = 2036 //エラー番号 : 2036
	XlErrRef   XlCVError = 2023 //エラー番号 : 2023
	XlErrValue XlCVError = 2015 //エラー番号 : 2015
)
