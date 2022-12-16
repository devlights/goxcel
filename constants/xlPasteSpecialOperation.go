package constants

type (
	// XlPasteSpecialOperation は、ワークシートの対象セルで数値データが計算される方法を表します。
	XlPasteSpecialOperation int
)

// XlPasteSpecialOperation -- ワークシートの対象セルで数値データが計算される方法を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlPasteSpecialOperationAdd      XlPasteSpecialOperation = 2     //コピーしたデータは、対象セルの値に追加されます。
	XlPasteSpecialOperationDivide   XlPasteSpecialOperation = 5     //コピーしたデータは、対象セルの値によって分割されます。
	XlPasteSpecialOperationMultiply XlPasteSpecialOperation = 4     //コピーしたデータには、対象セルの値が掛けられます。
	XlPasteSpecialOperationNone     XlPasteSpecialOperation = -4142 //貼り付け操作で計算は行われません。
	XlPasteSpecialOperationSubtract XlPasteSpecialOperation = 3     //コピーしたデータは、対象セルの値が引かれます。
)
