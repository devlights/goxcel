package constants

type (
	// XlCalculation は、計算モードを表します。
	XlCalculation int
)

// XlCalculation -- 計算モードを指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlCalculationAutomatic     XlCalculation = -4105 //Excel が再計算を制御します。
	XlCalculationManual        XlCalculation = -4135 //ユーザーが要求すると、計算が完了します。
	XlCalculationSemiautomatic XlCalculation = 2     //Excel が再計算を制御しますが、テーブル内の変更は無視します。
)
