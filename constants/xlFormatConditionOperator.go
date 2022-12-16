package constants

type (
	// XlFormatConditionOperator は、数式をセル内の値に対して比較するため、または xlBetween および xlNotBetween の場合 2 つの数式を比較するために使用する演算子を表します。
	XlFormatConditionOperator int
)

// XlFormatConditionOperator -- 数式をセル内の値に対して比較するため、または xlBetween および xlNotBetween の場合 2 つの数式を比較するために使用する演算子を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlBetween      XlFormatConditionOperator = 1 //間。2 つの数式が指定されている場合にのみ使用できます。
	XlEqual        XlFormatConditionOperator = 3 //等しい
	XlGreater      XlFormatConditionOperator = 5 //次の値より大きい
	XlGreaterEqual XlFormatConditionOperator = 7 //以上
	XlLess         XlFormatConditionOperator = 6 //次の値より小さい
	XlLessEqual    XlFormatConditionOperator = 8 //以下
	XlNotBetween   XlFormatConditionOperator = 2 //次の値の間以外。2 つの数式が指定されている場合にのみ使用できます。
	XlNotEqual     XlFormatConditionOperator = 4 //等しくない
)
