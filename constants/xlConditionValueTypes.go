package constants

type (
	// XlConditionValueTypesは、使用できる条件値の種類を表します。
	XlConditionValueTypes int
)

// XlConditionValueTypes -- 使用できる条件値の種類を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlConditionValueAutomaticMax XlConditionValueTypes = 7  //最長のデータ バーは、範囲の最大値に比例します。
	XlConditionValueAutomaticMin XlConditionValueTypes = 6  //最短のデータ バーは、範囲の最小値に比例します。
	XlConditionValueFormula      XlConditionValueTypes = 4  //数式が使用されます
	XlConditionValueHighestValue XlConditionValueTypes = 2  //値の一覧の最高値
	XlConditionValueLowestValue  XlConditionValueTypes = 1  //値の一覧の最低値
	XlConditionValueNone         XlConditionValueTypes = -1 //条件値なし
	XlConditionValueNumber       XlConditionValueTypes = 0  //数字が使用されます
	XlConditionValuePercent      XlConditionValueTypes = 3  //パーセンテージが使用されます
	XlConditionValuePercentile   XlConditionValueTypes = 5  //百分位が使用されます
)
