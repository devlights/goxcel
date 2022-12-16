package constants

type (
	// XlAutoFilterOperator は、フィルターによって適用される 2 つの条件を関連付けるために使用する演算子を表します。
	XlAutoFilterOperator int
)

// XlAutoFilterOperator -- フィルターによって適用される 2 つの条件を関連付けるために使用する演算子を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlAnd             XlAutoFilterOperator = 1  //抽出条件 1 と抽出条件 2 の論理演算子 AND
	XlBottom10Items   XlAutoFilterOperator = 4  //表示される最低値項目 (抽出条件 1 で指定される項目数)
	XlBottom10Percent XlAutoFilterOperator = 6  //表示される最低値項目 (抽出条件 1 で指定される割合)
	XlFilterCellColor XlAutoFilterOperator = 8  //セルの色
	XlFilterDynamic   XlAutoFilterOperator = 11 //動的フィルター
	XlFilterFontColor XlAutoFilterOperator = 9  //フォントの色
	XlFilterIcon      XlAutoFilterOperator = 10 //フィルター アイコン
	XlFilterValues    XlAutoFilterOperator = 7  //フィルターの値
	XlOr              XlAutoFilterOperator = 2  //抽出条件 1 または抽出条件 2 の論理演算子 OR
	XlTop10Items      XlAutoFilterOperator = 3  //表示される最高値項目 (抽出条件 1 で指定される項目数)
	XlTop10Percent    XlAutoFilterOperator = 5  //表示される最高値項目 (抽出条件 1 で指定される割合)
)
