package constants

type (
	// XlSearchDirection は、検索する場合の検索方向を表します。
	XlSearchDirection int
)

// XlSearchDirection -- 検索する場合の検索方向を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlSearchDirectionNext     XlSearchDirection = 1 //範囲内で、一致する次の値を検索します。
	XlSearchDirectionPrevious XlSearchDirection = 2 //範囲内で、一致する前の値を検索します。
)
