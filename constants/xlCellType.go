package constants

type (
	// XlCellType は、セルの種類を表します。
	XlCellType int
)

// XlCellType -- セルの種類を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlCellTypeAllFormatConditions  XlCellType = -4172 //表示形式が設定されているセル
	XlCellTypeAllValidation        XlCellType = -4174 //条件の設定が含まれているセル
	XlCellTypeBlanks               XlCellType = 4     //空白セル
	XlCellTypeComments             XlCellType = -4144 //コメントが含まれているセル
	XlCellTypeConstants            XlCellType = 2     //定数が含まれているセル
	XlCellTypeFormulas             XlCellType = -4123 //数式が含まれているセル
	XlCellTypeLastCell             XlCellType = 11    //使われたセル範囲内の最後のセル
	XlCellTypeSameFormatConditions XlCellType = -4173 //同じ表示形式が設定されているセル
	XlCellTypeSameValidation       XlCellType = -4175 //同じ条件の設定が含まれているセル
	XlCellTypeVisible              XlCellType = 12    //すべての可視セル
)
