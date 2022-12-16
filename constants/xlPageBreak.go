package constants

type (
	// XlPageBreak は、ワークシートのページブレークの場所を指定します。
	XlPageBreak int
)

//goland:noinspection ALL
const (
	// Excel が自動的に改ページを追加します。
	XlPageBreakAutomatic XlPageBreak = -4105
	// 改ページは手動で挿入されます。
	XlPageBreakManual XlPageBreak = -4135
	// 改ページはワークシートに挿入されません。
	XlPageBreakNone XlPageBreak = -4142
)
