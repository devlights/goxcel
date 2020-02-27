package constants

type (
	// XlPageOrientationは、ワークシートを印刷する場合のページの方向を表します。
	XlPageOrientation int
)

// XlPageOrientation -- ワークシートを印刷する場合のページの方向を指定します。
//
// REFERENCES::
//   - https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlpageorientation
//
//noinspection GoUnusedConst
const (
	XlPortrait  XlPageOrientation = 1 // 縦
	XlLandscape XlPageOrientation = 2 // 横
)
