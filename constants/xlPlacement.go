package constants

type (
	// XlPlacementは、シェイプが基になるセルを接続する方法を表します。
	XlPlacement int
)

// XlPlacement --シェイプが基になるセルを接続する方法を指定します。
//
// REFERENCES::
//   - https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xlplacement?view=excel-pia
//
//noinspection GoUnusedConst
const (
	XlMoveAndSize  XlPlacement = 1 //オブジェクトはセルと共に移動し、セルに合わせてサイズが変更されます。
	XlMove         XlPlacement = 2 //オブジェクトはセルと共に移動します。
	XlFreeFloating XlPlacement = 3 //オブジェクトは自由に動きます。
)
