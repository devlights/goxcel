package constants

type (
	// XlCopyPictureFormat は、コピーされる画像の形式を指定します。
	XlCopyPictureFormat int
)

// XlCopyPictureFormat -- コピーされる画像の形式を指定します。
//
// REFERENCES::
//   - https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xlcopypictureformat?view=excel-pia
//
// noinspection GoUnusedConst
const (
	XlBitmap  XlCopyPictureFormat = 2     // ビットマップ (.bmp, .jpg, .gif)
	XlPicture XlCopyPictureFormat = -4147 // ドロー画像 (.png, .wmf, .mix)
)
