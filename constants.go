package goxcel

type XlFileFormat int

// XlFileFormat
//   - https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat?redirectedfrom=MSDN
//   - https://excwlvba.blogspot.com/2013/04/xlfileformat.html
//noinspection GoUnusedConst
const (
	XlFileFormatXlWorkbookDefault XlFileFormat = 50
	XlFileFormatXlOpenXMLWorkbook XlFileFormat = 51
	XlFileFormatXlExcel8          XlFileFormat = 56
)

type XlPageOrientation int

// XlPageOrientation
//   - https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlpageorientation
//noinspection GoUnusedConst
const (
	XlPageOrientationXlPortrait  XlPageOrientation = 1
	XlPageOrientationXlLandscape XlPageOrientation = 2
)
