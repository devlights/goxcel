package constants

type (
	// XlFileFormatは、ブックを保存する場合のファイル形式を表します。
	XlFileFormat int
)

// XlFileFormat -- ブックを保存する場合のファイル形式を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlAddIn                       XlFileFormat = 18    //Microsoft Excel 97-2003 アドイン
	XlAddIn8                      XlFileFormat = 18    //Microsoft Excel 97-2003 アドイン
	XlCSV                         XlFileFormat = 6     //CSV
	XlCSVMac                      XlFileFormat = 22    //Macintosh CSV
	XlCSVMSDOS                    XlFileFormat = 24    //MSDOS CSV
	XlCSVWindows                  XlFileFormat = 23    //Windows CSV
	XlCurrentPlatformText         XlFileFormat = -4158 //現在のプラットフォームのテキスト
	XlDBF2                        XlFileFormat = 7     //DBF2
	XlDBF3                        XlFileFormat = 8     //DBF3
	XlDBF4                        XlFileFormat = 11    //DBF4
	XlDIF                         XlFileFormat = 9     //DIF
	XlExcel12                     XlFileFormat = 50    //Excel12
	XlExcel2                      XlFileFormat = 16    //Excel2
	XlExcel2FarEast               XlFileFormat = 27    //Excel2 FarEast
	XlExcel3                      XlFileFormat = 29    //Excel3
	XlExcel4                      XlFileFormat = 33    //Excel4
	XlExcel4Workbook              XlFileFormat = 35    //Excel4 ブック
	XlExcel5                      XlFileFormat = 39    //Excel5
	XlExcel7                      XlFileFormat = 39    //Excel7
	XlExcel8                      XlFileFormat = 56    //Excel8
	XlExcel9795                   XlFileFormat = 43    //Excel9795
	XlHtml                        XlFileFormat = 44    //HTML 形式
	XlIntlAddIn                   XlFileFormat = 26    //International Add-In
	XlIntlMacro                   XlFileFormat = 25    //International Macro
	XlOpenDocumentSpreadsheet     XlFileFormat = 60    //OpenDocument スプレッドシートを開く
	XlOpenXMLAddIn                XlFileFormat = 55    //XML アドインを開く
	XlOpenXMLTemplate             XlFileFormat = 54    //XML テンプレートを開く
	XlOpenXMLTemplateMacroEnabled XlFileFormat = 53    //マクロを有効にした XML テンプレートを開く
	XlOpenXMLWorkbook             XlFileFormat = 51    //XML ブックを開く
	XlOpenXMLWorkbookMacroEnabled XlFileFormat = 52    //マクロを有効にした XML ブックを開く
	XlSYLK                        XlFileFormat = 2     //SYLK
	XlTemplate                    XlFileFormat = 17    //テンプレート
	XlTemplate8                   XlFileFormat = 17    //テンプレート 8
	XlTextMac                     XlFileFormat = 19    //Macintosh テキスト
	XlTextMSDOS                   XlFileFormat = 21    //MSDOS テキスト
	XlTextPrinter                 XlFileFormat = 36    //プリンター テキスト
	XlTextWindows                 XlFileFormat = 20    //Windows テキスト
	XlUnicodeText                 XlFileFormat = 42    //Unicode テキスト
	XlWebArchive                  XlFileFormat = 45    //Web アーカイブ
	XlWJ2WD1                      XlFileFormat = 14    //WJ2WD1
	XlWJ3                         XlFileFormat = 40    //WJ3
	XlWJ3FJ3                      XlFileFormat = 41    //WJ3FJ3
	XlWK1                         XlFileFormat = 5     //WK1
	XlWK1ALL                      XlFileFormat = 31    //WK1ALL
	XlWK1FMT                      XlFileFormat = 30    //WK1FMT
	XlWK3                         XlFileFormat = 15    //WK3
	XlWK3FM3                      XlFileFormat = 32    //WK3FM3
	XlWK4                         XlFileFormat = 38    //WK4
	XlWKS                         XlFileFormat = 4     //ワークシート
	XlWorkbookDefault             XlFileFormat = 51    //ブックの既定
	XlWorkbookNormal              XlFileFormat = -4143 // ブックの標準
	XlWorks2FarEast               XlFileFormat = 28    //Works2 FarEast
	XlWQ1                         XlFileFormat = 34    //WQ1
	XlXMLSpreadsheet              XlFileFormat = 46    //XML スプレッドシート
)
