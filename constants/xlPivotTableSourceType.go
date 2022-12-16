package constants

type (
	// XlPivotTableSourceType は、レポート データのソースを表します。
	XlPivotTableSourceType int
)

// XlPivotTableSourceType -- レポート データのソースを指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlConsolidation XlPivotTableSourceType = 3     //複数のワークシート範囲
	XlDatabase      XlPivotTableSourceType = 1     //Excel のリスト/データベース
	XlExternal      XlPivotTableSourceType = 2     //外部のアプリケーションのデータ
	XlPivotTable    XlPivotTableSourceType = -4148 //既存のピボットテーブル レポート
	XlScenario      XlPivotTableSourceType = 4     //データは、[シナリオの登録と管理] を使用して作成されたシナリオに基づきます。
)
