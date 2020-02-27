package constants

type (
	// XlsxConstantsは、Excel の各種メソッドで使用される定数を表します。
	XlsxConstants int
)

// XlsxConstants -- この列挙により、Excel の各種メソッドで使用される定数がまとめられます。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	Xl3DBar                 XlsxConstants = -4099 //3D 横棒
	Xl3DEffects1            XlsxConstants = 13    //3-D 1
	Xl3DEffects2            XlsxConstants = 14    //3-D 2
	Xl3DSurface             XlsxConstants = -4103 //3D 表面
	XlAbove                 XlsxConstants = 0     //上
	XlAccounting1           XlsxConstants = 4     //会計 1
	XlAccounting2           XlsxConstants = 5     //会計 2
	XlAccounting4           XlsxConstants = 17    //会計 4
	XlAdd                   XlsxConstants = 2     //追加
	XlAll                   XlsxConstants = -4104 //すべて
	XlAccounting3           XlsxConstants = 6     //会計 3
	XlAllExceptBorders      XlsxConstants = 7     //罫線を除くすべて
	XlAutomatic             XlsxConstants = -4105 //自動
	XlBar                   XlsxConstants = 2     //自動
	XlBelow                 XlsxConstants = 1     //下
	XlBidi                  XlsxConstants = -5000 //右から左へ記述する言語
	XlBidiCalendar          XlsxConstants = 3     //BidiCalendar
	XlBoth                  XlsxConstants = 1     //両方
	XlBottom                XlsxConstants = -4107 //下
	XlCascade               XlsxConstants = 7     //重ねて表示
	XlCenter                XlsxConstants = -4108 //中央
	XlCenterAcrossSelection XlsxConstants = 7     //選択範囲内で中央
	XlChart4                XlsxConstants = 2     //グラフ 4
	XlChartSeries           XlsxConstants = 17    //グラフ系列
	XlChartShort            XlsxConstants = 6     //グラフ短縮
	XlChartTitles           XlsxConstants = 18    //グラフ タイトル
	XlChecker               XlsxConstants = 9     //市松模様
	XlCircle                XlsxConstants = 8     //円
	XlClassic1              XlsxConstants = 1     //一般 1
	XlClassic2              XlsxConstants = 2     //一般 2
	XlClassic3              XlsxConstants = 3     //一般 3
	XlClosed                XlsxConstants = 3     //更新不可
	XlColor1                XlsxConstants = 7     //色 1
	XlColor2                XlsxConstants = 8     //色 2
	XlColor3                XlsxConstants = 9     //色 3
	XlColumn                XlsxConstants = 3     //縦棒グラフ
	XlCombination           XlsxConstants = -4111 //複合グラフ
	XlComplete              XlsxConstants = 4     //完了
	XlConstants             XlsxConstants = 2     //定数
	XlContents              XlsxConstants = 2     //値
	XlContext               XlsxConstants = -5002 //対象
	XlCorner                XlsxConstants = 2     //コーナー
	XlCrissCross            XlsxConstants = 16    //クリスクロス
	XlCross                 XlsxConstants = 4     //交差
	XlCustom                XlsxConstants = -4114 //ユーザー定義
	XlDebugCodePane         XlsxConstants = 13    //デバッグ コード ペイン
	XlDefaultAutoFormat     XlsxConstants = -1    //既定のオートフォーマット
	XlDesktop               XlsxConstants = 9     //デスクトップ
	XlDiamond               XlsxConstants = 2     //ひし形
	XlDirect                XlsxConstants = 1     //直接
	XlDistributed           XlsxConstants = -4117 //均等割り付け
	XlDivide                XlsxConstants = 5     //除算
	XlDoubleAccounting      XlsxConstants = 5     //二重下線 (会計)
	XlDoubleClosed          XlsxConstants = 5     //二重引用符 (右)
	XlDoubleOpen            XlsxConstants = 4     //二重引用符 (左)
	XlDoubleQuote           XlsxConstants = 1     //二重引用符
	XlDrawingObject         XlsxConstants = 14    //描画オブジェクト
	XlEntireChart           XlsxConstants = 20    //グラフ全体
	XlExcelMenus            XlsxConstants = 1     //Excel メニュー
	XlExtended              XlsxConstants = 3     //拡張
	XlFill                  XlsxConstants = 5     //塗りつぶし
	XlFirst                 XlsxConstants = 0     //先頭
	XlFixedValue            XlsxConstants = 1     //固定値
	XlFloating              XlsxConstants = 5     //浮動
	XlFormats               XlsxConstants = -4122 //書式
	XlFormula               XlsxConstants = 5     //数式
	XlFullScript            XlsxConstants = 1     //フル スクリプト
	XlGeneral               XlsxConstants = 1     //標準
	XlGray16                XlsxConstants = 17    //灰色 16
	XlGray25                XlsxConstants = -4124 //灰色 25
	XlGray50                XlsxConstants = -4125 //灰色 50
	XlGray75                XlsxConstants = -4126 //灰色 75
	XlGray8                 XlsxConstants = 18    //灰色 8
	XlGregorian             XlsxConstants = 2     //グレゴリオ暦
	XlGrid                  XlsxConstants = 15    //グリッド
	XlGridline              XlsxConstants = 22    //目盛線
	XlHigh                  XlsxConstants = -4127 //高
	XlHindiNumerals         XlsxConstants = 3     //ヒンディー語の数字
	XlIcons                 XlsxConstants = 1     //アイコン
	XlImmediatePane         XlsxConstants = 12    //イミディエイト ペイン
	XlInside                XlsxConstants = 2     //内側
	XlInteger               XlsxConstants = 2     //整数
	XlJustify               XlsxConstants = -4130 //両端揃え
	XlLast                  XlsxConstants = 1     //末尾
	XlLastCell              XlsxConstants = 11    //最後のセル
	XlLatin                 XlsxConstants = -5001 //ラテン語
	XlLeft                  XlsxConstants = -4131 //左
	XlLeftToRight           XlsxConstants = 2     //左から右へ
	XlLightDown             XlsxConstants = 13    //暗くする
	XlLightHorizontal       XlsxConstants = 11    //横線
	XlLightUp               XlsxConstants = 14    //明るくする
	XlLightVertical         XlsxConstants = 12    //縦線
	XlList1                 XlsxConstants = 10    //リスト 1
	XlList2                 XlsxConstants = 11    //リスト 2
	XlList3                 XlsxConstants = 12    //リスト 3
	XlLocalFormat1          XlsxConstants = 15    //ローカル書式設定 1
	XlLocalFormat2          XlsxConstants = 16    //ローカル書式設定 2
	XlLogicalCursor         XlsxConstants = 1     //論理カーソル
	XlLong                  XlsxConstants = 3     //長整数型
	XlLotusHelp             XlsxConstants = 2     //Lotus ヘルプ
	XlLow                   XlsxConstants = -4134 //低
	XlLTR                   XlsxConstants = -5003 //LTR
	XlMacrosheetCell        XlsxConstants = 7     //マクロシート セル
	XlManual                XlsxConstants = -4135 //手動
	XlMaximum               XlsxConstants = 2     //最大
	XlMinimum               XlsxConstants = 4     //最小値
	XlMinusValues           XlsxConstants = 3     //マイナス値
	XlMixed                 XlsxConstants = 2     //混在
	XlMixedAuthorizedScript XlsxConstants = 4     //混在承認スクリプト
	XlMixedScript           XlsxConstants = 3     //混在スクリプト
	XlModule                XlsxConstants = -4141 //モジュール
	XlMultiply              XlsxConstants = 4     //乗算
	XlNarrow                XlsxConstants = 1     //狭い
	XlNextToAxis            XlsxConstants = 4     //軸の下/左
	XlNoDocuments           XlsxConstants = 3     //ドキュメントなし
	XlNone                  XlsxConstants = -4142 //なし
	XlNotes                 XlsxConstants = -4144 //メモ
	XlOff                   XlsxConstants = -4146 //オフ
	XlOn                    XlsxConstants = 1     //オン
	XlOpaque                XlsxConstants = 3     //塗りつぶし
	XlOpen                  XlsxConstants = 2     //開く
	XlOutside               XlsxConstants = 3     //外側
	XlPartial               XlsxConstants = 3     //一部
	XlPartialScript         XlsxConstants = 2     //スクリプトの一部
	XlPercent               XlsxConstants = 2     //パーセント
	XlPlus                  XlsxConstants = 9     //プラス記号
	XlPlusValues            XlsxConstants = 2     //プラス値
	XlReference             XlsxConstants = 4     //参照先
	XlRight                 XlsxConstants = -4152 //右
	XlRTL                   XlsxConstants = -5004 //RTL
	XlScale                 XlsxConstants = 3     //倍率
	XlSemiautomatic         XlsxConstants = 2     //半自動
	XlSemiGray75            XlsxConstants = 10    //SemiGray75
	XlShort                 XlsxConstants = 1     //短い形式の日付 (スラッシュ区切り)
	XlShowLabel             XlsxConstants = 4     //ラベルの表示
	XlShowLabelAndPercent   XlsxConstants = 5     //ラベルとパーセンテージを表示する
	XlShowPercent           XlsxConstants = 3     //パーセンテージを表示する
	XlShowValue             XlsxConstants = 2     //値を表示する
	XlSimple                XlsxConstants = -4154 //シンプル
	XlSingle                XlsxConstants = 2     //下線
	XlSingleAccounting      XlsxConstants = 4     //下線 (会計)
	XlSingleQuote           XlsxConstants = 2     //一重引用符
	XlSolid                 XlsxConstants = 1     //実線
	XlSquare                XlsxConstants = 1     //四角
	XlStar                  XlsxConstants = 5     //星
	XlStError               XlsxConstants = 4     //St エラー
	XlStrict                XlsxConstants = 2     //Strict
	XlSubtract              XlsxConstants = 3     //減算
	XlSystem                XlsxConstants = 1     //システム
	XlTextBox               XlsxConstants = 16    //テキスト ボックス
	XlTiled                 XlsxConstants = 1     //並べて表示
	XlTitleBar              XlsxConstants = 8     //タイトル バー
	XlToolbar               XlsxConstants = 1     //?ツールバー
	XlToolbarButton         XlsxConstants = 2     //ツールバー ボタン
	XlTop                   XlsxConstants = -4160 //上
	XlTopToBottom           XlsxConstants = 1     //上から下へ
	XlTransparent           XlsxConstants = 2     //透明
	XlTriangle              XlsxConstants = 3     //三角形
	XlVeryHidden            XlsxConstants = 2     //表示しない
	XlVisible               XlsxConstants = 12    //表示
	XlVisualCursor          XlsxConstants = 2     //表示カーソル
	XlWatchPane             XlsxConstants = 11    //ウォッチ ペイン
	XlWide                  XlsxConstants = 3     //広い
	XlWorkbookTab           XlsxConstants = 6     //ブック見出し
	XlWorksheet4            XlsxConstants = 1     //ワークシート 4
	XlWorksheetCell         XlsxConstants = 3     //ワークシート セル
	XlWorksheetShort        XlsxConstants = 5     //ワークシート短縮
)