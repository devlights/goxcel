package constants

type (
	// XlIMEMode は、入力モード値を表します。
	XlIMEMode int
)

// XlIMEMode -- 入力モード値を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlIMEModeAlpha        XlIMEMode = 8  //半角英数字
	XlIMEModeAlphaFull    XlIMEMode = 7  //全角英数字
	XlIMEModeDisable      XlIMEMode = 3  //無効
	XlIMEModeHangul       XlIMEMode = 10 //ハングル
	XlIMEModeHangulFull   XlIMEMode = 9  //全角ハングル
	XlIMEModeHiragana     XlIMEMode = 4  //ひらがな
	XlIMEModeKatakana     XlIMEMode = 5  //カタカナ
	XlIMEModeKatakanaHalf XlIMEMode = 6  //半角カタカナ
	XlIMEModeNoControl    XlIMEMode = 0  //コントロールなし
	XlIMEModeOff          XlIMEMode = 2  //オフ (英語モード)
	XlIMEModeOn           XlIMEMode = 1  //モード オン
)
