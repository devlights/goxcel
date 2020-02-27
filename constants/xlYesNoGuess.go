package constants

type (
	// XlYesNoGuessは、先頭の行に見出しを含めるかどうかを表します。
	XlYesNoGuess int
)

// XlYesNoGuess -- 先頭の行に見出しを含めるかどうかを指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	XlGuess XlYesNoGuess = 0 //見出しがあるかどうか、ある場合はその場所を Excel が特定します。
	XlNo    XlYesNoGuess = 2 //既定値。範囲全体が並べ替えの対象になります。
	XlYes   XlYesNoGuess = 1 //範囲全体が並べ替えられません。
)
