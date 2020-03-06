package constants

type (
	// NumberFormatLocal は、表示書式を表します。
	NumberFormatLocal string
)

// NumberFormatLocal -- 表示形式を指定します。
//
// REFERENCES::
//   - https://www.tipsfound.com/vba/07015
//
//noinspection GoUnusedConst
const (
	FormatNormal     NumberFormatLocal = "G/標準"
	FormatNumber     NumberFormatLocal = "0_"
	FormatCurrency   NumberFormatLocal = `\#,##0;\-#,##0`
	FormatAccounting NumberFormatLocal = `_ * #,##0_ ;_ * -#,##0_ ;_ * "-"_ ;_ @_`
	FormatDate       NumberFormatLocal = `yyyy/m/d`
	FormatTime       NumberFormatLocal = `[$-F400]h:mm:ss AM/PM`
	FormatPercentage NumberFormatLocal = `0%`
	FormatString     NumberFormatLocal = "@"
)
