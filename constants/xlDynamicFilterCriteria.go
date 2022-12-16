package constants

type (
	// XlDynamicFilterCriteria は、フィルターの条件を表します。
	XlDynamicFilterCriteria int
)

// XlDynamicFilterCriteria -- フィルターの条件を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	XlFilterAboveAverage              XlDynamicFilterCriteria = 33 //平均を上回る値をすべてフィルター。
	XlFilterAllDatesInPeriodApril     XlDynamicFilterCriteria = 24 //4 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodAugust    XlDynamicFilterCriteria = 28 //8 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodDecember  XlDynamicFilterCriteria = 32 //12 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodFebruray  XlDynamicFilterCriteria = 22 //2 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodJanuary   XlDynamicFilterCriteria = 21 //1 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodJuly      XlDynamicFilterCriteria = 27 //7 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodJune      XlDynamicFilterCriteria = 26 //6 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodMarch     XlDynamicFilterCriteria = 23 //3 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodMay       XlDynamicFilterCriteria = 25 //5 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodNovember  XlDynamicFilterCriteria = 31 //11 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodOctober   XlDynamicFilterCriteria = 30 //10 月の日付をすべてフィルター。
	XlFilterAllDatesInPeriodQuarter1  XlDynamicFilterCriteria = 17 //第 1 四半期の日付をすべてフィルター。
	XlFilterAllDatesInPeriodQuarter2  XlDynamicFilterCriteria = 18 //第 2 四半期の日付をすべてフィルター。
	XlFilterAllDatesInPeriodQuarter3  XlDynamicFilterCriteria = 19 //第 3 四半期の日付をすべてフィルター。
	XlFilterAllDatesInPeriodQuarter4  XlDynamicFilterCriteria = 20 //第 4 四半期の日付をすべてフィルター。
	XlFilterAllDatesInPeriodSeptember XlDynamicFilterCriteria = 29 //9 月の日付をすべてフィルター。
	XlFilterBelowAverage              XlDynamicFilterCriteria = 34 //平均未満の値をすべてフィルター。
	XlFilterLastMonth                 XlDynamicFilterCriteria = 8  //先月に関する値をすべてフィルター。
	XlFilterLastQuarter               XlDynamicFilterCriteria = 11 //前四半期に関する値をすべてフィルター。
	XlFilterLastWeek                  XlDynamicFilterCriteria = 5  //先週に関する値をすべてフィルター。
	XlFilterLastYear                  XlDynamicFilterCriteria = 14 //前年に関する値をすべてフィルター。
	XlFilterNextMonth                 XlDynamicFilterCriteria = 9  //来月に関する値をすべてフィルター。
	XlFilterNextQuarter               XlDynamicFilterCriteria = 12 //次の四半期に関する値をすべてフィルター。
	XlFilterNextWeek                  XlDynamicFilterCriteria = 6  //次週に関する値をすべてフィルター。
	XlFilterNextYear                  XlDynamicFilterCriteria = 15 //来年に関する値をすべてフィルター。
	XlFilterThisMonth                 XlDynamicFilterCriteria = 7  //今月に関する値をすべてフィルター。
	XlFilterThisQuarter               XlDynamicFilterCriteria = 10 //今四半期に関する値をすべてフィルター。
	XlFilterThisWeek                  XlDynamicFilterCriteria = 4  //今週に関する値をすべてフィルター。
	XlFilterThisYear                  XlDynamicFilterCriteria = 13 //今年に関する値をすべてフィルター。
	XlFilterToday                     XlDynamicFilterCriteria = 1  //今日に関する値をすべてフィルター。
	XlFilterTomorrow                  XlDynamicFilterCriteria = 3  //明日に関する値をすべてフィルター。
	XlFilterYearToDate                XlDynamicFilterCriteria = 16 //今日から 1 年前までの値をすべてフィルター。
	XlFilterYesterday                 XlDynamicFilterCriteria = 2  //昨日に関する値をすべてフィルター。
)
