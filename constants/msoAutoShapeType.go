package constants

type (
	// MsoAutoShapeTypeは、 AutoShapeオブジェクトの図形の種類を表します。
	MsoAutoShapeType int
)

// MsoAutoShapeType -- AutoShapeオブジェクトの図形の種類を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	msoShape16pointStar                      MsoAutoShapeType = 94  //星 16。
	MsoShape24pointStar                      MsoAutoShapeType = 95  //星 24。
	MsoShape32pointStar                      MsoAutoShapeType = 96  //星 32。
	MsoShape4pointStar                       MsoAutoShapeType = 91  //星 4。
	MsoShape5pointStar                       MsoAutoShapeType = 92  //星 5。
	MsoShape8pointStar                       MsoAutoShapeType = 93  //星 8。
	MsoShapeActionButtonBackorPrevious       MsoAutoShapeType = 129 //[戻る] または [前へ] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonBeginning            MsoAutoShapeType = 131 //[上旬] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonCustom               MsoAutoShapeType = 125 //既定の画像またはテキストのないボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonDocument             MsoAutoShapeType = 134 //[文書] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonEnd                  MsoAutoShapeType = 132 //[終了] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonForwardorNext        MsoAutoShapeType = 130 //[進む] または [次へ] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonHelp                 MsoAutoShapeType = 127 //[ヘルプ] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonHome                 MsoAutoShapeType = 126 //[ホーム] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonInformation          MsoAutoShapeType = 128 //[情報] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonMovie                MsoAutoShapeType = 136 //[ビデオ] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonReturn               MsoAutoShapeType = 133 //[戻る] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeActionButtonSound                MsoAutoShapeType = 135 //[サウンド] ボタン。マウスクリックおよびマウスオーバー動作をサポートします。
	MsoShapeArc                              MsoAutoShapeType = 25  //円弧。
	MsoShapeBalloon                          MsoAutoShapeType = 137 //吹き出し。
	MsoShapeBentArrow                        MsoAutoShapeType = 41  //90°の曲線に続くブロック矢印。
	MsoShapeBentUpArrow                      MsoAutoShapeType = 44  //90°の鋭角線に続くブロック矢印。既定では上向きです。
	MsoShapeBevel                            MsoAutoShapeType = 15  //斜角。
	MsoShapeBlockArc                         MsoAutoShapeType = 20  //アーチ。
	MsoShapeCan                              MsoAutoShapeType = 13  //円柱。
	MsoShapeChevron                          MsoAutoShapeType = 52  //山形。
	MsoShapeCircularArrow                    MsoAutoShapeType = 60  //180°の曲線に続くブロック矢印。
	MsoShapeCloudCallout                     MsoAutoShapeType = 108 //雲形吹き出し。
	MsoShapeCross                            MsoAutoShapeType = 11  //十字形。
	MsoShapeCube                             MsoAutoShapeType = 14  //直方体。
	MsoShapeCurvedDownArrow                  MsoAutoShapeType = 48  //下カーブ ブロック矢印。
	MsoShapeCurvedDownRibbon                 MsoAutoShapeType = 100 //下カーブ リボン。
	MsoShapeCurvedLeftArrow                  MsoAutoShapeType = 46  //左カーブ ブロック矢印。
	MsoShapeCurvedRightArrow                 MsoAutoShapeType = 45  //右カーブ ブロック矢印。
	MsoShapeCurvedUpArrow                    MsoAutoShapeType = 47  //上カーブ ブロック矢印。
	MsoShapeCurvedUpRibbon                   MsoAutoShapeType = 99  //上カーブリボン。
	MsoShapeDiamond                          MsoAutoShapeType = 4   //ひし形。
	MsoShapeDonut                            MsoAutoShapeType = 18  //ドーナツ。
	MsoShapeDoubleBrace                      MsoAutoShapeType = 27  //中かっこ。
	MsoShapeDoubleBracket                    MsoAutoShapeType = 26  //大かっこ。
	MsoShapeDoubleWave                       MsoAutoShapeType = 104 //小波。
	MsoShapeDownArrow                        MsoAutoShapeType = 36  //下向きブロック矢印。
	MsoShapeDownArrowCallout                 MsoAutoShapeType = 56  //下矢印の付いた吹き出し。
	MsoShapeDownRibbon                       MsoAutoShapeType = 98  //リボンの端よりも下に中央面があるリボン。
	MsoShapeExplosion1                       MsoAutoShapeType = 89  //爆発。
	MsoShapeExplosion2                       MsoAutoShapeType = 90  //爆発。
	MsoShapeFlowchartAlternateProcess        MsoAutoShapeType = 62  //代替処理フローチャート記号。
	MsoShapeFlowchartCard                    MsoAutoShapeType = 75  //カード フローチャート記号。
	MsoShapeFlowchartCollate                 MsoAutoShapeType = 79  //照合フローチャート記号。
	MsoShapeFlowchartConnector               MsoAutoShapeType = 73  //結合子フローチャート記号。
	MsoShapeFlowchartData                    MsoAutoShapeType = 64  //データ フローチャート記号。
	MsoShapeFlowchartDecision                MsoAutoShapeType = 63  //判断フローチャート記号。
	MsoShapeFlowchartDelay                   MsoAutoShapeType = 84  //論理積ゲート フローチャート記号。
	MsoShapeFlowchartDirectAccessStorage     MsoAutoShapeType = 87  //直接アクセス記憶フローチャート記号。
	MsoShapeFlowchartDisplay                 MsoAutoShapeType = 88  //表示フローチャート記号。
	MsoShapeFlowchartDocument                MsoAutoShapeType = 67  //書類フローチャート記号。
	MsoShapeFlowchartExtract                 MsoAutoShapeType = 81  //抜き出しフローチャート記号。
	MsoShapeFlowchartInternalStorage         MsoAutoShapeType = 66  //内部記憶フローチャート記号。
	MsoShapeFlowchartMagneticDisk            MsoAutoShapeType = 86  //磁気ディスク フローチャート記号。
	MsoShapeFlowchartManualInput             MsoAutoShapeType = 71  //手操作入力フローチャート記号。
	MsoShapeFlowchartManualOperation         MsoAutoShapeType = 72  //手作業フローチャート記号。
	MsoShapeFlowchartMerge                   MsoAutoShapeType = 82  //組み合わせフローチャート記号。
	MsoShapeFlowchartMultidocument           MsoAutoShapeType = 68  //複数書類フローチャート記号。
	MsoShapeFlowchartOffpageConnector        MsoAutoShapeType = 74  //他ページ結合子フローチャート記号。
	MsoShapeFlowchartOr                      MsoAutoShapeType = 78  //論理和フローチャート記号。
	MsoShapeFlowchartPredefinedProcess       MsoAutoShapeType = 65  //定義済み処理フローチャート記号。
	MsoShapeFlowchartPreparation             MsoAutoShapeType = 70  //準備フローチャート記号。
	MsoShapeFlowchartProcess                 MsoAutoShapeType = 61  //処理フローチャート記号。
	MsoShapeFlowchartPunchedTape             MsoAutoShapeType = 76  //せん孔テープ フローチャート記号。
	MsoShapeFlowchartSequentialAccessStorage MsoAutoShapeType = 85  //順次アクセス記憶フローチャート記号。
	MsoShapeFlowchartSort                    MsoAutoShapeType = 80  //分類フローチャート記号。
	MsoShapeFlowchartStoredData              MsoAutoShapeType = 83  //記憶データ フローチャート記号。
	MsoShapeFlowchartSummingJunction         MsoAutoShapeType = 77  //和接合フローチャート記号。
	MsoShapeFlowchartTerminator              MsoAutoShapeType = 69  //端子フローチャート記号。
	MsoShapeFoldedCorner                     MsoAutoShapeType = 16  //メモ。
	MsoShapeHeart                            MsoAutoShapeType = 21  //ハート。
	MsoShapeHexagon                          MsoAutoShapeType = 10  //六角形。
	MsoShapeHorizontalScroll                 MsoAutoShapeType = 102 //横巻き。
	MsoShapeIsoscelesTriangle                MsoAutoShapeType = 7   //二等辺三角形。
	MsoShapeLeftArrow                        MsoAutoShapeType = 34  //左向きブロック矢印。
	MsoShapeLeftArrowCallout                 MsoAutoShapeType = 54  //左矢印の付いた吹き出し。
	MsoShapeLeftBrace                        MsoAutoShapeType = 31  //左中かっこ。
	MsoShapeLeftBracket                      MsoAutoShapeType = 29  //左大かっこ。
	MsoShapeLeftRightArrow                   MsoAutoShapeType = 37  //左右ブロック矢印。
	MsoShapeLeftRightArrowCallout            MsoAutoShapeType = 57  //左右矢印の付いた吹き出し。
	MsoShapeLeftRightUpArrow                 MsoAutoShapeType = 40  //左、右、および上の 3 方向ブロック矢印。
	MsoShapeLeftUpArrow                      MsoAutoShapeType = 43  //左および上矢印の 2 方向ブロック矢印。
	MsoShapeLightningBolt                    MsoAutoShapeType = 22  //稲妻。
	MsoShapeLineCallout1                     MsoAutoShapeType = 109 //枠付きで、水平の吹き出し線の付いた吹き出し。
	MsoShapeLineCallout1AccentBar            MsoAutoShapeType = 113 //水平の強調線の付いた吹き出し。
	MsoShapeLineCallout1BorderandAccentBar   MsoAutoShapeType = 121 //枠付きで、水平の強調線の付いた吹き出し。
	MsoShapeLineCallout1NoBorder             MsoAutoShapeType = 117 //水平線の付いた吹き出し。
	MsoShapeLineCallout2                     MsoAutoShapeType = 110 //斜めの直線の付いた吹き出し。
	MsoShapeLineCallout2AccentBar            MsoAutoShapeType = 114 //斜めの吹き出し線と強調線の付いた吹き出し。
	MsoShapeLineCallout2BorderandAccentBar   MsoAutoShapeType = 122 //枠、斜めの直線、および強調線の付いた吹き出し。
	MsoShapeLineCallout2NoBorder             MsoAutoShapeType = 118 //枠および斜めの吹き出し線のない吹き出し。
	MsoShapeLineCallout3                     MsoAutoShapeType = 111 //折れ線の付いた吹き出し。
	MsoShapeLineCallout3AccentBar            MsoAutoShapeType = 115 //折れた吹き出し線と強調線の付いた吹き出し。
	MsoShapeLineCallout3BorderandAccentBar   MsoAutoShapeType = 123 //枠、折れた吹き出し線、強調線の付いた吹き出し。
	MsoShapeLineCallout3NoBorder             MsoAutoShapeType = 119 //枠および折れた吹き出し線のない吹き出し。
	MsoShapeLineCallout4                     MsoAutoShapeType = 112 //U 字型の吹き出し線分の付いた吹き出し。
	MsoShapeLineCallout4AccentBar            MsoAutoShapeType = 116 //強調線および U 字型の吹き出し線分の付いた吹き出し。
	MsoShapeLineCallout4BorderandAccentBar   MsoAutoShapeType = 124 //枠線、強調線、および U 字型の吹き出し線分の付いた吹き出し。
	MsoShapeLineCallout4NoBorder             MsoAutoShapeType = 120 //枠線および U 字型の吹き出し線分のない呼び出し。
	MsoShapeMixed                            MsoAutoShapeType = -2  //値のみを返します。その他の状態の組み合わせを示します。
	MsoShapeMoon                             MsoAutoShapeType = 24  //月。
	MsoShapeNoSymbol                         MsoAutoShapeType = 19  //禁止。
	MsoShapeNotchedRightArrow                MsoAutoShapeType = 50  //右向きの V 字型矢印。
	MsoShapeNotPrimitive                     MsoAutoShapeType = 138 //サポートされていません。
	MsoShapeOctagon                          MsoAutoShapeType = 6   //八角形。
	MsoShapeOval                             MsoAutoShapeType = 9   //楕円。
	MsoShapeOvalCallout                      MsoAutoShapeType = 107 //円形吹き出し。
	MsoShapeParallelogram                    MsoAutoShapeType = 2   //平行四辺形。
	MsoShapePentagon                         MsoAutoShapeType = 51  //ホームベース。
	MsoShapePlaque                           MsoAutoShapeType = 28  //ブローチ。
	MsoShapeQuadArrow                        MsoAutoShapeType = 39  //4 方向ブロック矢印。
	MsoShapeQuadArrowCallout                 MsoAutoShapeType = 59  //4 方向矢印の付いた吹き出し。
	MsoShapeRectangle                        MsoAutoShapeType = 1   //四角形。
	MsoShapeRectangularCallout               MsoAutoShapeType = 105 //四角形吹き出し。
	MsoShapeRegularPentagon                  MsoAutoShapeType = 12  //ホームベース。
	MsoShapeRightArrow                       MsoAutoShapeType = 33  //右向きブロック矢印。
	MsoShapeRightArrowCallout                MsoAutoShapeType = 53  //右矢印の付いた吹き出し。
	MsoShapeRightBrace                       MsoAutoShapeType = 32  //右中かっこ。
	MsoShapeRightBracket                     MsoAutoShapeType = 30  //右大かっこ。
	MsoShapeRightTriangle                    MsoAutoShapeType = 8   //直角三角形。
	MsoShapeRoundedRectangle                 MsoAutoShapeType = 5   //角丸四角形。
	MsoShapeRoundedRectangularCallout        MsoAutoShapeType = 106 //角丸四角形吹き出し。
	MsoShapeSmileyFace                       MsoAutoShapeType = 17  //スマイル。
	MsoShapeStripedRightArrow                MsoAutoShapeType = 49  //先にストライプの付いた右向きのブロック矢印。
	MsoShapeSun                              MsoAutoShapeType = 23  //太陽。
	MsoShapeTrapezoid                        MsoAutoShapeType = 3   //台形。
	MsoShapeUpArrow                          MsoAutoShapeType = 35  //上向きブロック矢印。
	MsoShapeUpArrowCallout                   MsoAutoShapeType = 55  //上矢印の付いた吹き出し。
	MsoShapeUpDownArrow                      MsoAutoShapeType = 38  //上下 2 方向ブロック矢印。
	MsoShapeUpDownArrowCallout               MsoAutoShapeType = 58  //上下のブロック矢印の付いた吹き出し。
	MsoShapeUpRibbon                         MsoAutoShapeType = 97  //リボンの端よりも上に中央面があるリボン。
	MsoShapeUTurnArrow                       MsoAutoShapeType = 42  //U 字型のブロック矢印。
	MsoShapeVerticalScroll                   MsoAutoShapeType = 101 //縦巻き。
	MsoShapeWave                             MsoAutoShapeType = 103 //大波。
)
