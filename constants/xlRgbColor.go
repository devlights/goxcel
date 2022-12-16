package constants

type (
	// XlRgbColor は、カラーコード値を表します。
	XlRgbColor int
)

// XlRgbColor -- カラーコード値を指定します。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
// noinspection GoUnusedConst
const (
	RgbBlack                XlRgbColor = 0        //黒
	RgbNavy                 XlRgbColor = 8388608  //ネイビー
	RgbNavyBlue             XlRgbColor = 8388608  //ネイビーブルー
	RgbDarkBlue             XlRgbColor = 9109504  //濃い青
	RgbMediumBlue           XlRgbColor = 13434880 //淡い青
	RgbBlue                 XlRgbColor = 16711680 //青
	RgbDarkGreen            XlRgbColor = 25600    //濃い緑
	RgbGreen                XlRgbColor = 32768    //緑
	RgbTeal                 XlRgbColor = 8421376  //青緑
	RgbDarkCyan             XlRgbColor = 9145088  //濃いシアン
	RgbLightCyan            XlRgbColor = 9145088  //明るい水色
	RgbDeepSkyBlue          XlRgbColor = 16760576 //深いスカイブルー
	RgbDarkTurquoise        XlRgbColor = 13749760 //濃いターコイズ
	RgbMediumSpringGreen    XlRgbColor = 10156544 //淡いスプリンググリーン
	RgbLime                 XlRgbColor = 65280    //黄緑
	RgbSpringGreen          XlRgbColor = 8388352  //スプリンググリーン
	RgbAqua                 XlRgbColor = 16776960 //水色
	RgbMidnightBlue         XlRgbColor = 7346457  //ミッドナイトブルー
	RgbDodgerBlue           XlRgbColor = 16748574 //ドジャーブルー
	RgbLightSeaGreen        XlRgbColor = 11186720 //薄いシーグリーン
	RgbForestGreen          XlRgbColor = 2263842  //フォレストグリーン
	RgbSeaGreen             XlRgbColor = 5737262  //シーグリーン
	RgbDarkSlateGray        XlRgbColor = 5197615  //濃いスレートグレー
	RgbLimeGreen            XlRgbColor = 3329330  //ライムグリーン
	RgbMediumSeaGreen       XlRgbColor = 7451452  //淡いシーグリーン
	RgbTurquoise            XlRgbColor = 13688896 //ターコイズ
	RgbRoyalBlue            XlRgbColor = 14772545 //ロイヤルブルー
	RgbSteelBlue            XlRgbColor = 11829830 //スチールブルー
	RgbDarkSlateBlue        XlRgbColor = 9125192  //濃いスレートブルー
	RgbMediumTurquoise      XlRgbColor = 13422920 //淡いターコイズ
	RgbIndigo               XlRgbColor = 8519755  //インディゴ
	RgbDarkOliveGreen       XlRgbColor = 3107669  //濃いオリーブグリーン
	RgbCadetBlue            XlRgbColor = 10526303 //カデットブルー
	RgbCornflowerBlue       XlRgbColor = 15570276 //コーンフラワーブルー
	RgbMediumAquamarine     XlRgbColor = 11206502 //淡いアクアマリン
	RgbDimGray              XlRgbColor = 6908265  //ディムグレー
	RgbSlateBlue            XlRgbColor = 13458026 //スレートブルー
	RgbOliveDrab            XlRgbColor = 2330219  //オリーブドラブ
	RgbSlateGray            XlRgbColor = 9470064  //スレートグレー
	RgbLightSlateGray       XlRgbColor = 10061943 //薄いスレートグレー
	RgbMediumSlateBlue      XlRgbColor = 15624315 //淡いスレートブルー
	RgbLawnGreen            XlRgbColor = 64636    //若草色
	RgbChartreuse           XlRgbColor = 65407    //シャルトルーズ
	RgbAquamarine           XlRgbColor = 13959039 //アクアマリン
	RgbMaroon               XlRgbColor = 128      //栗色
	RgbPurple               XlRgbColor = 8388736  //紫
	RgbOlive                XlRgbColor = 32896    //オリーブ
	RgbGray                 XlRgbColor = 8421504  //灰色
	RgbGrey                 XlRgbColor = 8421504  //灰色
	RgbSkyBlue              XlRgbColor = 15453831 //スカイブルー
	RgbLightSkyBlue         XlRgbColor = 16436871 //薄いスカイブルー
	RgbBlueViolet           XlRgbColor = 14822282 //青紫
	RgbDarkRed              XlRgbColor = 139      //濃い赤
	RgbDarkMagenta          XlRgbColor = 9109643  //濃いマゼンタ
	RgbDarkSeaGreen         XlRgbColor = 9419919  //濃いシーグリーン
	RgbLightGreen           XlRgbColor = 9498256  //明るい緑
	RgbMediumPurple         XlRgbColor = 14381203 //淡い紫
	RgbDarkViolet           XlRgbColor = 13828244 //濃い紫
	RgbPaleGreen            XlRgbColor = 10025880 //ペールグリーン
	RgbDarkOrchid           XlRgbColor = 13382297 //濃いオーキッド
	RgbYellowGreen          XlRgbColor = 3329434  //イエローグリーン
	RgbSienna               XlRgbColor = 2970272  //シェンナ
	RgbBrown                XlRgbColor = 2763429  //茶
	RgbDarkGray             XlRgbColor = 11119017 //濃い灰色
	RgbLightBlue            XlRgbColor = 15128749 //明るい青
	RgbGreenYellow          XlRgbColor = 3145645  //グリーンイエロー
	RgbPaleTurquoise        XlRgbColor = 15658671 //ペールターコイズ
	RgbLightSteelBlue       XlRgbColor = 14599344 //薄いスチールブルー
	RgbPowderBlue           XlRgbColor = 15130800 //パウダーブルー
	RgbFireBrick            XlRgbColor = 2237106  //れんが色
	RgbDarkGoldenrod        XlRgbColor = 755384   //濃いゴールデンロッド
	RgbMediumOrchid         XlRgbColor = 13850042 //淡いオーキッド
	RgbRosyBrown            XlRgbColor = 9408444  //ローズブラウン
	RgbDarkKhaki            XlRgbColor = 7059389  //濃いカーキ
	RgbSilver               XlRgbColor = 12632256 //銀色
	RgbMediumVioletRed      XlRgbColor = 8721863  //淡いバイオレットレッド
	RgbIndianRed            XlRgbColor = 6053069  //インディアンレッド
	RgbPeru                 XlRgbColor = 4163021  //ペルー
	RgbTan                  XlRgbColor = 9221330  //タン
	RgbLightGray            XlRgbColor = 13882323 //薄い灰色
	RgbThistle              XlRgbColor = 14204888 //あざみ色
	RgbOrchid               XlRgbColor = 14053594 //オーキッド
	RgbGoldenrod            XlRgbColor = 2139610  //ゴールデンロッド
	RgbPaleVioletRed        XlRgbColor = 9662683  //ペールバイオレットレッド
	RgbCrimson              XlRgbColor = 3937500  //深紅
	RgbGainsboro            XlRgbColor = 14474460 //ゲーンズボロ
	RgbPlum                 XlRgbColor = 14524637 //プラム
	RgbBurlyWood            XlRgbColor = 8894686  //バーリーウッド
	RgbLavender             XlRgbColor = 16443110 //ラベンダー
	RgbDarkSalmon           XlRgbColor = 8034025  //濃いサーモンピンク
	RgbViolet               XlRgbColor = 15631086 //紫色
	RgbPaleGoldenrod        XlRgbColor = 7071982  //ペールゴールデンロッド
	RgbLightCoral           XlRgbColor = 8421616  //薄いさんご
	RgbKhaki                XlRgbColor = 9234160  //カーキ
	RgbAliceBlue            XlRgbColor = 16775408 //アリスブルー
	RgbHoneydew             XlRgbColor = 15794160 //ハニーデュー
	RgbAzure                XlRgbColor = 16777200 //空色
	RgbSandyBrown           XlRgbColor = 6333684  //サンディブラウン
	RgbWheat                XlRgbColor = 11788021 //小麦
	RgbBeige                XlRgbColor = 14480885 //ベージュ
	RgbWhiteSmoke           XlRgbColor = 16119285 //ホワイトスモーク
	RgbMintCream            XlRgbColor = 16449525 //ミントクリーム
	RgbGhostWhite           XlRgbColor = 16775416 //ゴーストホワイト
	RgbSalmon               XlRgbColor = 7504122  //サーモンピンク
	RgbAntiqueWhite         XlRgbColor = 14150650 //アンティークホワイト
	RgbLinen                XlRgbColor = 15134970 //リネン
	RgbLightGoldenrodYellow XlRgbColor = 13826810 //薄いゴールデンロッドイエロー
	RgbOldLace              XlRgbColor = 15136253 //オールドレース
	RgbRed                  XlRgbColor = 255      //赤
	RgbFuchsia              XlRgbColor = 16711935 //明るい紫
	RgbDeepPink             XlRgbColor = 9639167  //深いピンク
	RgbOrangeRed            XlRgbColor = 17919    //オレンジレッド
	RgbTomato               XlRgbColor = 4678655  //トマト
	RgbHotPink              XlRgbColor = 11823615 //ホットピンク
	RgbCoral                XlRgbColor = 5275647  //さんご
	RgbDarkOrange           XlRgbColor = 36095    //濃いオレンジ
	RgbLightSalmon          XlRgbColor = 8036607  //薄いサーモンピンク
	RgbOrange               XlRgbColor = 42495    //オレンジ
	RgbLightPink            XlRgbColor = 12695295 //薄いピンク
	RgbPink                 XlRgbColor = 13353215 //ピンク
	RgbGold                 XlRgbColor = 55295    //ゴールド
	RgbPeachPuff            XlRgbColor = 12180223 //ピーチパフ
	RgbNavajoWhite          XlRgbColor = 11394815 //ナバホホワイト
	RgbMoccasin             XlRgbColor = 11920639 //モカシン
	RgbBisque               XlRgbColor = 12903679 //ビスク
	RgbMistyRose            XlRgbColor = 14804223 //ミスティローズ
	RgbBlanchedAlmond       XlRgbColor = 13495295 //ブランシュアーモンド
	RgbPapayaWhip           XlRgbColor = 14020607 //パパイヤホイップ
	RgbLavenderBlush        XlRgbColor = 16118015 //ラベンダーブラッシュ
	RgbSeashell             XlRgbColor = 15660543 //シーシェル
	RgbCornsilk             XlRgbColor = 14481663 //コーンシルク
	RgbLemonChiffon         XlRgbColor = 13499135 //レモンシフォン
	RgbFloralWhite          XlRgbColor = 15792895 //フローラルホワイト
	RgbSnow                 XlRgbColor = 16448255 //スノー
	RgbYellow               XlRgbColor = 65535    //黄
	RgbLightYellow          XlRgbColor = 14745599 //明るい黄
	RgbIvory                XlRgbColor = 15794175 //アイボリー
	RgbWhite                XlRgbColor = 16777215 //白
)
