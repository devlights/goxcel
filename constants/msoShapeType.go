package constants

type (
	// MsoShapeTypeは、図形の種類または範囲を表します。
	MsoShapeType int
)

// MsoShapeType -- 図形の種類または範囲を指定します。 Shape.Typeの判定に使います。
//
// REFERENCES::
//   - https://excel-ubara.com/EXCEL/EXCEL905.html
//
//noinspection GoUnusedConst
const (
	MsoAutoShape         MsoShapeType = 1  //オートシェイプ
	MsoCallout           MsoShapeType = 2  //引き出し線
	MsoCanvas            MsoShapeType = 20 //キャンバス
	MsoChart             MsoShapeType = 3  //グラフ
	MsoComment           MsoShapeType = 4  //コメント
	MsoDiagram           MsoShapeType = 21 //ダイアグラム
	MsoEmbeddedOLEObject MsoShapeType = 7  //埋め込み OLE オブジェクト
	MsoFormControl       MsoShapeType = 8  //フォーム コントロール
	MsoFreeform          MsoShapeType = 5  //フリーフォーム
	MsoGroup             MsoShapeType = 6  //グループ
	MsoIgxGraphic        MsoShapeType = 24 //SmartArt グラフィック
	MsoInk               MsoShapeType = 22 //インク
	MsoInkComment        MsoShapeType = 23 //インク コメント
	MsoLine              MsoShapeType = 9  //直線
	MsoLinkedOLEObject   MsoShapeType = 10 //リンク OLE オブジェクト
	MsoLinkedPicture     MsoShapeType = 11 //リンク画像
	MsoMedia             MsoShapeType = 16 //メディア
	MsoOLEControlObject  MsoShapeType = 12 //OLE コントロール オブジェクト
	MsoPicture           MsoShapeType = 13 //画像
	MsoPlaceholder       MsoShapeType = 14 //プレースホルダー
	MsoScriptAnchor      MsoShapeType = 18 //スクリプト アンカー
	MsoShapeTypeMixed    MsoShapeType = -2 //図形の種類の組み合わせ
	MsoTable             MsoShapeType = 19 //テーブル
	MsoTextBox           MsoShapeType = 17 //テキスト ボックス
	MsoTextEffect        MsoShapeType = 15 //テキスト効果
)
