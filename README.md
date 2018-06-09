### Offie文書（xls, doc, ppt）の『印刷日』を取得するには

Microsoft Officeの古典的なファイル形式（``xls``, ``doc``, ``ppt``）は，いずれもCFBF（[Compound File Binary Format](https://en.wikipedia.org/wiki/Compound_File_Binary_Format)）と呼ばれる形式のバイナリファイルです。

**注記**：「Windows 複合形式」「[OLE 構造化記憶](https://en.wikipedia.org/wiki/COM_Structured_Storage)」と表現されることもあります。

CFBFファイルは，さまざまなタイプのデータが『出し入れ』できるよう，セクター・アロケーションテーブル・ディレクトリといった論理的な構造を持っており，ファイルシステムに設計が似ています。セクターのサイズが``512``バイトあるいは``4096``バイトに固定されており，データの総量よりも余分にサイズを占有するため，大抵のCFBFファイルには圧縮が施されてます。

[CFBF](https://msdn.microsoft.com/en-us/library/dd942138.aspx)の設計は公開されています。

CFBFファイルは，構造的なファイルであるという点でJSONやXMLに似ています。

下記のプラグインを使用すれば，CFBFファイルをJSONに変換することができます。

https://github.com/miyako/4d-plugin-CFBF

Offie文書（[XLS](https://msdn.microsoft.com/en-us/library/office/cc313106(v=office.12).aspx), [DOC](https://msdn.microsoft.com/en-us/library/office/cc313153(v=office.12).aspx), [PPT](https://msdn.microsoft.com/en-us/library/office/cc313154(v=office.12).aspx)）の構造は公開されています。

CFBFの構造がわかれば，DOC・XLS・PPTなどの複合バイナリファイルから個別のデータを取り出すことができます。

個別のデータもそれぞれがバイナリ形式の構造体です。構造体の定義は，上述した資料に加え，[DTYP](https://msdn.microsoft.com/en-us/library/cc230273.aspx), [OAUT](https://msdn.microsoft.com/en-us/library/cc237549.aspx), [OLEPS](https://msdn.microsoft.com/en-us/library/dd942421.aspx), [OSHARED](https://msdn.microsoft.com/en-us/library/office/cc313156(v=office.12).aspx)など，いくつかの仕様書に記述されています。それらを参照して，4DのBLOBコマンドを使用することにより，これらを数値・日付・テキストなどのデータ型に変換することができます。

**ポイント**

* XLS, DOC, PPTはどれもCFBF，つまり複合バイナリファイルである
* CFBFは[4d-plugin-CFBF](https://github.com/miyako/4d-plugin-CFBF)で個別のバイナリデータに分解できる
* 個別のバイナリデータは，BLOBコマンドにより，数値・日付・テキストなどのデータ型に変換することができる
* 個別のバイナリデータに何が記録されているのかは，[XLS](https://msdn.microsoft.com/en-us/library/office/cc313106(v=office.12).aspx), [DOC](https://msdn.microsoft.com/en-us/library/office/cc313153(v=office.12).aspx), [PPT](https://msdn.microsoft.com/en-us/library/office/cc313154(v=office.12).aspx)に説明されている
* データがどのような形式で記録されているのかは，[DTYP](https://msdn.microsoft.com/en-us/library/cc230273.aspx), [OAUT](https://msdn.microsoft.com/en-us/library/cc237549.aspx), [OLEPS](https://msdn.microsoft.com/en-us/library/dd942421.aspx), [OSHARED](https://msdn.microsoft.com/en-us/library/office/cc313156(v=office.12).aspx)に説明されている

#### 印刷日を取得する

スプレッドシートをExcel 97-2004形式（``xls``）で保存します。

CFBFプラグインでこれを解析すると，下記のようなJSONが返されます。

```
{
	"storages" : [
		{
			"name" : "\u0001CompObj",
			"size" : 115,
			"data" : 1
		},
		{
			"name" : "Workbook",
			"size" : 15813,
			"data" : 2
		},
		{
			"name" : "\u0005SummaryInformation",
			"size" : 34940,
			"data" : 3
		},
		{
			"name" : "\u0005DocumentSummaryInformation",
			"size" : 252,
			"data" : 4
		}
	]
}
```

ストレージ名に制御文字（``\u0005``）が含まれていますが，仕様書を参照すると，これが正式名称であることが確認できます。（[MS-XLS].pdf, 2.1.7.4）

印刷日は，``\u0005SummaryInformation``のプロパティのひとつです。

```
ARRAY OBJECT($storages;0)
OB GET ARRAY($XLS;"storages";$storages)
C_TEXT($name)

For ($i;1;Size of array($storages))
	
	$storage:=$storages{$i}
	$name:=OB Get($storage;"name";Is text)
	
	Case of 
		: (Match regex("\\u0005DocumentSummaryInformation";$name))

		: (Match regex("\\u0005SummaryInformation";$name))

		: ($name="CompObj")

		: ($name="Workbook")
		
	End case 
	
End for 
```

``data``は，BLOB配列の要素番号です。つまり，``$2->{OB Get($storage;"data";Is longint)})``で``\u0005SummaryInformation``の値を取り出すことができます。

取り出したデータの形式はOSHAREDに準拠しています。``0x0000``~``0x0007``番地には，いくつかの基本情報が書かれています。

```
$byteOrder:=($1{0} << 8)+$1{1}  //0xFFFE is a reserved value
$version:=($1{2} << 8)+$1{3}  //version number of the property set
$OSMajorVersion:=$1{4}  //major version of the operating system that created the file
$OSMinorVersion:=$1{5}  //minor version of the operating system that created the file
$OSType:=($1{6} << 8)+$1{7}
```

``0x0010``番地は，アプリケーションのクラスIDです。

```
C_BLOB($applicationClsid)
COPY BLOB($1;$applicationClsid;8;0;0x0010)
```

``0x0018``番地は，セクションの数です。それぞれのセクションは，プロパティの集合です。

```
$pos:=0x0018
$cSections:=BLOB to longint($1;PC byte ordering;$pos)
```

``0x001C``番地からは，各セクションのクラスID，開始位置，サイズ，プロパティ数が書かれています。セクションの開始位置にジャンプすると，各プロパティの識別子と開始位置（オフセット）が書かれています。プロパティの識別子がわかれば，その形式が特定できます。形式が``TypedPropertyValue``の場合，プロパティの開始位置にジャンプすることにより，実際の形式が特定できます。

たとえば，印刷日プロパティの識別子は``PIDSI_LASTPRINTED``つまり``0x000B``です。この形式は，``TypedPropertyValue``なので，まずプロパティの開始位置にジャンプします。すると値は``0x0040``つまり``FILETIME``形式であることがわかります。これは，``8``バイト（``64``ビット）の整数値であり，1601年1月1日から経過した時間を100ナノ秒単位で表現したものです。

4Dに64ビット整数型の変数はありませんので，[テキストで整数を計算する](https://github.com/miyako/4d-tips-text-integer-maths)か，近似値で構わないのであれば，下記のように実数でこれを計算することができます。

**注記**: MicrosoftのF``ILETIME``構造体は，1601年1月1日を起点としていますが，これを4Dの``Add to date``で処理すると，かなりの誤差が発生します。UNIX時間に変換（``11644473600``秒を追加）してから``Add to date``で処理すれば，正確な日付と時刻が得られます。

```
C_BLOB($1;$FILETIME)
C_TEXT($0)

SET BLOB SIZE($FILETIME;8)

If (BLOB size($1)=8)
	
	  //byte swap
	
	$FILETIME{0}:=$1{7}
	$FILETIME{1}:=$1{6}
	$FILETIME{2}:=$1{5}
	$FILETIME{3}:=$1{4}
	$FILETIME{4}:=$1{3}
	$FILETIME{5}:=$1{2}
	$FILETIME{6}:=$1{1}
	$FILETIME{7}:=$1{0}
	
	$propOffset:=0
	
	$dwLowDateTimeH:=0xFFFF & BLOB to integer($FILETIME;Macintosh byte ordering;$propOffset)
	$dwLowDateTimeL:=0xFFFF & BLOB to integer($FILETIME;Macintosh byte ordering;$propOffset)
	
	$dwHighDateTimeH:=0xFFFF & BLOB to integer($FILETIME;Macintosh byte ordering;$propOffset)
	$dwHighDateTimeL:=0xFFFF & BLOB to integer($FILETIME;Macintosh byte ordering;$propOffset)
	
	$dwLowDateTime:=($dwLowDateTimeH*0x00010000)+$dwLowDateTimeL
	$dwHighDateTime:=($dwHighDateTimeH*0x00010000)+$dwHighDateTimeL

	$seconds:=((4294967296*$dwLowDateTime)+$dwHighDateTime)/10000000
	$unixtime:=$seconds-11644473600

	$days:=$unixtime/86400
	$date:=Add to date(!1970-01-01!;0;0;$days)
	$time:=$unixtime%86400

	$0:=String(Year of($date);"0000")+"-"+\
	String(Month of($date);"00")+"-"+\
	String(Day of($date);"00")+"T"+\
	String(Time($time);HH MM SS)+".000Z"
	
End if 
```

* サンプル

```
$path:=Get 4D folder(Current resources folder)+"sample.xls"

C_BLOB($CFBF)
DOCUMENT TO BLOB($path;$CFBF)

C_TEXT($json)
ARRAY BLOB($bytes;0)

CFBF PARSE DATA ($CFBF;$json;$bytes)

$XLS:=PARSE_OLE (JSON Parse($json);->$bytes)

SET TEXT TO PASTEBOARD(JSON Stringify($XLS))
```

* 結果

```
{
	"SummaryInformation": {
		"byteOrder": 65279,
		"version": 0,
		"OSMajorVersion": 3,
		"OSMinorVersion": 10,
		"applicationClsid": "{00000000-0000-0000-0000-000000000000}",
		"OSType": 256,
		"cSections": 1,
		"sections": [
			{
				"sectionOffset": 48,
				"formatId": "{F29F85E0-4FF9-1068-91AB-08002B27B3D9}",
				"cbSection": 35012,
				"cProps": 13,
				"properties": [
					{
						"CODEPAGE": 65001,
						"TITLE": "タイトル",
						"SUBJECT": "件名",
						"AUTHOR": "Microsoft Office ユーザー",
						"KEYWORDS": "キーワード",
						"COMMENTS": "こめこめ",
						"LASTAUTHOR": "Microsoft Office ユーザー",
						"APPNAME": "Microsoft Macintosh Excel",
						"LASTPRINTED": "2018-06-09T21:27:32.000Z",
						"CREATE_DTM": "2018-06-08T09:11:34.000Z",
						"LASTSAVE_DTM": "2018-06-09T21:27:40.000Z",
						"DOC_SECURITY": 0
					}
				]
			}
		]
	},
	"DocumentSummaryInformation": {
		"byteOrder": 65279,
		"version": 0,
		"OSMajorVersion": 3,
		"OSMinorVersion": 10,
		"applicationClsid": "{00000000-0000-0000-0000-000000000000}",
		"OSType": 256,
		"cSections": 2,
		"sections": [
			{
				"sectionOffset": 68,
				"formatId": "{D5CDD502-2E9C-101B-9793-08002B2CF9AE}",
				"cbSection": 272,
				"cProps": 11,
				"properties": [
					{
						"CODEPAGE": 65001,
						"CATEGORY": "分類",
						"MANAGER": "管理者",
						"COMPANY": "フォーディー",
						"VERSION": "15.0",
						"SCALE": false,
						"LINKSDIRTY": false,
						"SHAREDDOC": false
					}
				]
			},
			{
				"sectionOffset": 340,
				"formatId": "{D5CDD505-2E9C-101B-9793-08002B2CF9AE}",
				"cbSection": 80,
				"cProps": 3,
				"properties": [
					{
						"0": null,
						"CODEPAGE": 65001
					}
				]
			}
		]
	}
}
```

### ダウンロード

https://github.com/4D-JP/4d-tips-office-document-information/releases/tag/1.0
