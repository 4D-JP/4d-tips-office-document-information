# 4d-tips-office-document-information
Offie文書（xls, doc, ppt）の『印刷日』を取得するには

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

### 印刷日を取得する

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
