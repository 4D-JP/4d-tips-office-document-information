# 4d-tips-office-document-information
Offie文書（xls, doc, ppt）の『印刷日』を取得するには

Microsoft Officeの古典的なファイル形式（``xls``, ``doc``, ``ppt``）は，いずれもCFBF（[Compound File Binary Format](https://en.wikipedia.org/wiki/Compound_File_Binary_Format)）と呼ばれる形式のバイナリファイルです。

**注記**：「Windows 複合形式」「[OLE 構造化記憶](https://en.wikipedia.org/wiki/COM_Structured_Storage)」と表現されることもあります。

CFBFファイルは，さまざまなタイプのデータが『出し入れ』できるよう，セクター・アロケーションテーブル・ディレクトリといった論理的な構造を持っており，ファイルシステムに設計が似ています。セクターのサイズが``512``バイトあるいは``4096``バイトに固定されており，データの総量よりも余分にサイズを占有するため，大抵のCFBFファイルには圧縮が施されてます。

[CFBF](https://msdn.microsoft.com/en-us/library/dd942138.aspx)の設計は公開されています。

CFBFファイルは，構造的なファイルであるという点でJSONやXMLに似ています。

下記のプラグインを使用すれば，CFBFファイルをJSONに変換することができます。

https://github.com/miyako/4d-plugin-CFBF

Offie文書（[xls](https://msdn.microsoft.com/en-us/library/office/cc313106(v=office.12).aspx), [doc](https://msdn.microsoft.com/en-us/library/office/cc313153(v=office.12).aspx), [ppt](https://msdn.microsoft.com/en-us/library/office/cc313154(v=office.12).aspx)）の構造は公開されています。

CFBFの構造がわかれば，DOC・XLS・PPTなどの複合バイナリファイルから個別のデータを取り出すことができます。

個別のデータもそれぞれがバイナリ形式の構造体です。構造体の定義は，上述した資料（MS-XLS, MS-DOC, MS-PPT）に加え，MS-DTYP, MS-OAUT, MS-OLEPS, MS-OSHAREDなど，いくつかの仕様書に記述されています。それらを参照して，4DのBLOBコマンドを使用することにより，これらを数値・日付・テキストなどのデータ型に変換することができます。

**ポイント**

* XLS, DOC, PPTはどれもCFBF，つまり複合バイナリファイルである
* CFBFは[4d-plugin-CFBF](https://github.com/miyako/4d-plugin-CFBF)で個別のバイナリデータに分解できる
* 個別のバイナリデータは，BLOBコマンドにより，数値・日付・テキストなどのデータ型に変換することができる
* 個別のバイナリデータに何が記録されているのかは，MS-XLS, MS-DOC, MS-PPTに説明されている
* データがどのような形式で記録されているのかは，MS-DTYP, MS-OAUT, MS-OLEPS, MS-OSHAREDに説明されている

