# IndexSearch
This program search keywords by Windows Search function.

パスと用語を指定するとWindows Searchを使って検索します。

# 使い方

Dosプロンプトまたは、PowerShellで使用してください。

以下オプションがあります。
 - --path 検索対象のフルパス。WindowsのIndexが設定されていないパスを指定すると、検索結果が出ません。
検索結果が出てこない場合は、”Windows Searchの設定”を確認してください。
指定しない場合は、カレントディレクトリ配下を検索します。
 - --word 検索対象の用語。""でくくってください。"Azure SQL"のように連続した単語も指定できます。
Windows Search SQLのCONTAINS術後の引数となります。以下も参照してください。
https://docs.microsoft.com/ja-jp/windows/win32/search/-search-sql-contains
