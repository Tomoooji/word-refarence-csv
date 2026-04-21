# word-refarence-csv

CSVファイルからMS Wordの参考文献を作成するプログラム

## ファイル構成

```text
word-refarence-csv/
├─ README.md          - プログラムの説明(このファイル)
├─ ConvertXML.py      - Pythonプログラム
└─ Report_Source.csv  - 参考資料管理用csvファイル
```

他にも.gitignoreとかLICENCEとかは気にしなくてよいです。

## 使い方

1. Pythonプログラムが動く環境を用意してください。(Anacondaでもokなはず)
2. Pythonプログラムとcsvファイル(名前を変えないでください)を同じフォルダの中にいれてください。
3. csvファイルの中に書籍/書籍のセクション/ジャーナルの記事/webサイトの記事の情報を入力してください。<注意！>Excelでcsvファイルを開くとxlsz形式での保存を推奨されますが,プログラムが読み込めなくなるのでしないでください。
4. Pythonプログラムを実行してください。csvファイルの情報が多い場合、少し時間がかかるかもしれません。
5. プログラムと同じフォルダのなかに`Report_Source.xml`というファイルが作成されたのを確認したら、Wordを立ち上げて参考資料を読み込みたい文書を開いてください。
6. 参考資料タブを選択して資料文献の管理から、作成されたxmlファイルを選択してください。
7. 読み込みたい文献を左側で選択して、コピーを押して文書中に取り込むことができます。
8. Wordの文書内で参考資料の番号を挿入したり、文献目録を挿入したりできます。
9. 参考資料のデータを更新した場合は 1~6 の手順をもう一度行ってください。

## csvファイルの入力様式

| Tag   | SourceType     | Author       | Editor               | Title                | Year | Month | Day | __Book__ | BookTitle       | Pablisher | Edition | Pages     | __Journal__ | JournalNamee | Volume | Issue | Pages | __Website__ | InternalSiteTitle  | ProductionCompany | URL                         | YearAccessed | MonthAccessed | DayAccessed |
| ----- | -------------- | ------------ | -------------------- | -------------------- | ---- | ----- | --- | -------------- | --------------- | --------- | ------- | --------- | ----------------- | ------------ | ------ | ----- | ----- | ----------------- | ------------------ | ----------------- | --------------------------- | ------------ | ------------- | ----------- |
| test1 | JournalArticle | Tomoji Kanae |                      | Test Journal Article | 2026 | 04    | 21  |                |                 |           |         |           |                   | BioScience   | 3      | 380   | 24-48 |                   |                    |                   |                             |              |               |             |
| test2 | Book           | Tomoji Kanae |                      | Test Book            | 2026 | 04    | 21  |                |                 | K.G.Univ  | 3       | 380-24048 |                   |              |        |       |       |                   |                    |                   |                             |              |               |             |
| test3 | BookSection    | Tomoji Kanae | School of Bioscience | Test Book Section    | 2026 | 04    | 21  |                | Test Book Title | K.G.Univ  | 3       | 380-24048 |                   |              |        |       |       |                   |                    |                   |                             |              |               |             |
| test4 | InternetSite   | Tomoji Kanae |                      | Test WEBsite         | 2026 | 04    | 21  |                |                 |           |         |           |                   |              |        |       |       |                   | Test WEBsite Title | K.G.Univ          | https://github.com/Tomoooji | 2026         | 04            | 21          |

各資料の情報が横列に並んでいて、列を増やす/減らすことで文献を増やす/減らすことが可能です。

各情報の内容は以下の表と対応しています。
    全体用
    - Tag
    - SourceType
    - Author
    - Editor
    - Title
    - Year
    - Month
    - Day
    Book用
    - BookTitle
    - Publisher
    - Edition
    - Pages
    Journal用
    - JournalName
    - Volume
    - Issue
    - Pages
    Website用
    - InternetSiteTitle
    - ProductionCompany
    - URL
    - YearAccessed
    - MonthAccessed
    - Day Accessed
    