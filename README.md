# word-refarence-csv

CSVファイルからMS Wordの参考文献を作成するプログラム

## ファイル構成

```text
word-refarence-csv/
├─ README.md          - プログラムの説明(このファイル)
├─ ConvertXML.py      - Pythonプログラム
└─ Report_Data.csv  - 参考資料管理用csvファイル
```

他の.gitignoreとかLICENCEとかは気にしなくてよいです。

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
    - Tag タグ(任意の文字列)  
    - SourceType 資料タイプ(Book/BookSection/JounalArticle/InternetSite)  
    - Author 著者名  
    - Editor 編集者名(BookSectionのみ)  
    - Title タイトル  
    - Year 年  
    - Month 月(なくても可)  
    - Day 日(なくても可)  
    Book用 (__Book__のところは空欄)  
    - BookTitle 書籍のタイトル(BookSectionのみ)  
    - Publisher 出版者  
    - Edition 版  
    - Pages ページ  
    Journal用 (__Journal__は空欄)  
    - JournalName 雑誌名  
    - Volume 巻  
    - Issue 号  
    - Pages ページ  
    Website用　(__Website__は空欄)  
    - InternetSiteTitle サイト名  
    - ProductionCompany 作成会社  
    - URL url  
    - YearAccessed アクセス年  
    - MonthAccessed アクセス月  
    - Day Accessed アクセス日  

## Product
```xml
This XML file does not appear to have any style information associated with it. The document tree is shown below.
<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="">
    <b:Source>
        <b:Tag>test1</b:Tag>
        <b:SourceType>JournalArticle</b:SourceType>
        <b:Author>
            <b:Author>
                <b:NameList>
                    <b:Person>
                        <b:Last>Tomoji Kanae</b:Last>
                    </b:Person>
                </b:NameList>
            </b:Author>
        </b:Author>
        <b:Title>Test Journal Article</b:Title>
        <b:Year>2026</b:Year>
        <b:Month>04</b:Month>
        <b:Day>21</b:Day>
        <b:JournalNamee>BioScience</b:JournalNamee>
        <b:Volume>3</b:Volume>
        <b:Issue>380</b:Issue>
        <b:Pages>24-48</b:Pages>
    </b:Source>
    <b:Source>
        <b:Tag>test2</b:Tag>
        <b:SourceType>Book</b:SourceType>
        <b:Author>
            <b:Author>
                <b:NameList>
                    <b:Person>
                        <b:Last>Tomoji Kanae</b:Last>
                    </b:Person>
                </b:NameList>
            </b:Author>
        </b:Author>
        <b:Title>Test Book</b:Title>
        <b:Year>2026</b:Year>
        <b:Month>04</b:Month>
        <b:Day>21</b:Day>
        <b:Pablisher>K.G.Univ</b:Pablisher>
        <b:Edition>3</b:Edition>
        <b:Pages>380-24048</b:Pages>
    </b:Source>
    <b:Source>
        <b:Tag>test3</b:Tag>
        <b:SourceType>BookSection</b:SourceType>
        <b:Author>
            <b:Author>
                <b:NameList>
                    <b:Person>
                        <b:Last>Tomoji Kanae</b:Last>
                    </b:Person>
                </b:NameList>
            </b:Author>
            <b:Editor>
                <b:NameList>
                    <b:Person>
                        <b:Last>School of Bioscience</b:Last>
                    </b:Person>
                </b:NameList>
            </b:Editor>
        </b:Author>
        <b:Title>Test Book Section</b:Title>
        <b:Year>2026</b:Year>
        <b:Month>04</b:Month>
        <b:Day>21</b:Day>
        <b:BookTitle>Test Book Title</b:BookTitle>
        <b:Pablisher>K.G.Univ</b:Pablisher>
        <b:Edition>3</b:Edition>
        <b:Pages>380-24048</b:Pages>
    </b:Source>
    <b:Source>
        <b:Tag>test4</b:Tag>
        <b:SourceType>InternetSite</b:SourceType>
        <b:Author>
            <b:Author>
                <b:NameList>
                    <b:Person>
                        <b:Last>Tomoji Kanae</b:Last>
                    </b:Person>
                </b:NameList>
            </b:Author>
        </b:Author>
        <b:Title>Test WEBsite</b:Title>
        <b:Year>2026</b:Year>
        <b:Month>04</b:Month>
        <b:Day>21</b:Day>
    </b:Source>
</b:Sources>
```

---
作成者:金栄智治(38024048)  
最終更新日:2026/04/21  
エラーとかでたら教えてください(o_ _)o  