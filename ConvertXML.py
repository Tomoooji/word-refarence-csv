import os,csv

csvFileName = "Report_Data.csv"
xmlFileName = "Report_Source.xml"

areas = {
    "Book":(10,13),
    "BookSection":(9,13),
    "JournalArticle":(13,18),
    "InternetSite":(29,24)
    }

site_type = 1

with open(os.path.join(os.getcwd(),csvFileName),"r",encoding='utf-8') as data:
    csv_data = [[value for value in row] for row in csv.reader(data)]

with open(os.path.join(os.getcwd(),xmlFileName),"w",encoding='utf-8') as result:
    print('<?xml version="1.0"?>', file=result)
    print('<b:Sources SelectedStyle="" xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">', end="", file=result)
    for source in csv_data[1:]:
        print(f'<b:Source>', end="", file=result)
        
        print(f'<b:Tag>{source[0]}</b:Tag>', end="", file=result)
        print(f'<b:SourceType>{source[1]}</b:SourceType>', end="", file=result)
        if source[2]:
            print(f'<b:Author>', end="", file=result)
            print(f'<b:Author><b:NameList><b:Person><b:Last>{source[2]}</b:Last></b:Person></b:NameList></b:Author>', end="", file=result)
            if source[3]:print(f'<b:Editor><b:NameList><b:Person><b:Last>{source[3]}</b:Last></b:Person></b:NameList></b:Editor>', end="", file=result)
            print(f'</b:Author>', end="", file=result)
        print(f'<b:Title>{source[4]}</b:Title>', end="", file=result)
        print(f'<b:Year>{source[5]}</b:Year>', end="", file=result)
        if source[5]:print(f'<b:Month>{source[6]}</b:Month>', end="", file=result)
        if source[6]:print(f'<b:Day>{source[7]}</b:Day>', end="", file=result)
        
        for idx in range(*areas[source[site_type]]):
            print(f'<b:{csv_data[0][idx]}>{source[idx]}</b:{csv_data[0][idx]}>', end="", file=result)
        
        print('</b:Source>', end="", file=result)
    print('</b:Sources>', end="", file=result)