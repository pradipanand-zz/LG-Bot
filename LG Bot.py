from docx import Document
import pandas as pd
import datetime,time#,xlsxwriter

doc_loc = 'C:/Python/Project/excel/'
source_doc = 'sample_doc.docx'
input_excel = 'read_data.xlsx'
destination_doc = 'name_authorization_letter.doc' #this is just format of document
excel_list = pd.DataFrame()

def create_doc(document,string1,name,string2,string3,string4,position,location,Address,Letter_created):
    #print('====>',string3)
    for p in document.paragraphs:

        if 'string1' in p.text:
            inline = p.runs
            for i in range(len(inline)):
                #print('i catch1')
                if 'string1' in inline[i].text:
                    #print('i catch2')
                    text = inline[i].text.replace('string1', string1)
                    inline[i].text = text
                    break

        if 'name' in p.text:
            inline = p.runs
            for i in range(len(inline)):
                #print('i catch1')
                if 'name' in inline[i].text:
                    #print('i catch2')
                    text = inline[i].text.replace('name', name)
                    inline[i].text = text
                    break
                    #inline[i].text = re.sub(r"[^\u0900-\u097F]+", "", text)
        if 'string2' in p.text:
            inline = p.runs
            for i in range(len(inline)):
                #print('i catch1')
                if 'string2' in inline[i].text:
                    #print('i catch2')
                    text = inline[i].text.replace('string2', string2)
                    inline[i].text = text
                    break

        if 'string3' in p.text:
            inline = p.runs
            for i in range(len(inline)):
                #print('i catch1')
                if 'string3' in inline[i].text:
                    #print('i catch2')
                    text = inline[i].text.replace('string3', string3)
                    inline[i].text = text
                    break

        if 'string4' in p.text:
            inline = p.runs
            for i in range(len(inline)):
                #print('i catch1')
                if 'string4' in inline[i].text:
                    #print('i catch2')
                    text = inline[i].text.replace('string4', string4)
                    inline[i].text = text
                    break

        if 'position' in p.text:
            inline = p.runs
            for i in range(len(inline)):
                #print('i catch1')
                if 'position' in inline[i].text:
                    #print('i catch2')
                    text = inline[i].text.replace('position', position)
                    inline[i].text = text
                    break

        if 'location' in p.text:
            #print('oooooooooooooo',location)
            #print('i catch1')
            inline = p.runs
            for i in range(len(inline)):
                #print('i catch2',i)
                if 'location' in inline[i].text:
                    #print('i catch3')
                    text = inline[i].text.replace('location', location)
                    inline[i].text = text
                    break

        if 'Address' in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if 'Address' in inline[i].text:
                    #print('i catch2')
                    text = inline[i].text.replace('Address', Address)
                    inline[i].text = text
                    break

        if 'Letter_created' in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if 'Letter_created' in inline[i].text:
                    #print('i catch2')
                    text = inline[i].text.replace('Letter_created', Letter_created)
                    inline[i].text = text
                    break
    return document

data                = pd.read_excel(doc_loc + input_excel,encoding='utf-8')
length_of_record    = len(data)
#get header into dataframe
#header_list = pd.read_excel(doc_loc + input_excel,nrows=0)
header_list         = ['string1', 'name', 'string2', 'string3', 'string4', 'position', 'location', 'Address', 'Letter_created','date']
list_of_variable = []

for p in range(0,length_of_record):
    document            = Document(doc_loc+source_doc)
    string1             = data.iloc[p]['string1']
    name                = data.iloc[p]['name'] #for file name
    string2             = data.iloc[p]['string2']
    string3             = data.iloc[p]['string3']
    string4             = data.iloc[p]['string4']
    position            = data.iloc[p]['position']
    location            = data.iloc[p]['location']
    Address             = data.iloc[p]['Address']
    Letter_created      = data.iloc[p]['Letter_created']
    today               = data.iloc[p]['date']

    if data.iloc[p]['Letter_created'] != 'Y':
        document = create_doc(document,string1,name,string2,string3,string4,position,location,Address,Letter_created)
        #today=datetime.datetime.today().strftime('%m-%d-%Y-%H-%M-%S')
        today=datetime.datetime.today().strftime('%m-%d-%Y')
        document.save(doc_loc + name + '_authorization_letter_'+ today + '.docx')
        #below creating tp update if record has been processed
        Letter_created      = 'Y'
        #print('time-',today)
        print(p+1,'-doc is completed')

    list_of_field       = [string1,name,string2,string3,string4,position,location,Address,Letter_created,today]
    list_of_variable.append(list_of_field)

df=pd.DataFrame(list_of_variable,columns=header_list)
df.to_excel(doc_loc + input_excel,index=False)

print('All documents have been completed')

time.sleep(20)
