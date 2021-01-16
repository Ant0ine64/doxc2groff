#!/bin/python3

import docx

dic = {
        'heading1': ".NH",
        'heading2': ".NH 2",
        'heading2': ".NH 3",
        'paragraph': ".PP"
        }

document = docx.Document("input.docx")
outfile = open("output.ms", "w")

for para in document.paragraphs:
    if para.style.name=='Heading 1':
        print("titre : " + para.text)
        outfile.write(dic['heading1'] + "\n" + para.text + "\n")
    elif para.style.name=='Heading 2':
        print("titre2 : " + para.text)
        outfile.write(dic['heading2'] + "\n" + para.text + "\n")
    else:
        print(para.text)
        outfile.write(dic['paragraph'] + "\n" + para.text + "\n")

outfile.close()
