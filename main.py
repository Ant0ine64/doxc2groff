#!/bin/python3

import docx

def change_align(align):
    outfile.write(dic['align_start'] + " " + align
            + "\n" + para.text + "\n"
            + dic['align_end'] + "\n")

def parse_paragraph_run(run):
    if (run.bold == True and run.italic == True):
        #write blod & italic
        return "\n" + dic['bold_italic'] + " \"" + run.text + "\"\n"
    elif run.bold == True:
        #wite bold
        return "\n" + dic['bold'] + " \"" + run.text + "\"\n"
    elif run.italic == True:
        #wite italic
        return "\n" + dic['italic'] + " \"" + run.text + "\"\n"
    elif run.underline == True:
        #write underline
        return "\n" + dic['underline'] + " \"" + run.text + "\"\n"
    else:
        #normal text
        return run.text


dic = {
        'title': ".TL",
        'heading1': ".NH",
        'heading2': ".NH 2",
        'heading2': ".NH 3",
        'paragraph': ".PP",
        'align_start': ".DS",
        'align_end': ".DE",
        'bold': ".B",
        'italic': ".I",
        'bold_italic': ".BI",
        'underline': ".UL"
        }

document = docx.Document("input.docx")
outfile = open("output.ms", "w")

for para in document.paragraphs:
    #test title
    if para.style.name=='Heading 1':
        # Happens once
        #print("titre : " + para.text)
        outfile.write(dic['title'] + "\n" + para.text + "\n")
    elif para.style.name=='Heading 2':
        #print("titre2 : " + para.text)
        outfile.write(dic['heading1'] + "\n" + para.text + "\n")
    elif para.style.name=='Heading 3':
        #print("titre2 : " + para.text)
        outfile.write(dic['heading2'] + "\n" + para.text + "\n")
    #test paragraph
    else:
        if para.alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER:
            #center
            change_align("C")
        elif para.alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT:
            #right
            change_align("R")
        else:
            #left normal
            print(para.text)
            text = ""
            for run in para.runs:
                text += parse_paragraph_run(run)
                #print("run : " + text)
            print(text)
            outfile.write(dic['paragraph'] + "\n" + text + "\n")

outfile.close()
