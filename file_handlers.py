#from urllib.parse import unquote
from rtf.Rtf2Markdown import getMarkdown
import olefile
import sys
import chardet
import json

def get_notes(sticky_notes_file_path):
    notes = []

    snt_file = olefile.OleFileIO(sticky_notes_file_path)

    for storage in snt_file.listdir(storages=True, streams=False):
        note_id = storage[0]  # UUID-like string representing the note ID
        note_text_rtf_file = '0'  # RTF content of the note

        with snt_file.openstream([note_id, note_text_rtf_file]) as note_content:
            rawdata = note_content.read()
            encoding = chardet.detect(rawdata)
            #print(encoding)
            note_text_rtf = rawdata.decode('ascii')
            #note_text_rtf = rawdata.decode('utf-8')
        #print(note_text_rtf)
        notes.append({'text': getMarkdown(note_text_rtf), 'color': None})

    snt_file.close()

    return notes

sn_path = sys.argv[1]
with open("notes.json","w") as wp:
    vals = get_notes(sn_path)
    snotes = json.dumps(vals,indent=4)
    print(snotes)
    wp.write(snotes)

with open("notes.txt","w") as wp:
    vals = get_notes(sn_path)
    for note in vals:
        print(note["text"])
        text = note["text"]
        wp.write(text)
        wp.write("------------")
        
