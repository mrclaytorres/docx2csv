import os
import os.path
import docx
import pandas as pd
import datetime
import time
import csv

paragraph_rows = []

def para2text(p):
  rs = p._element.xpath('.//w:t')
  return u" ".join([r.text for r in rs])

def getText(paperFilePath):
  doc = docx.Document(paperFilePath)
  for para in doc.paragraphs:
    paragraph_rows.append(para2text(para))


def doc2csv(input_file: str, output_file: str):

  time_start = datetime.datetime.now().replace(microsecond=0)
  directory = os.path.dirname(os.path.realpath(__file__))

  paperFilePath = input_file

  # Extract the document file
  getText(paperFilePath)

  # save the paragraphs into a csv file
  now = datetime.datetime.now().strftime('%Y%m%d-%Hh%M')
  print('Saving to a CSV file...\n')
  data = {"Paragraphs": paragraph_rows}
  df=pd.DataFrame(data=data)
  df.index+=1

  filename = output_file + now + ".csv"

  print(f'{filename} saved sucessfully.\n')

  file_path = os.path.join(directory,'csvfiles/', filename)
  df.to_csv(file_path)

  time_end = datetime.datetime.now().replace(microsecond=0)
  runtime = time_end - time_start
  print(f"Script runtime: {runtime}.\n")

if __name__ == '__main__':
  import sys
  input_file = sys.argv[1]
  output_file = sys.argv[2]
  doc2csv(input_file, output_file)