import pandas as pd
import sys

def writeString2File(s,filename):
  f_in = open(filename,'w')
  f_in.write(s)
  f_in.close()

def replaceSpecialCharacters(s_in):
  s_in = s_in.replace("'","")
  s_in = s_in.replace("á","a")
  s_in = s_in.replace("ć","c")
  s_in = s_in.replace("č","c")
  s_in = s_in.replace("ç","c")
  s_in = s_in.replace("Ç","C")
  s_in = s_in.replace("Đ","D")
  s_in = s_in.replace("đ","d")
  s_in = s_in.replace("ë","e")
  s_in = s_in.replace("é","e")
  s_in = s_in.replace("ğ","g")
  s_in = s_in.replace("Ï","I")
  s_in = s_in.replace("ı","i")
  s_in = s_in.replace("ş","s")
  s_in = s_in.replace("Ö","O")
  s_in = s_in.replace("ö","o")
  s_in = s_in.replace("ü","u")
  s_in = s_in.replace("Ü","U")
  print(s_in)
  return s_in

def NTVODataFrame2GOCSVstring(data):
  s = ""
  for index,row in data.iterrows():
    if index < 3:
      continue
    studentNR = row[0]
    completename = row[1]
    splitname = completename.split(", ")
    surname = splitname[0]
    name = splitname[1]
    surname = surname.replace(" ",".")
    name = name.replace(" ",".")
    studentinfo = str(studentNR) + "," + name + ",,," + surname +","+ studentNR+"@hr.nl" + ",,,,,,,," + ("\n" if index < len(data) - 1 else "")
    s = s + studentinfo  
  return replaceSpecialCharacters(s)

if __name__== "__main__" :
  if len(sys.argv)!=2:
    exit("\n!!!Please input arguments correctly.\nfilename.py ntvo_file.xls\nExiting...")
  in_ntvo_file=sys.argv[1]
  out_ntvo_file=in_ntvo_file.split("/")[-1].split(".")[0]+"_2GO.csv"
  data=pd.read_excel(in_ntvo_file,sheet_name='Presentielijst',header = None)
  s = NTVODataFrame2GOCSVstring(data)
  input("\nGoing to create now the csv file. Do not forget to add the classname on the first row of the template. Format below:\nclass_year,class_name\nOK?\n")
  writeString2File(s,out_ntvo_file)
  