import csv
import requests
from bs4 import BeautifulSoup
import pandas

list_of_info=[]
cookie_data = {
    "PHPSESSID":"e4b150rvvi8g2pja5i2gb3k0kt"
}
for score in range(76):
    url = f"http://deptref.clvsc.tyc.edu.tw/myscorelist.php?myscore={score}&listtype=1"
    print(url)
    req = requests.get(url, verify=False,cookies=cookie_data)
    htmlpage = BeautifulSoup(req.text)
    list_of_data = htmlpage.find_all("h1",{'class':'title is-5 has-text-link'})

    for i in list_of_data:
        temp = i.text.replace("/ ","")
        if(temp.find('21資管類')!=-1):
            temp = temp.replace(' 21資管類/','(資管類) ').split(' ')
        else:
            temp = temp.split(' ')
        if(len(temp)==6):
           temp[3]+=temp.pop(4)
        temp.append(score)
        temp[4] = temp[4].replace("人","")
        list_of_info.append(temp)
file = open(r"myscore.csv","w",encoding='utf-8')
#file = open(r"C:\Users\jimmy\OneDrive\桌面\Python\Python爬蟲\myscore.csv","w",encoding='utf-8')
file_csv = csv.writer(file)
file_csv.writerow(['代號','校名','系名','標準','考取級分'])
file_csv.writerows(list_of_info)
file.close()
read_file = pandas.read_csv (r"myscore.csv",encoding='utf-8')
read_file.to_excel (r"myscore.xlsx", index = None, header=True)