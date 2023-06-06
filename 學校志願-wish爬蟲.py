import csv
import bs4
import requests
from bs4 import BeautifulSoup
import pandas

list_of_info=[]
list_of_id=[]

cookie_data = {
    "PHPSESSID":"e4b150rvvi8g2pja5i2gb3k0kt"
}
url = r'http://deptref.clvsc.tyc.edu.tw/studlist.php?'
req = requests.get(url, verify=False,cookies=cookie_data)
htmlpage = BeautifulSoup(req.text)
#print(htmlpage.text)
list_of_data = htmlpage.find_all("li",{'style':'cursor:default;'})
for i in list_of_data:
    list_of_id.append(i.text[:6])

for id in list_of_id:
    print(id)
    url = f'http://deptref.clvsc.tyc.edu.tw/memberlist.php?wish={id}&listtype=1'
    req = requests.get(url,verify=False,cookies=cookie_data)
    htmlpage = BeautifulSoup(req.text)
    list_of_data = htmlpage.find_all('li')
    school_name = str(htmlpage.find('span',{'class':'is-school'}).text)
    school_major = str(htmlpage.find('span',{'class':'has-text-link'}).text)
    school_limit = str(htmlpage.find('span',{'class':'is-note'}).text)
    if school_limit.find('21資管類')!=-1:
        school_major+='(資管類)'
    if school_major[0]=='/':school_major = school_major[1:]
    for i in list_of_data:
        temp = str(i.text).split(' ')
        temp.append(school_name)
        temp.append(school_major)
        temp[0] = temp[0][:3]
        temp[1] = temp[1].replace("＠＠","")+temp.pop(2).replace("(","").replace(")","")
        list_of_info.append(temp)
file = open(r"wish.csv","w",encoding='utf-8')
file_csv = csv.writer(file)
file_csv.writerow(['班級','姓氏與級分','學校名稱','系名'])
file_csv.writerows(list_of_info)
file.close()
read_file = pandas.read_csv (r"wish.csv",encoding='utf-8')
read_file.to_excel (r"wish.xlsx", index = None, header=True)