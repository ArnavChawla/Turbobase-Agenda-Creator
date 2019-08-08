from __future__ import print_function
from flask import Flask
from os import environ
from flask import request, render_template, send_file
import requests
from test import MailMerge
import time
from bs4 import BeautifulSoup as soup
from datetime import date
import os
import datetime
import calendar
import json
from PyDictionary import PyDictionary
abbr_to_num = {name: num for num, name in enumerate(calendar.month_name) if num}
class date():
    def __init__(self,s):
        a = s.split(' ')
        self.day = a[0]
        self.month = a[1]
        self.year = a[2].split(',')[0]
    def toSring(self):
        print("{} {} {}".format(self.day,self.month,self.year))
        return("{} {} {}".format(self.day,self.month,self.year))
class combined():
    def __init__(self,url,date):
        self.date= date.toSring()
        self.url = "http://www.turbobase.com"+url
app = Flask(__name__)
@app.route('/')
def my_form():
    combarr = []
    start = datetime.datetime.now()
    session = requests.session()
    r = session.get("http://www.turbobase.com/talk")
    print(session.cookies["JSESSIONID"])
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'en-US,en;q=0.9',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Content-Length': '88',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': 'JSESSIONID='+session.cookies["JSESSIONID"],
        'Host': 'www.turbobase.com',
        'Origin': 'http://www.turbobase.com',
        'Referer': 'http://www.turbobase.com/talk/login',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36'
    }
    data = {
        'formEmpty': 'true',
        'formUsed': 'Login',
        'emailAddress': '',
        'password': ''
    }
    r = session.post(r.url,headers=headers,data=data,cookies=session.cookies)
    r = session.get("http://www.turbobase.com/talk/meetings/list")
    bs = soup(r.text,"html.parser")
    data = []
    table = bs.find_all('table')[1]
    table_body = table.find('table')
    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('a')
        cols = [ele["href"] for ele in cols]
        data.append([ele for ele in cols if ele])
    del data[0]
    urls = []
    for item in reversed(data):
        urls.append(item[0])
    data=[]
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        data.append([ele for ele in cols if ele])
    del(data[0])
    dates = []
    for item in reversed(data):
        dates.append(item[1])
    datearr = []
    for item in dates:
        datearr.append(date(item))

    i = len(datearr)-1
    for item in reversed(datearr):
        #print(item.year)
        if(int(item.year) < 2019):
            datearr.pop(i)
            urls.pop(i)
        i-=1
    i = len(datearr)-1
    for item in reversed(datearr):
        #print(item.month)
        if(abbr_to_num[str(item.month)] < start.month):
            datearr.pop(i)
            urls.pop(i)
        i-=1
    i = len(datearr)-1
    for item in reversed(datearr):
        #print(item.day)

        if(int(item.day) < int(start.day)):
            print(item.day)
            if(abbr_to_num[str(item.month)] == start.month):
                datearr.pop(i)
                urls.pop(i)
        i-=1
    #print(len(datearr) == len(urls))
    count = 0
    i = 0
    for i in range(len(datearr)):
        if count < 4:
            combarr.append(combined(urls[i],datearr[i]))
            count+=1
    return render_template("my-form.html", data=combarr)
@app.route('/', methods=['POST'])
def my_form_post():
    template = os.getcwd()+"/template.docx"
    document = MailMerge(template)
    print(request.form)
    url = request.form["date"]
    session = requests.session()
    r = session.get("http://www.turbobase.com/talk")
    print(session.cookies["JSESSIONID"])
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'en-US,en;q=0.9',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Content-Length': '88',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': 'JSESSIONID='+session.cookies["JSESSIONID"],
        'Host': 'www.turbobase.com',
        'Origin': 'http://www.turbobase.com',
        'Referer': 'http://www.turbobase.com/talk/login',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36'
    }
    data = {
        'formEmpty': 'true',
        'formUsed': 'Login',
        'emailAddress': '',
        'password': ''
    }
    r = session.post(r.url,headers=headers,data=data,cookies=session.cookies)
    r = session.get(url)
    bs = soup(r.text,"html.parser")
    data = []
    table = bs.find_all('table')[1]
    table_body = table.find('table')
    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        data.append([ele for ele in cols if ele])

    i = 0
    roles = []
    print(data)
    for i in range(0,len(data)):
        st = str(data[i])
        b = st.split("\',")
        x = str(b[0]).split('[\'')

        try:
            roles.append(str(b[1]).split('\'')[1].split('\']')[0])
            print("++")
        except:
            roles.append("")
            print('asd')
        print(len(roles))
    document.merge(TITLE=roles[1],date=roles[0],office=roles[19],fat=roles[20],toastmaster=roles[4],tabletopics=roles[5],gramarian=roles[7],wizard=roles[21],timer=roles[8],general=roles[6],speaker1=roles[9],speaker2=roles[10],speaker3=roles[11],speaker4=roles[12],eval1=roles[14],eval2=roles[15],eval3=roles[16],eval4=roles[17])
    document.write(os.getcwd()+"/files/{}.docx".format(roles[0].replace(" ","").split(",")[0]))
    time.sleep(0.05)
    return send_file(os.getcwd()+"/files/{}.docx".format(roles[0].replace(" ","").split(",")[0]),as_attachment=True)

@app.route('/gramarian')
def here():
    return render_template("gramrian.html")

@app.route("/createslide", methods=["POST"])
def bet():
    template = os.getcwd()+"/grammartemp.docx"
    document=MailMerge(template)
    print(request.form)
    word = request.form["input"]
    dictionary=str(PyDictionary.googlemeaning(word))
    dictionary = dictionary.split('\n')
    document.merge(WORD=word.upper(),part=dictionary[0].split(": ")[1].rstrip(),defo=dictionary[1].split(".")[0])
    document.write(os.getcwd()+"/files/gramarian.docx")

    return send_from_directory(directory=os.getcwd()+"/files", filename='gramarian.docx',as_attachment=True)


if __name__ == '__main__':
    app.run()
