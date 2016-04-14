import os
import scrapy
from selenium.webdriver.common.action_chains import ActionChains
from scrapy import signals
from scrapy.xlib.pydispatch import dispatcher
from scrapy.http import TextResponse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import xlwt
from ocr.testing import read_captcha
import cv2,urllib
import numpy as np
from Tkinter import *


def url_to_image(url):
    # download the image, convert it to a NumPy array, and then read
    # it into OpenCV format
    resp = urllib.urlopen(url)
    image = np.asarray(bytearray(resp.read()), dtype="uint8")
    image = cv2.imdecode(image, cv2.IMREAD_COLOR)
 
    # return the image
    return image

class Result(scrapy.Spider):
    name = "btech"
    s_roll = 1309110001
    e_roll = None
    count = 1
    top = [["",0],["",1000]]
    allowed_domains = ['http://new.aktu.co.in/']
    start_urls = ['http://new.aktu.co.in/']

    def fetch(self,entries):
        sroll = entries[1]
        self.s_roll = int(entries[0][1].get())
        self.e_roll  = int(entries[1][1].get())
        print self.s_roll+1
        self.root.destroy()

    def makeform(self,fields):
       entries = []
       for field in fields:
          row = Frame(self.root)
          lab = Label(row, width=15, text=field, anchor='w')
          ent = Entry(row)
          row.pack(side=TOP, fill=X, padx=5, pady=10)
          lab.pack(side=TOP)
          ent.pack(side=TOP, expand=YES, fill=X , pady=10)
          entries.append((field, ent))
       return entries

    def __init__(self, filename=None):
        fields = 'First Roll No.' ,'Last Roll No.'
        self.root = Tk()
        self.root.title('Result')
        self.root.geometry("200x200")
        ents = self.makeform( fields)
        self.root.bind('<Return>', (lambda event, e=ents: self.fetch(e)))   
        b1 = Button(self.root, text='submit',
                command=(lambda e=ents: self.fetch(e)))
        b1.pack(side=TOP, padx=5, pady=5)
        self.root.mainloop()
        self.driver = webdriver.Chrome()
        #self.driver = webdriver.Firefox()
        self.workbook = xlwt.Workbook()
        self.sheet = self.workbook.add_sheet('Sheet_1')
        dispatcher.connect(self.spider_closed, signals.spider_closed)

    def spider_closed(self, spider):
        self.driver.close()
        self.workbook.save('result.xls')

    def add_in_sheet(self,item):
        try :
            
            self.count += 1
            if self.count == 2:
                self.sheet.write(0,0,'Name')
                self.sheet.write(0,1,"Father's Name")
                self.sheet.write(0,2,'Roll No.')
                self.sheet.write(0,3,'Enrollment No.')
                self.sheet.write(0,4,'Branch')
                self.sheet.write(0,5,'College')
                self.sheet.write(0,6,item['s1'])
                self.sheet.write(0,7,item['s2'])
                self.sheet.write(0,8,item['s3'])
                self.sheet.write(0,9,item['s4'])
                self.sheet.write(0,10,item['s5'])
                self.sheet.write(0,11,item['s6'])
                self.sheet.write(0,12,'GP')
                self.sheet.write(0,13,'Total')
            self.sheet.write(self.count,0,item['name'])
            self.sheet.write(self.count,1,item['father'])
            self.sheet.write(self.count,2,item['roll'])
            self.sheet.write(self.count,3,item['enroll'])
            self.sheet.write(self.count,4,item['branch'])
            self.sheet.write(self.count,5,item['clg'])
            self.sheet.write(self.count,6,item[item['s1']])
            self.sheet.write(self.count,7,item[item['s2']])
            self.sheet.write(self.count,8,item[item['s3']])
            self.sheet.write(self.count,9,item[item['s4']])
            self.sheet.write(self.count,10,item[item['s5']])
            self.sheet.write(self.count,11,item[item['s6']])
            self.sheet.write(self.count,12,item['gp'])
            self.sheet.write(self.count,13,item['tot'])
            if int(item['tot']) > self.top[0][1]:
                self.top[0][1] = int(item['tot'])
                self.top[0][0] = item['name']
            if int(item['tot']) < self.top[1][1]:
                self.top[1][1] = int(item['tot']) 
                self.top[1][0] = item['name']
            
        except :
            return
        

    def parse_result(self, response):
        try :
            item = {}
            # Load the current page into Selenium
            # self.driver.get(response)
            try:
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_imgstud"]')))
            except TimeoutException:
                print "result not found"
                input("result page")
                return 0
            # Sync scrapy and selenium so they agree on the page we're looking at then let scrapy take over
            resp = TextResponse(url=self.driver.current_url, body=self.driver.page_source, encoding='utf-8');
            temp = format(resp.xpath('//*[@id="lblname"]/text()').extract())
            item['name'] = temp[3:-2]

            temp = format(resp.xpath('//*[@id="lblfname"]/text()').extract())
            item['father'] = temp[3:-2]

            temp = format(resp.xpath('//*[@id="lblrollno"]/text()').extract())
            item['roll'] = temp[3:-2]

            temp = format(resp.xpath('//*[@id="lblenrollno"]/text()').extract())
            item['enroll'] = temp[3:-2]

            temp = format(resp.xpath('//*[@id="lblbranch"]/text()').extract())
            item['branch'] = temp[3:-2]

            temp = format(resp.xpath('//*[@id="lblcollegename"]/text()').extract())
            item['clg'] = temp[3:-2]

            temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/b/text()').extract())
            item['s1'] = temp[3:-2]

            t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]/b/text()').extract())
            t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[4]/b/text()').extract())
            item[item['s1']] = t1[3:-2] + ' , '  + t2[3:-2]

            temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/b/text()').extract())
            item['s2'] = temp[3:-2]
            t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[3]/b/text()').extract())
            t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[4]/b/text()').extract())
            item[item['s2']] = t1[3:-2] + ' , '  + t2[3:-2]

            temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/b/text()').extract())
            item['s3'] = temp[3:-2]
            t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[3]/b/text()').extract())
            t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[4]/b/text()').extract())
            item[item['s3']] = t1[3:-2] + ' , '  + t2[3:-2]    

            temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/b/text()').extract())
            item['s4'] = temp[3:-2]
            t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td[3]/b/text()').extract())
            t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td[4]/b/text()').extract())
            item[item['s4']] = t1[3:-2] + ' , '  + t2[3:-2]

            temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[6]/td[2]/b/text()').extract())
            item['s5'] = temp[3:-2]
            t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[6]/td[3]/b/text()').extract())
            t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[6]/td[4]/b/text()').extract())
            item[item['s5']] = t1[3:-2] + ' , '  + t2[3:-2]

            temp = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]/b/text()').extract())
            item['s6'] = temp[3:-2]
            t1 = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[7]/td[3]/b/text()').extract())
            t2  = format(resp.xpath('//*[@id="Pane0_content"]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[7]/td[4]/b/text()').extract())
            item[item['s6']] = t1[3:-2] + ' , '  + t2[3:-2]

            temp = format(resp.xpath('//*[@id="ctl00_ContentPlaceHolder1_tr1"]/td[3]/text()').extract())
            item['gp'] = temp[3:-2]
            temp = format(resp.xpath('//*[@id="Pane0_content"]/table[3]/tbody/tr[2]/td[3]/text()').extract())
            print temp[5:-7]
            item['tot'] = temp[5:-7]
            self.add_in_sheet(item)
            return 1
        except :          
            return 1

    def fill_captcha(self,resp):
        captcha_url = format(resp.xpath('//*[@id="ctl00_ContentPlaceHolder1_divSearchRes"]/center/table/tbody/tr[4]/td/center/div/div/img/@src').extract())
        url = "http://new.aktu.co.in/" + captcha_url[3:-2]
        captcha_value = read_captcha(url_to_image(url))
        print captcha_value
        captcha_input = self.driver.find_element_by_name('ctl00$ContentPlaceHolder1$txtCaptcha')
        captcha_input.clear()
        captcha_input.send_keys(captcha_value)
        submit = self.driver.find_element_by_name('ctl00$ContentPlaceHolder1$btnSubmit')
        actions = ActionChains(self.driver)
        time.sleep(3)
        actions.click(submit)
        actions.perform()
        resp = TextResponse(url=self.driver.current_url, body=self.driver.page_source, encoding='utf-8');
        return resp

    def parse(self, response):
        try :
            while self.s_roll <= self.e_roll:
                self.driver.get('http://new.aktu.co.in/')
                try:
                    WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ContentPlaceHolder1_divSearchRes"]/center/table/tbody/tr[4]/td/center/div/div/img')))
                except:
                    continue
    	        # Sync scrapy and selenium so they agree on the page we're looking at then let scrapy take over
                resp = TextResponse(url=self.driver.current_url, body=self.driver.page_source, encoding='utf-8');
                rollno = self.driver.find_element_by_name('ctl00$ContentPlaceHolder1$TextBox1')
                rollno.send_keys(self.s_roll)
                try :
                    resp = self.fill_captcha(resp)
                    print format(resp.xpath('//*[@id="ContentPlaceHolder1_Label1"]/text()').extract())
                    while "Incorrect" in format(resp.xpath('//*[@id="ContentPlaceHolder1_Label1"]/text()').extract()):
                        resp = self.fill_captcha(resp)
                except :
                    continue
                self.parse_result(self.driver.current_url)
                self.s_roll += 1
            self.count +=3
            self.sheet.write(self.count,0,"First")
            self.sheet.write(self.count,1,self.top[0][0])
            self.sheet.write(self.count+1,0,"Last")
            self.sheet.write(self.count+1,1,self.top[1][0])
        except :
            self.parse(response)
        finally :
            return

