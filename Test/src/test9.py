'''
Created on May 5, 2014

@author: teresatan
'''

from bs4 import BeautifulSoup as bs
from scrapy.contrib.spiders import CrawlSpider, Rule
from scrapy.contrib.linkextractors.sgml import SgmlLinkExtractor
from scrapy.selector import HtmlXPathSelector
from scrapy.http import Request
from scrapy.spider import Spider


import urllib2
import csv
import os
import re
import math
import numpy
import datetime
import json
import requests
import xlwt
import xlrd
import parsedatetime as pdt
#from IPython import embed

class yahooSpider(CrawlSpider):
    name = "yahooSpider"
    allowed_domains = ['yahoo.com']






#workbook = xlrd.open_workbook('input.xls')
#sheet = workbook.sheet_by_index(0)
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Sheet1")   #sheet.write(row, column, 'value')
sheet.write(0,0,'Name')
sheet.write(0,1,'Comment')
sheet.write(0,2,'Date')
sheet.write(0,3,'Thumbs-up')
sheet.write(0,4,'Thumbs-down')
#from IPython import embed

#Current time (hour)
time = datetime.datetime.now()

s = requests.session()

#####Number of total comments######
url_start='http://news.yahoo.com/_xhr/contentcomments/get_comments/'
params_start = {'content_id':'2c4772d7-9388-33e0-989d-3f0e4f6558bb',
          '_device':'full',
          'sortBy':'latest',
          'isNext':'true',
          '_media.modules.content_comments.switches._enable_view_others':'1',
          '_media.modules.content_comments.switches._enable_mutecommenter':'1',
          'enable_collapsed_comment':'1'}             


#####################Gives the total amount of comments for this article##########################
commentCount = s.get(url_start, params=params_start)
obj_count = commentCount.content
allcomment_count = json.loads(obj_count)
totalCommentsCount = allcomment_count["totalCommentCount"]
#print totalCommentsCount


url='http://news.yahoo.com/_xhr/contentcomments/get_comments/'
params = {'content_id':'2c4772d7-9388-33e0-989d-3f0e4f6558bb',
          '_device':'full',
          'count':'100',
          'sortBy':'latest',
          'isNext':'true',
          #'pageNumber':'2',
          #'data-page': '2',
          '_media.modules.content_comments.switches._enable_view_others':'1',
          '_media.modules.content_comments.switches._enable_mutecommenter':'1',
          'enable_collapsed_comment':'1'}             



#m = re.search('(?=id\":','content_id=2c4772d7-9388-33e0-989d-3f0e4f6558bb&amp;alias_id=story%3Dpistorius-allegedly-made-sinister-remark-court-134853464--spt\" class=\"rapid-noclick-resp\" role=\"link\" title=\"Email to friends\" data-share')















comments_html = s.get(url, params=params)
#print (comments_html.url)


obj = comments_html.content


#####All comments including HTML#####
allcomments = json.loads(obj)

commentList_json = allcomments["commentList"]
#print commentList_json

soup = bs(commentList_json)
#print soup.prettify()

timestampTag = soup.findAll("span", {"class" : "comment-timestamp"})
#print timestampTag
commentTag = soup.findAll("p", {"class" : "comment-content "})
#print commentTag
thumbsupTag = soup.findAll("div", {"class" : "int vote-box up"})
#print thumbsupTag
thumbsdownTag = soup.findAll("div", {"class" : "int vote-box down"})
#print thumbsdownTag
nameTag = soup.findAll("span", {"class" : "int profile-link "})
#print nameTag


##################################### THIS PRINTS OUT ALL THE COMMENTS!!!! ########################################
comment_array = []
time_array = []
all_array = []


####################comments and time ###########################
row_num = 1
###### Writes comments into excel sheet #######
for com in commentTag:
    sheet.write(row_num,1,''.join(com.find(text=True)).strip())
    row_num = row_num + 1 


row_num = 1
###### Writes time into excel sheet #######
for tim in timestampTag:
    #sheet.write(row_num,2,''.join(tim.find(text=True)).strip())
    time_ago = ''.join(tim.find(text=True)).strip()
    #print time_ago
    cal = pdt.Calendar()
    #print cal
    newtime = cal.parse(time_ago)
    newtime = str(newtime)
    #n_strip = newtime.strip()
    yr = re.search('(?<=tm_year=)(\d+)', newtime)
    mth = re.search('(?<=tm_mon=)(\d+)', newtime)
    dy = re.search('(?<=tm_mday=)(\d+)', newtime)
    #print yr.group(0)
    #print mth.group(0)
    #print dy.group(0)
    #print mth.group(0) + "/" + dy.group(0) + "/" + yr.group(0)
    dte = mth.group(0) + "/" + dy.group(0) + "/" + yr.group(0)
    sheet.write(row_num,2,dte)
    #newtime = time.struct_time(time_ago)
    #print cur_time
    #print newtime
    #newtime.tm_mon + "/" + newtime.tm_day + "/" + newtime.tm_year
    row_num = row_num + 1
    #break


row_num = 1
###### Writes the number of thumbs-up into excel sheet#########
for thumbsup in thumbsupTag:
    #print ''.join(thumbsup.findAll(text=True)).strip()
    sheet.write(row_num,3,''.join(thumbsup.findAll(text=True)).strip())
    row_num = row_num + 1


row_num = 1
###### Writes the number of thumbs-down into excel sheet#########
for thumbsdown in thumbsdownTag:
    #print ''.join(thumbsup.findAll(text=True)).strip()
    sheet.write(row_num,4,''.join(thumbsdown.findAll(text=True)).strip())
    row_num = row_num + 1


row_num = 1
###### Writes the names of the commentors into excel sheet#########
for names in nameTag:
    #print ''.join(names.findAll(text=True)).strip()
    sheet.write(row_num,0,''.join(names.findAll(text=True)).strip())
    row_num = row_num + 1
    
workbook.save("excel_files/yahoo.xls") 