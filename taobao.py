import requests
import json
import matplotlib.pyplot as plt
from wordcloud import WordCloud, ImageColorGenerator
import PIL.Image as Image
import numpy as np
import jieba
import xlwt
import sys
import os
import shutil
import subprocess


Target_File = ""
# 这里需要替换为你自己的 Cookie
YourCookie = "lid=yourname; cna=E7bFFGrA/AQCAQ4RFiE1VDBd; hng=CN%7Czh-CN%7CCNY%7C156; enc=tF7ZZdQryuVJKqmAq8JceFK8I5zl6RMzQculbb4d3gT7zBSJPLddOJBd2RE3Zko7hyoZqjKjmXPcrpXr1H9CKw%3D%3D; tfstk=cWodBeZlbdBL4uMUu0LgVEuPKtmdZb8TawNc2Ltq9_7GlqjRiTmmD0ckd-txDdC..; sgcookie=EL0bhCbSgJwtuQtJTJwUS; t=6f5e39f76283b43de0d478cfb81df1af; uc3=id2=Vyh%2FYV2ZAZQY&nk2=F4Ik0bqkgZs%3D&lg2=URm48syIIVrSKA%3D%3D&vt3=F8dBxGR%2Fgrbh5Qb%2BbQg%3D; tracknick=yourname; uc4=id4=0%40VX9OrQst308plZSgDdMhds7UpSI%3D&nk4=0%40FZi2hHZcZfFJvOx8uBFMu%2BF0ww%3D%3D; lgc=yourname; _tb_token_=e891f53510533; cookie2=102a46c5481ed4e45882034fd2898956; x5sec=7b22726174656d616e616765723b32223a226532636566323936386130613237353962643732363964646363373438356333434a695975665546454b7267333879692f616d6f7641453d227d; l=eBx5Z9RPqfa2QNNhBOfanurza77OvIRXjuPzaNbMiT5POw56PWKFWZje0LYBCnhVnsCDR37EODnaBvTUZyUihZXRFJXn9MptndLh.; isg=BCoqh57FAEuc1Y9BxM0GxDlge5bMm671Y6wpkLTj6X0M58uhnCocBdFRdhN7PCaN"

# 获取字符串长度，一个中文的长度为2
def len_byte(value):
    length = len(value)
    utf8_length = len(value.encode('utf-8'))
    length = (utf8_length - length) / 2 + length
    return int(length)*256

def getMaxLength(value, oldLength):
    return max(oldLength, len_byte(value))

class TaoBao:
    lastPage = 1
    url="https://rate.tmall.com/list_detail_rate.htm"
    header={
        "cookie":YourCookie,
        "referer":"https://detail.tmall.com/item.htm",
        "user-agent":"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 UBrowser/6.2.4098.3 Safari/537.37",
    }
    params={                                #必带信息
        "itemId":"",                       #商品id  
        "sellerId":"",
        "currentPage":"",                  #页码
        "order":"3",                        #排序方式:1:按时间降序  3：默认排序
        "callback":"jsonp775",
        }
    def __init__(self, id:str, sellerId:str):
        self.params['itemId']=id
        self.params['sellerId']=sellerId

    def getPageData(self,pageIndex:int):
        self.params["currentPage"]=str(pageIndex)
        print("正在请求第 {} 页评论".format(pageIndex))
        req=requests.get(self.url,self.params,headers=self.header,timeout = 2).content.decode('utf-8');     #解码，并且去除str中影响json转换的字符（\n\rjsonp(...)）;
        req=req[req.find('{'):req.rfind('}')+1]
        ret = json.loads(req);
        return ret;

    def setOrder(self,way:int):
        self.params["order"]=way;

    def getAllData(self):
        Data=self.getPageData(1)
        self.lastPage= Data['rateDetail']['paginator']['lastPage']
        rateCount = Data['rateDetail']['rateCount']['total']
        print("一共有 {} 条评论，共计 {} 页".format(rateCount, self.lastPage))
        for i in range(2,self.lastPage+1):
            Data['rateDetail']['rateList'].extend(self.getPageData(i)['rateDetail']['rateList'])
        return Data;

    def simplifyData(self, data):
        ret = []
        for comment in data['rateDetail']['rateList']:
            item = {};
            item['commentTime'] = comment['rateDate']
            item['comment'] = comment['rateContent']
            item['type'] = comment['auctionSku']
            item['appendComment'] = ''
            if comment['appendComment']:
                item['appendComment'] = comment['appendComment']['content']
            ret.append(item);
            
        return ret;

    def analyze(self, data):
        #导入停用词
        stoplist = [line.strip() for line in open('stop.txt','r',encoding="utf8").readlines() ]
        #打开本地数据文件，注意文件名不能用中文
        # text_from_file_with_apath = open('new.txt',encoding="utf8").read()
        text_from_file_with_apath = data
        #去除停用词
        for stop in stoplist:
            jieba.del_word(stop)
        
        #用jieba模块进行挖掘关键词
        wordlist_after_jieba = jieba.cut(text_from_file_with_apath, cut_all=False)
        wl_space_split = " ".join(wordlist_after_jieba)
        
        coloring=np.array(Image.open("cat_new.jpg"))
        my_wordcloud = WordCloud(background_color="white",
                                 mask=coloring,
                                 width=2000, height=1000,
                                 font_path="Hiragino Sans GB.ttc",
                                 max_words=400,
                                 max_font_size=80,
                                 min_font_size=10,
                                 random_state=40).generate(wl_space_split)
        # plt.imshow(my_wordcloud)
        # plt.axis("off")
        # plt.show()
        my_wordcloud.to_file(Target_File+'/WordCloud.png');

    def toExcel(self, data):
        wb = xlwt.Workbook()
        # 添加一个表
        ws = wb.add_sheet('test')
        style = xlwt.XFStyle()
        style.alignment.wrap = 1

        line = 0;
        # 3个参数分别为行号，列号，和内容
        # 需要注意的是行号和列号都是从0开始的
        ws.write(line, 0, '时间')
        ws.write(line, 1, '评论')
        ws.write(line, 2, '追加评论')
        ws.write(line, 3, '分类')

        time_len = 10*256
        comment_len = 10*256
        appendComment_len = 10*256
        itemType_len = 10*256

        for item in data:
            line += 1

            time_len = max(len_byte(item['commentTime']), time_len)
            comment_len = max(len_byte(item['comment']), comment_len)
            appendComment_len = max(len_byte(item['appendComment']), appendComment_len)
            itemType_len = max(len_byte(item['type']), itemType_len);

            ws.write(line, 0, item['commentTime'])
            ws.write(line, 1, item['comment'], style)
            ws.write(line, 2, item['appendComment'], style)
            ws.write(line, 3, item['type'])

        ws.col(0).width = min(time_len, 80*256)
        ws.col(1).width = min(comment_len, 80*256)
        ws.col(2).width = min(appendComment_len, 80*256)
        ws.col(3).width = min(itemType_len, 80*256)
        # 保存excel文件
        wb.save(Target_File+'/data.xls')

if len(sys.argv) < 3:
    print("必须输入商品 ID 和店铺 ID")
    sys.exit()
itemId = sys.argv[1]
sellerId = sys.argv[2]
#itemId = "602623204515"
#sellerId = "2206529520669"
taobao=TaoBao(itemId, sellerId)

myJsonData = {}


# with open(Target_File+'/data.json', 'r') as f:
#     myJsonData = json.load(f)

Target_File = "output/{}_{}".format(itemId, sellerId)

if os.path.exists(Target_File):
    shutil.rmtree(Target_File)
os.makedirs(Target_File);

# myJsonData=taobao.getPageData(1)

myJsonData=taobao.getAllData()
with open(Target_File+'/data.json', 'w') as f:
    json.dump(myJsonData, f, ensure_ascii=False, indent=4, separators=(',', ':'))


simplifyData = taobao.simplifyData(myJsonData)

taobao.toExcel(simplifyData)

allComment = ""
for item in simplifyData:
    allComment = allComment+"\n"+item['comment']+"\n"+item['appendComment']+"\n"
taobao.analyze(allComment)

subprocess.call(["open", Target_File])