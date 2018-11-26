# coding:utf-8

import thirdpart.requests as requests
from thirdpart.docx import Document
from thirdpart.docx.shared import Pt
from thirdpart.docx.shared import Inches
from thirdpart.docx.shared import Cm
from thirdpart.docx.oxml.ns import qn

import re
import json
import os

class PyDocx:

    def __init__(self):
        self.document=Document()
        self.pagesizelists = {
            "A4": {
                "width": Cm(21),
                "height": Cm(29.7)
            },
            "A3": {
                "width": Cm(29.7),
                "height": Cm(42)
            },
            "letter": {
                "width": Cm(21.59),
                "height": Cm(27.94)
            }
        }

    def pagewidth(self,section,pagewidth):
        #页面宽度
        self.document.sections[section].page_width=pagewidth

    def pageheight(self,section,pageheight):
        #页面高度
        self.document.sections[section].page_height=pageheight

    def pagesize(self,section,pagesizelist):
        self.document.sections[section].page_width=self.pagesizelists[pagesizelist]["width"]
        self.document.sections[section].page_height=self.pagesizelists[pagesizelist]["height"]

    def verticalpage(self,section):
        #竖向页面
        if self.document.sections[section].page_height<self.document.sections[section].page_width:
            tempheight=self.document.sections[section].page_height
            tempwidth=self.document.sections[section].page_width
            self.document.sections[section].page_height=tempwidth
            self.document.sections[section].page_width=tempheight

    def horizontalpage(self,section):
        #横向页面
        if self.document.sections[section].page_height>self.document.sections[section].page_width:
            tempheight=self.document.sections[section].page_height
            tempwidth=self.document.sections[section].page_width
            self.document.sections[section].page_height=tempwidth
            self.document.sections[section].page_width=tempheight


    def addcontent(self):
        self.document.add_paragraph("test")


    def savedoc(self,filepath=os.path.join(os.getcwd(),"doc"),filename="texts.docx"):
        self.document.save(os.path.join(filepath,filename))
        print("Save success!")


class BaiduDownload():
    def __init__(self,baiduurl):
        self.baiduurl=baiduurl
        self.savepath=os.path.join(os.getcwd(),"downloadjson")
        self.requesturl=""
        self.downloadjsonlist=[]
        self.title=self.getdoctitle()

    """
    获取doc title
    """
    def getdoctitle(self):
        html=requests.get(self.baiduurl)
        #print(html.encoding)#编码1
        #print(html.apparent_encoding)#编码2
        html=html.text
        title=re.split("<title>",html)[1]
        title=re.split("_",title)[0]
        title=title.encode("ISO-8859-1").decode("gb2312")
        return title


    #Step 1.获取百度下载请求地址
    def getrequesturl(self):
        temprequesturl=re.split("view/",self.baiduurl)[1]
        temprequesturl=re.split(".html",temprequesturl)[0]
        self.requesturl=f"https://wenku.baidu.com/browse/getrequest?doc_id={temprequesturl}&pn=1&rn=50&type=html"
        #https://wenku.baidu.com/browse/getrequest?doc_id=d81484a0db38376baf1ffc4ffe4733687e21fcb0&pn=1&rn=50&type=html
        return self.requesturl

    #Step 2.由Json数据获取下载pageIndex，pageloadUrl
    #返回https://wkbjbos.bdimg.com/#######的下载地址
    def getjsonurl(self):
        self.downloadjsonlist=[]
        webjsondata=requests.get(self.requesturl)
        webjsondata=webjsondata.text
        jsondata=json.loads(webjsondata)
        pageloadurls=jsondata["json"]
        for pages in pageloadurls:
            self.downloadjsonlist.append(pages["pageLoadUrl"])
        return self.downloadjsonlist

    #Step3.从downjsonlist下载JSON格式doc数据
    def downloaddocdata(self):
        i=0
        for js in self.downloadjsonlist:
            with open(os.path.join(self.savepath, f"{i}.json"), "w") as f:
                text = requests.get(js).text
                text = text.encode("utf-8", "ingore").decode("unicode_escape")[8:-1]
                #jsontext = json.loads(text)
                # print(json.dumps(jstext,indent=4, sort_keys=False))
                f.write(text)
                print(f"下载{i}.json文件成功")
                print("开始转换文件")
                i=i+1

    #Step4.JSON转换成doc格式
    def jsonconvert2doc(self):
        mydoc=PyDocx()
        i=0
        print(f"共{len(self.downloadjsonlist)}页")
        for js in self.downloadjsonlist:
            with open(os.path.join(self.savepath, f"{i}.json"), "r") as f:
                text=f.read()
                #print(text)
                docjsondata=json.loads(text)
                #print(docjsondata)
                #设置页面大小

                mydoc.document.styles['Normal'].font.name=u"宋体"
                mydoc.pagewidth(0,Pt(docjsondata["page"]["pw"]))
                mydoc.pageheight(0,Pt(docjsondata["page"]["ph"]))
                #获取内容增加内容
                npword=""
                for c in docjsondata["body"]:
                    if (not isinstance(c["c"],dict)) and (not c["c"]==" "):
                        npword=npword+c["c"]
                    else:
                        np = mydoc.document.add_paragraph()
                        run=np.add_run(npword)
                        run.font.name = 'Times New Roman'#西文字体
                        run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')#中文字体
                        run.font.size=Pt(c["p"]["h"])
                        npword=""

                i=i+1
            mydoc.savedoc(filename=self.title+".docx")


def main():
    #original=r"https://wenku.baidu.com/view/d81484a0db38376baf1ffc4ffe4733687e21fcb0.html?sxts=1541821552899"
    original=r"https://wenku.baidu.com/view/7a4e9036b8f67c1cfbd6b881.html"
    #requrl=oriurl2requrl("https://wenku.baidu.com/view/d81484a0db38376baf1ffc4ffe4733687e21fcb0.html?sxts=1541821552899")
    mybaidudownlaod=BaiduDownload(original)
    mybaidudownlaod.getrequesturl()
    mybaidudownlaod.getjsonurl()
    mybaidudownlaod.downloaddocdata()
    mybaidudownlaod.jsonconvert2doc()
if __name__=="__main__":
    main()
