import os
import xml.dom.minidom

import xlrd
import xlsxwriter

def Lists_Write(file,lists):
    workbook=xlsxwriter.Workbook(file)
    sheet=workbook.add_worksheet(u"sheet")
    for i in range(len(lists)):
        for j in range(len(lists[i])):
            sheet.write(j,i,lists[i][j])
    workbook.close()
    print('文件写入成功！保存至'+file)
def start(file):
    with open(file,encoding="UTF-8") as f:
        f=f.read().replace(':','')
        dom=xml.dom.minidom.parseString(f)
        root=dom.documentElement
        tag=root.getElementsByTagName('cser')
        tittle=[]
        cvdata=[]
        static=0
        if len(tag[0].getElementsByTagName('cstrCache'))!=len(tag[0].getElementsByTagName('cnumCache')):
            for i in range(len(tag)):#共有几组数据
                #获得所有标题
                data=tag[i].getElementsByTagName('ctx')
                for ij in range(len(data)):
                    cvtitle=data[ij].getElementsByTagName('cv')
                    tittle.append(cvtitle[0].firstChild.data) #0指仅有一个
                cc=tag[i].getElementsByTagName('cnumCache')
                #获得所有数据
                for j in range(len(cc)):
                    if static%2==0:
                        cvdata.append('SPLIT')
                    else:
                        cvdata.append('SPLIT')
                    ccd=cc[j].getElementsByTagName('cv')
                    for k in range(len(ccd)):
                        cvdata.append(ccd[k].firstChild.data)
                    static+=1
        else:
            for i in range(len(tag)):#共有几组数据cser
            #获得所有标题
                data=tag[i].getElementsByTagName('ctx')
                for ij in range(len(data)):
                    cvtitle=data[ij].getElementsByTagName('cv')
                    tittle.append(cvtitle[0].firstChild.data) #0指仅有一个
                ccx=tag[i].getElementsByTagName('cstrCache')
                ccy=tag[i].getElementsByTagName('cnumCache')
                #获得所有数据
                for j in range(len(ccx)):
                    cvdata.append('SPLIT')
                    ccd=ccx[j].getElementsByTagName('cpt')
                    for k in range(len(ccd)):
                        cvdata.append(ccd[k].firstChild.firstChild.data)
                for m in range(len(ccy)):
                    cvdata.append('SPLIT')
                    ccd=ccy[m].getElementsByTagName('cpt')
                    for n in range(len(ccd)):
                        cvdata.append(ccd[n].firstChild.firstChild.data)
        return [cvdata,tittle]
def embed(fpath):
    if (os.path.exists(fpath)):
        list = os.listdir(fpath)
        for i in range(0,len(list)):
            path = os.path.join(fpath,list[i])
            if os.path.isfile(path):
                if (list[i].split('.'))[1]=='xml':
                        lists=start(path)
                        listss,tittles=lists[0],lists[1]
                        strs=''.join(str(i)+',' for i in listss)
                        listsss=strs.split('SPLIT')
                        newlist=[]
                        for cl in listsss:
                            if cl=='':
                                listsss.remove(cl)
                        for ne in listsss:
                            new=ne.strip(',').split(',')
                            newlist.append(new)
                        Lists_Write(fpath+'/'+(list[i].split('.'))[0]+'-'+''.join(str(z)+',' for z in tittles).replace('/','').replace(':','')+'.xlsx',newlist)
    else:
        print('文件夹不存在！')
#embed('D:/1/word/charts')#输入XML所在文件夹路径
#def fitwordchart(wordfile):