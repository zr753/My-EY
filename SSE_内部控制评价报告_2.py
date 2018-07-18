
# coding: utf-8

# In[1]:


import urllib.parse  
import urllib.request
import json
import re
import time
import requests
import os
import xlsxwriter
import xlrd
from pdfminer.pdfparser import PDFParser  # 从一个文件中获取数据
from pdfminer.pdfparser import PDFDocument  # 保存获取的数据，和PDFParser是相互关联的
from pdfminer.pdfinterp import PDFPageInterpreter  # 处理页面内容
from pdfminer.pdfdevice import PDFDevice  # 将其翻译成你需要的格式
from pdfminer.pdfinterp import PDFResourceManager  # 用于存储共享资源，如字体或图片 
from pdfminer.layout import LAParams  # 设定参数进行分析
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal  # 识别解析完成的元素类型

proxy_info = { 'host' : 'CNWEB.ey.net','port' : 8080 }
proxy_support = urllib.request.ProxyHandler({"http" : "http://%(host)s:%(port)d" % proxy_info})
opener = urllib.request.build_opener(proxy_support)
urllib.request.install_opener(opener)

print("该程序是用于识别上海证券交易所上市公司内部控制评价报告\n注意！改程序为爬虫程序，频繁使用会造成IP封禁，为避免对EY造成影响，建议一日使用一次！\n如果对程序有任何疑问，请联系作者Matt Zhang， Email: bowenpossible@126.com")
os.system('pause')

data_total=[]
Year = "2017"
exe_path= os.getcwd()

def get_info(page,Year):
    global data_total
    url=r"http://query.sse.com.cn/search/getSearchResult.do"
    values={
            'search': 'qwjs',
            'jsonCallBack': 'jQuery111208667413195081248_1523539789774',
            'page':page,
            'searchword':'T_L CTITLE T_D E_KEYWORDS T_JT_E likeT_L' + Year + '年度内部控制评价报告T_RT_R',
            'orderby': '-CRELEASETIME',
            'perpage': '10',
            '_': '1523539789780',
            }#get方式所需参数
    data=urllib.parse.urlencode(values)#将参数转换为str
    theurl=url+"?"+data#合并参数到请求链接中去，在get方式中，数据请求是通过带参数的链接访问的
    headers = {
            'accept': "*/*",
            'accept-language': "zh-CN,zh;q=0.9",
            'connection': "keep-alive",
            'cookie': "yfx_c_g_u_id_10000042=_ck18031622331116025898951797532; VISITED_FUND_CODE=%5B%22500006%22%5D; VISITED_STOCK_CODE=%5B%22600287%22%2C%22600247%22%2C%22600612%22%2C%22600000%22%5D; VISITED_COMPANY_CODE=%5B%22500006%22%2C%22600287%22%2C%22600247%22%2C%22600612%22%2C%22600000%22%5D; VISITED_MENU=%5B%229045%22%2C%2210005%22%2C%228451%22%2C%228639%22%2C%228352%22%2C%228314%22%2C%228319%22%2C%2210006%22%2C%228312%22%2C%228349%22%2C%228307%22%5D; seecookie=%5B500006%5D%3A%u57FA%u91D1%u88D5%u9633%2C%5B600287%5D%3A%u6C5F%u82CF%u821C%u5929%2C%5B600247%5D%3AST%u6210%u57CE%2C%5B600612%5D%3A%u8001%u51E4%u7965%2C%5B600000%5D%3A%u6D66%u53D1%u94F6%u884C%2C%u81EA%u5F8B%2C%u81EA%u67E5%2C2017%u5E74%2C2017%u5E74%20%u81EA%u5F8B%2C2017%u5E74%20%u81EA%u6211%2C2017%u5E74%20%u5185%u90E8%2C%u5185%u90E8%20%u68C0%u67E5%2C%u5185%u90E8%u63A7%u5236%2C%u5185%u90E8%u63A7%u5236%u8BC4%u4EF7%u62A5%u544A; yfx_f_l_v_t_10000042=f_t_1521210791596__r_t_1523539683092__v_t_1523539789363__r_c_3",
            'dnt': "1",
            'host': "query.sse.com.cn",
            'referer': "http://www.sse.com.cn/home/search/?webswd=%E5%86%85%E9%83%A8%E6%8E%A7%E5%88%B6%E8%AF%84%E4%BB%B7%E6%8A%A5%E5%91%8A",
            'user-agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36",
            }#get方式所需header  'accept-encoding': "gzip, deflate",
    req = urllib.request.Request(theurl,None,headers)#合并url和header内容
    response = urllib.request.urlopen(req)#请求内容
    the_page = response.read()#读取请求内容
    #print(the_page)
    m = re.match(".*jQuery111208667413195081248_1523539789774\((.*)\).*",the_page.decode("utf8"))#将返回的json里的内容提取出来
    hjson = json.loads(m.group(1))#解析json内容
    #print(hjson)
    for content in hjson["data"]:
        data_total.append([content['CTITLE_TXT'],'www.sse.com.cn'+content['CURL']])
    if page==1:
        return hjson["countPage"]
    
def download(exe_path, i):
    global data_total
    proxies={"http": "http://CNWEB.ey.net:8080", "https": "http://CNWEB.ey.net:8080", }
    file_url = data_total[i][1]
    r = requests.get("http://" + file_url, proxies=proxies, stream=True)
    with open(os.path.join(exe_path,data_total[i][0].replace("*", "_")+".pdf"), "wb") as pdf:
        for chunk in r.iter_content(chunk_size=1024):
            if chunk:
                pdf.write(chunk)
    print("Downloaded " + str(i+1) + " / " + str(len(data_total)))
    time.sleep(1)

def execute_Download():
    global data_total
    global exe_path
    error_count = 0
    while error_count <= 5:
        if data_total == []:
            break
        error_download = []
        for i in range(0,len(data_total)):
            try:
                download(exe_path, i)
            except:
                print("ip封禁中...请等待30秒")
                time.sleep(30)
                error_download.append(data_total[i])
                error_count = error_count + 1
                if error_count > 5:
                    print("ip封禁超过5次...正在转储余下数据，并生成文件...")
                    for j in range (i+1, len(data_total)):
                        error_download.append(data_total[j])
                    workbook = xlsxwriter.Workbook(os.path.join(exe_path,"Error_Download_List.xlsx"))
                    worksheet = workbook.add_worksheet("Error_Download_List")
                    for i in range (0, len(error_download)):
                        worksheet.write(i, 0, error_download[i][0])
                        worksheet.write(i, 1, error_download[i][1])
                    workbook.close()
                    error_download = []
                    return "Error"
            else:
                continue
        data_total = error_download
        
    return "Success"

def scan_PDF():
    log=[]
    for files in os.walk(exe_path):  # 遍历文件夹下的文件  
        for i in range (0, len(files[2])):  #[0]是根目录，[1]是子文件夹，[2]是子文件
            if os.path.splitext(files[2][i])[1]=='.pdf' or os.path.splitext(files[2][i])[1]=='.PDF':  # 验证后缀名是否为pdf，os.path.splitext方法用来分割后缀名
                file_name = files[2][i]
                try:
    #*******************************************************获取数据********************************************************
                    file = open(os.path.join(exe_path,file_name), 'rb')
                    parser = PDFParser(file)  # Translate the binary file to recognizable datastream (PDFParser object)
                    document = PDFDocument()  # 创建一个PDF文档
                    parser.set_document(document)  # 连接分析文档
                    document.set_parser(parser)
                    document.initialize()  # pdf 初始化 
                    resource=PDFResourceManager()
                    laparams=LAParams()
                    device=PDFPageAggregator(resource,laparams=laparams)  # 创建一个PDF设备对象
                    interpreter=PDFPageInterpreter(resource,device)  # 创建一个PDF解释器对象
                    content=[]
                    i =0 
                    for page in document.get_pages():
                        interpreter.process_page(page)  # 解释器完成解释后，将内容传送至device里（device里包含资源管理器的PDF资源）  
                        layout=device.get_result()
                        for raw_data in layout:
                            if (isinstance(raw_data, LTTextBoxHorizontal)):  # 这里来判断取得的元素，我们这里要取得的是TextBox内容，其他内容或样例请直接print(x)查看
                                #print(i,raw_data.get_text())
                                content.append(raw_data.get_text())  # 将获得的内容赋值至content中
                                #i=i+1
    #**********************************************************************************************************************

    #*******************************************************识别数据********************************************************
                    tag = 0  # 验证是否存在重大缺陷等问题存在，如发现异常则加1
                    validation = -1 # 验证是否存在数据无法读取，默认为无法读取，能读取数据后赋值为0
                    for i in range(0, len(content)):
                        if content[i].split(" ")[0].find("公司代码") > -1:
                            print("公司代码： "+ content[i].split("：")[1].split(" ")[0])
                            sec_code = content[i].split("：")[1].split(" ")[0]  # 读取证券代码 
                            print("公司简称： "+ content[i].split("：")[2])  
                            sec_name = content[i].split("：")[2].strip()  # 读取证券名称
                            validation = 0
                        if (content[i].find("重大缺陷") > -1 or content[i].find("重要缺陷") > -1 or content[i].find("一致") > -1) and content[i+1].find("√") > -1 :
                            #print(content[i])
                            #print(content[i+1])
                            if content[i].find("一致") > -1:
                                question = content[i].split(" ")[1]
                                if content[i+1].find("√否") > -1:
                                    answer = "不一致"
                                    tag = tag + 1
                                elif content[i+1].find("√是") > -1:
                                    answer = "一致"
                                else:
                                    continue
                                log.append([sec_code, sec_name, question, answer])
                            elif (content[i].find("重大缺陷") > -1 and content[i].split(" ")[1] == "重大缺陷") or (content[i].find("重要缺陷") > -1 and content[i].split(" ")[1] == "重要缺陷"): 
                                question = content[i+1].split(" ")[0]
                                if content[i+1].find("√是") > -1:
                                    answer = "是"
                                    tag = tag +1
                                elif content[i+1].find("√否") > -1:
                                    answer = "否"
                                else:
                                    continue
                                log.append([sec_code, sec_name, question, answer])
                            elif (content[i].find("重大缺陷") > -1 and content[i].split(" ")[1] != "重大缺陷") or (content[i].find("重要缺陷") > -1 and content[i].split(" ")[1] != "重要缺陷"):
                                question = content[i].split(" ")[1]
                                if content[i+1].find("√是") > -1:
                                    answer = "是"
                                    tag = tag +1
                                elif content[i+1].find("√否") > -1:
                                    answer = "否"
                                else:
                                    continue
                                log.append([sec_code, sec_name, question, answer])
                            else:
                                continue
                    print(tag, validation)


                    if tag > 0:  # 如果tag大于0，表示有异常，则将异常结果写入文件
                        file = open(os.path.join(exe_path, 'Final_Result.txt'), 'a+')  # 写入访问记录
                        file.write(sec_code+" "+sec_name+"\r\n")
                        file.close()

                    if validation == -1:  # 如果等于-1则为数据无法读取，需要人工核查
                        file = open(os.path.join(exe_path, 'Error_Log.txt'), 'a+')  # 写入访问记录
                        file.write(os.path.join(exe_path, file_name)+"\r\n")
                        file.close()
    #***********************************************************************************************************************
                except: # 如果文件本身错误，则写入错误日志，待人工核查
                    file = open(os.path.join(exe_path, 'Error_Log.txt'), 'a+')  # 写入访问记录
                    file.write(os.path.join(exe_path, file_name)+"\r\n")
                    file.close()
                else:
                    continue

    workbook = xlsxwriter.Workbook(os.path.join(exe_path,"Log_File.xlsx"))
    worksheet = workbook.add_worksheet("Log_File")

    columns=['Sec_Code', 'Sec_Name', 'Question', 'Answer']
    for i in range (0, 4):
        worksheet.write(0, i, columns[i])

    for i in range (0, len(log)):
        for j in range (0, 4):
            worksheet.write(i+1, j, log[i][j])

    workbook.close()

    print("识别完成")
    os.system('pause')

def execute_get_Info(Year):
    try:
        total_page=get_info(1,Year)
        time.sleep(0.5)
        print("Page "+ "1" +" / " + total_page)
    except:
        print("该IP已经被Ban，请关闭程序隔6小时后再试..")
        return "Error"
    else:
        if total_page == "1":
            print("完成信息爬取，准备下载PDF文件...")
            time.sleep(15)
            return "Success"
        else:
            for i in range(2,int(total_page)+1):
                try:
                    get_info(i,Year)
                    print("Page "+ str(i) +" / " + total_page)
                    time.sleep(0.5)
                except:
                    print("该IP已经被Ban，请关闭程序隔6小时后再试..")
                    return "Error"
                else:
                    continue    
    print("完成信息爬取，准备下载PDF文件...")
    time.sleep(15)
    return "Success"

if os.path.exists(os.path.join(exe_path,"Error_Download_List.xlsx")) == True:
    print("检测到错误文件，继续根据上次断点进行下载...")
    excel_path = os.path.join(exe_path,"Error_Download_List.xlsx")
    excel_data = xlrd.open_workbook(excel_path)
    table = excel_data.sheet_by_name('Error_Download_List')
    nrows = table.nrows
    for i in range(0, nrows):
        data_total.append([table.cell(i,0).value,table.cell(i,1).value])
    os.remove(os.path.join(exe_path,"Error_Download_List.xlsx"))
    msg=execute_Download()
    if msg == "Error":
        print("程序退出...请至少间隔6小时再尝试")
    else:
        scan_PDF()
else:
    Year = input("请输入您想要查找的年份：（格式为四位数，例：2017。对错误输入导致网站封禁IP等后果不承担任何责任！）")
    msg = execute_get_Info(Year)
    if msg == "Error":
        os.system('pause')
    else:
        msg = execute_Download()
        if msg == "Error":
            print("程序退出...请至少间隔6小时再尝试")
            os.system('pause')
        else:
            scan_PDF()

