# -*- coding: utf-8 -*-
import os
import json
import sys
import xlwt
import hashlib
import time
import urllib.request

dirname, filename = os.path.split(os.path.abspath(__file__)) 

def download_img(img_url):
    header = {} # 设置http header
    request = urllib.request.Request(img_url, headers=header)
    # try:
    response = urllib.request.urlopen(request)
    img_name =hashlib.md5(img_url.encode("utf-8")).hexdigest()+".png"
    avatar_filename = os.path.join(dirname,"out","avatar",img_name)
    if (response.getcode() == 200):
        with open(avatar_filename, "wb") as f:
            f.write(response.read()) # 将内容写入图片
        return img_name
    # except:
    #     return "failed"

def load_json_data(file):
    content = open(file, encoding='utf8')
    json_data=json.load(content)
    return json_data
#设置表格样式
def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style

def data_to_excel(robots):
    robots_tags=["昵称","头像地址","机器人类型 （1:游客,2:手机用户,3:上传昵称,4:上传头像和昵）","wechat_head_img_url","wxid"] 
    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding = 'utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('Table')
    #写第一行
    for i in robots_tags:
        print ("序号：%s   值：%s"%(robots_tags.index(i) , i))
        worksheet.write(0,robots_tags.index(i),i,set_style('Times New Roman',220,True))

       
    index =1
    for (k,robot) in  robots.items(): 
        print( "dict[%s]="% k,robot)
        worksheet.write(index,0,robot["nick_name"])
        local_avatar_filename = download_img(robot["head_img"])
        worksheet.write(index,1,local_avatar_filename)
        worksheet.write(index,2,4)
        worksheet.write(index,3,robot["head_img"])
        worksheet.write(index,4,robot["wxid"])
        
        index=index+1
    sttaf_time=time.strftime("%Y_%m_%d_%H_%M_%S", time.localtime())
    xls = os.path.join(dirname,"out","robot_"+sttaf_time+".xls")
    workbook.save(xls)

def main():
    print('开始处理数据',dirname)
    json_dir_path = os.path.join(dirname,"json")
    print('json 文件夹', json_dir_path)
    robots = {}
   
    #--遍历json 文件夹
    for root,dirs,files in os.walk(json_dir_path):
        for file in files:
                print( file )
                wechat_users = load_json_data( os.path.join(root,file))
                print( type( wechat_users ) )
                for wechat_user in wechat_users:
                    wxid = wechat_user["wxid"]
                    robots[wxid] = wechat_user
   
    data_to_excel(robots)


if __name__ == '__main__':
    main()