# -*- coding: utf-8 -*-
'''
author:Zhang Lianzhong
date:20191003
description:批量基于模板发送邮件。发送人信息读取自同目录下的excel文件
'''
from email.header import Header #处理邮件主题
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText # 处理邮件内容
from email.utils import parseaddr, formataddr #用于构造特定格式的收发邮件地址
import smtplib #用于发送邮件
from emails.ReadData import readRecieveInfo

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

#设置发送人邮箱信息
from_addr = 'qinghaihudream@163.com'
password = 'Zhang640324'

# 邮箱服务器地址
smtp_server = 'smtp.163.com'

#server = smtplib.SMTP_SSL(host=smtp_server,port=465)
server = smtplib.SMTP(host=smtp_server, port=25)
server.login(from_addr, password)

#读取接收人信息
#to_addrs = ['QQ号@qq.com','***@163.com']
recieverInfo = readRecieveInfo('invitee.xlsx')
print(recieverInfo)


for username,to_addr in recieverInfo:

    para1 = "<p>%s,你好！从校友总会老师处得知，你已经选择来到深圳工作发展，在此，表示热烈的欢迎。</p>" % (username)
    para2 = "<p><b>电子科技大学深圳校友会</b>隶属于电子科技大学校友总会，主要负责在深成电校友相关活动。深圳校友会自成立以来，除每年承办迎新、校友企业参观学习等活动以外，还组建了如跑团、羽毛球协会、足球协会、金融协会等比较专业性的群体。在大家遇到困难的时候，校友会也会尽力给大家提供帮助。一句话总结——<font color='#FF0000'>成电帮，帮成电。</font></p>"
    para3 = "<p>我是2012级外国语学院的关敏，也是电子科技大学深圳校友会的一名志愿者。我们已经为2018届来深的成电校友建好了<font color='#FF0000'>微信群——2018届成电深圳校友会</font>，欢迎你的加入！由于微信群目前已经超过100人，无法扫码加入，大家可通过加我的微信（定格：18682253039）或者让同届来深好友邀请进群即可，谢谢！</p>"
    para4 = "<p>同时，也欢迎大家关注深圳校友会公众号：成电帮（微信号：cdbang66）/电子科技大学深圳校友会（微信号：uestcsz），以了解最新的校友会相关讯息。重点来了！！！深圳校友会预计在9月15日/16日举办新一届来深校友的迎新会，以欢迎大家加入深圳校友会这个大家庭，让选择深漂的你，依然能感受到来自母校的关怀和温暖。如果你有意来参加，请回复本邮件，在邮件主题后+是否有意参加+姓名（eg：欢迎加入成电深圳校友会+有意参加+关敏），谢谢！</p>"
    para5 = "<p>---------------------------------------------------</p>"
    para6 = "<p>祝好！</p>"
    para7 = "<br/>"
    para8 = "关敏<br/>"
    para9 = "电话：18682253039<br/>"
    para10 = "邮箱：angeleguan@163.com"
    content = para1+para2+para3+para4+para5+para6+para7+para8+para9+para10
    # 设置邮件信息
    msg = MIMEMultipart('mixed')
    body = MIMEText(content, 'html', 'utf-8')
    # 如果不加下边这行代码的话，上边的文本是不会正常显示的，会把超文本的内容当做文本显示
    #body["Content-Disposition"] = 'attachment; filename="uestc.html"'
    msg.attach(body)

    msg['From'] = _format_addr('青海湖<%s>'%from_addr)
    msg['To'] = _format_addr('%s<%s>'%(username,to_addr))
    msg['subject'] = Header('欢迎加入电子科技大学深圳校友会（批量发送测试邮件，不用回复！）', 'utf-8')

    try:
        server.sendmail(from_addr, to_addr, msg.as_string())
    except:
        #print('发送失败，再次尝试')
        #server.sendmail(from_addr, to_addr, msg.as_string())
        print('发送邮件失败，姓名：%s,邮箱地址：%s'%(username, to_addr))
    print('成功发送邮件到:'+to_addr)

server.quit()