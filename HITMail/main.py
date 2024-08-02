import smtplib
from email.message import EmailMessage
from email.headerregistry import Address, Group
import email.policy
import mimetypes
import base64
import pandas
import time
import traceback

class SendEmail(object):
    """
    python3 发送邮件类
    格式： html
    可发送多个附件
    """
    def __init__(self, smtp_server, smtp_user, smtp_passwd, sender, recipient):
        # 发送邮件服务器,常用smtp.163.com
        self.smtp_server = smtp_server
        # 发送邮件的账号
        self.smtp_user = smtp_user
        # 发送账号的客户端授权码
        self.smtp_passwd = smtp_passwd
        # 发件人
        self.sender = sender
        # 收件人
        self.recipient = recipient

        # Use utf-8 encoding for headers. SMTP servers must support the SMTPUTF8 extension
        # https://docs.python.org/3.6/library/email.policy.html
        self.msg = EmailMessage(email.policy.SMTPUTF8)

    def set_header_content(self, subject, content):
        """
        设置邮件头和内容
        :param subject: 邮件题头
        :param content: 邮件内容
        :return:
        """
        self.msg['From'] = self.sender
        self.msg['To'] = self.recipient
        self.msg['Subject'] = subject
        self.msg.set_content(content, subtype="html")

    def set_accessories(self, path_list: list):
        """
        添加附件
        :param path_list: [{"path": ""}, {"name": ""}]
        :return:
        """
        for path_dict in path_list:
            path = path_dict['path']
            name = path_dict['name']
            # print(path, name)
            ctype, encoding = mimetypes.guess_type(path)
            if ctype is None or encoding is not None:
                # No guess could be made, or the file is encoded (compressed), so
                # use a generic bag-of-bits type.
                ctype = 'application/octet-stream'

            maintype, subtype = ctype.split('/', 1)
            with open(path, 'rb') as fp:
                self.msg.add_attachment(fp.read(), maintype, subtype, filename=name)
                # self.msg.add_attachment(fp.read(), maintype, subtype, filename=self.dd_b64(name))

    def send_email(self):
        """
        发送邮件
        :return:
        """
        with smtplib.SMTP_SSL(self.smtp_server, port=465) as smtp:
            # HELO向服务器标志用户身份
            smtp.ehlo_or_helo_if_needed()
            # 登录邮箱服务器
            smtp.login(self.smtp_user, self.smtp_passwd)
            print("Email:{}==>{}".format(self.sender, self.recipient))
            smtp.send_message(self.msg)
            print("成功发送!")

    @staticmethod
    def dd_b64(param):
        """
        对邮件header及附件的文件名进行两次base64编码，防止outlook中乱码。
        email库源码中先对邮件进行一次base64解码然后组装邮件
        :param param: 需要防止乱码的参数
        :return:
        """
        param = '=?utf-8?b?' + base64.b64encode(param.encode('UTF-8')).decode() + '?='
        param = '=?utf-8?b?' + base64.b64encode(param.encode('UTF-8')).decode() + '?='
        return param


if __name__ == '__main__':

    # smtp_server = "smtp.hit.edu.cn"
    # smtp_user = "xxx@hit.edu.cn"  # 发送邮件的账号
    # smtp_passwd = "xxxxxxxx"          # 发送账号的客户端授权码
    # sender = Address("xxx", "fansiyue", "hit.edu.cn")
    smtp_server = "smtp.163.com"
    smtp_user = "xxx@163.com"  # 发送邮件的账号
    smtp_passwd = "xxxxxxxxx"          # 发送账号的客户端授权码
    sender = Address("xxx", "xxx", "163.com")

    # majors = ['电气', '控制', '动力']
    # 手动更改专业学科名称
    majors = ['动力']

    content_pass = """
        <html>
            <p>某某同学：</p>
            <p>  您好，恭喜您通过哈尔滨工业大学（深圳）夏令营预推免面试，经材料审核与考核，您的面试成绩请查收附件。此成绩仅为您的
            推免面试成绩，最终的推免综合成绩将综合考虑学业背景、专业素养、综合素质、创新能力、外语水平及思想品德等方面综合评定，
            具体接收情况请关注后续QQ群信息。</p>
            <p>根据“哈尔滨工业大学（深圳）接收2025年推免生（含直博生）工作办法”的相关规定，推免生应在全国推免服务系统开通后当日内，
            登陆全国推免服务系统，选择深圳校区进行申请报名，完成网上录取确认。未在规定时间内完成确认的推免生，
            原则上我校将不再保留其接收资格。
</p>
            <p style="text-align: right">哈尔滨工业大学（深圳）机电工程与自动化学院</p>
            <p style="text-align: right">2024年8月2日</p>
        </html>
    """
    content_fail = """
        <html>
            <p>某某同学：</p>
            <p>您好，您参加哈尔滨工业大学（深圳）预推免复试，很遗憾地通知您，
            经材料审核与考核，未通过本次推免面试，感谢您选择哈尔滨工业大学（深圳），祝您前程似锦，一切顺利。</p>
            <p style="text-align: right">哈尔滨工业大学（深圳）机电工程与自动化学院</p>
            <p style="text-align: right">2024年8月2日</p>
        </html>
    """
    subject = "哈尔滨工业大学（深圳）机电工程与自动化学院预推免面试结果通知"

    for major in majors:
        print(f"------------------------------------{major}")
        df = pandas.read_excel('动力夏令营成绩邮箱.xlsx')
        # df = pandas.read_excel('../第二批结果/testData.xlsx', sheet_name=major)

        for i in range(len(df)):

            stu_name = df.loc[i]['姓名']
            stu_email_pre, stu_email_domain = df.loc[i]['电子邮箱'].split('@')
            # print(stu_email_pre, stu_email_domain)
            if df.loc[i]['面试结果'] == '通过':
                pass_flag = True
                # print("通过")
            else:
                pass_flag = False
                # print("不通过")
            if pass_flag:
                recipient = Group(addresses=[Address(stu_name, stu_email_pre, stu_email_domain)])
                path_list = [
                    {"path": f"D:/控制面试/Docx-20230727/Docx/results/pdf/动力/哈尔滨工业大学（深圳）推免面试成绩单-{stu_name}.pdf",
                    "name":  f"哈尔滨工业大学（深圳）推免面试成绩单-{stu_name}.pdf"}]
                content = content_pass.replace('某某', stu_name)
            else:
                # path_list = []
                recipient = Group(addresses=[Address(stu_name, stu_email_pre, stu_email_domain)])
                path_list = [
                    {"path": f"D:/控制面试/Docx-20230727/Docx/results/pdf/动力/哈尔滨工业大学（深圳）推免面试成绩单-{stu_name}.pdf",
                    "name":  f"哈尔滨工业大学（深圳）推免面试成绩单-{stu_name}.pdf"}]
                content = content_fail.replace('某某', stu_name)

            # 发送邮件
            sd = SendEmail(smtp_server, smtp_user, smtp_passwd, sender, recipient)  # 创建对象
            sd.set_header_content(subject, content)  # 设置题头和内容
            # if pass_flag:
            sd.set_accessories(path_list)  # 添加附件
            # else:
            #     pass

            # print(sd.recipient)
            try:
                sd.send_email()  # 发送邮件
            except:
                print("Error")
                print(traceback.print_exc())
                exit()
            
            # time.sleep(240)
