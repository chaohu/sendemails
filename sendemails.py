import json
import smtplib
import sys
import win32com.client
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from pathlib import Path

def sendMail(session, sender_address, sender_pass, recv_address, subject, content, attach_pdf, attach_map):
    # param session 邮件链接
    # param sender_address 发送邮件地址
    # param sender_pass 发送邮件认证口令
    # param recv_address 接收邮箱
    # param subject 邮件主题
    # param content 邮件内容
    # param attach_pdf 邮件附件PDF文件地址
    message = MIMEMultipart() #message结构体初始化
    message['From'] = sender_address #你自己的邮箱
    message['To'] = recv_address #要发送邮件的邮箱
    message['Subject'] = subject
    # content,发送内容,这个内容可以自定义,'plain'表示文本格式
    mail_content = MIMEText(content, 'plain')
    message.attach(mail_content)
    # 添加附件1
    with open(attach_pdf, 'rb') as attach_pdf_file:
        mail_attach_pdf = MIMEApplication(attach_pdf_file.read(), _subtype = 'pdf')
        mail_attach_pdf.add_header('Content-Disposition', 'attachment', filename = str(attach_pdf.split('\\')[-1]))
        message.attach(mail_attach_pdf)
    #添加附件2
    with open(attach_map, 'rb') as attach_map_file:
        mail_attach_map = MIMEApplication(attach_map_file.read(), _subtype = 'pdf')
        mail_attach_map.add_header('Content-Disposition', 'attachment', filename = str(attach_map.split('\\')[-1]))
        message.attach(mail_attach_map)
    # message结构体内容传递给text,变量名可以自定义
    text = message.as_string()
    # 主要功能,发送邮件
    session.sendmail(sender_address,recv_address,text)

def main():
    input('请按下任意键执行批量处理程序。')
    try:
        current_path = Path(Path(sys.executable).resolve(strict=True).parents[0], 'supportfiles')
        # 获取主要参数
        parameters_path = Path(current_path, 'parameters.json')
        if not Path(parameters_path).exists() or not Path(parameters_path).is_file():
            raise ValueError('参数文件（parameters.json）不存在')
        with open(parameters_path, 'r') as parameters_file:
            parameters = json.load(parameters_file)
        # 获取邮件正文模板内容
        content_path = Path(current_path, parameters['content_txt'] + '.txt')
        if not Path(content_path).exists() or not Path(content_path).is_file():
            raise ValueError('邮件正文模板文件（content.txt）不存在')
        with open(content_path, 'r', encoding='utf-8') as content_txt:
            content = content_txt.read()
        template_docx_path = Path(current_path, parameters['template_docx'] + '.doc')
        if not Path(template_docx_path).exists() or not Path(template_docx_path).is_file():
            raise ValueError('邮件附件模板文件（template.docx）不存在')
        target_xlsx_path = Path(current_path, parameters['target_xlsx'] + '.xlsx')
        if not Path(target_xlsx_path).exists() or not Path(target_xlsx_path).is_file():
            raise ValueError('邮件接受者文件（target.xlsx）不存在')
        attach_map_path = Path(current_path, parameters['attach_map'] + '.pdf')
        if not Path(attach_map_path).exists() or not Path(attach_map_path).is_file():
            raise ValueError('邮件附件地图文件（map.pdf）不存在')
        word = win32com.client.Dispatch('Word.Application')
        try:
            wtemplate = word.Documents.Open(str(template_docx_path), ReadOnly = True)
            try:
                excel = win32com.client.Dispatch('KET.Application')
                try:
                    etarget = excel.Workbooks.Open(str(target_xlsx_path))
                    try:
                        etarget_worksheet = etarget.WorkSheets.Item('Target')
                        name = ''
                        email = ''
                        # 这里是smtp网站的连接,可以通过谷歌邮箱查看,步骤请看下边
                        session = smtplib.SMTP('smtp.gmail.com',587)
                        # 连接tls
                        session.starttls()
                        try:
                            # 登陆邮箱
                            session.login(parameters['account'], parameters['password'])
                            for row in range(2, etarget_worksheet.Rows.Count):
                                name = etarget_worksheet.Cells.Item(row, 1).Value2
                                if not name: break
                                else:
                                    # 获取接受者性别
                                    gender = etarget_worksheet.Cells.Item(row, 2).Value2
                                    # 获取接受者邮箱地址
                                    email = etarget_worksheet.Cells.Item(row, 3).Value2
                                    print(email)
                                    # 处理邮件正文
                                    if not gender:
                                        content_final = content.replace('%性别%', '').replace('%姓名%', name)
                                    else:
                                        content_final = content.replace('%性别%', gender).replace('%姓名%', name)
                                    # 处理邮件附件
                                    #for paragraph in wtemplate.Paragraphs:
                                    #    format = paragraph.Format
                                    #    if '%姓名%' in paragraph.Range.Text:
                                    #        paragraph.Range.Text = paragraph.Range.Text.replace('%姓名%', name)
                                    #        paragraph.Format = format
                                    wtemplate.FormFields.Item('name').Result = name
                                    pdf_path = Path(current_path, 'temp', name)
                                    if not Path(pdf_path).exists() or not Path(pdf_path).is_dir():
                                        Path(pdf_path).mkdir(parents=True, exist_ok=True)
                                    pdf_full_path = Path(pdf_path, parameters['attach_pdf'] + '.pdf')
                                    wtemplate.SaveAs2(str(pdf_full_path), 17)
                                    # 发送邮件
                                    sendMail(session, parameters['account'], parameters['password'], email, parameters['Subject'], content_final, str(pdf_full_path), str(attach_map_path))
                                    # 发送成功标记
                                    print("send {} successfully".format(email))
                                    etarget_worksheet.Cells.Item(row, 4).Value2 = 'success'
                        finally:
                            # 关闭连接
                            session.quit()
                    finally:
                        etarget.Save()
                        etarget.Close(0)
                finally:
                    excel.Quit()
            finally:
                wtemplate.Close(0)
        finally:
            word.Quit()
    except Exception as exception:
        print('执行过程中出现异常:')
        print(exception)
    input('批量处理程序执行完毕，请检查运行结果，按任意键退出程序。')
    
main()