# -- coding: utf-8 --


import tornado.httpserver
import tornado.ioloop
import tornado.options
import tornado.web
import tornado.log
import tornado.httpclient
import tornado.gen

import logging

import hmac
import json
import time
import base64
import hashlib
from hashlib import sha1, md5

import random
# import math
import re

import os
from urllib import parse
import requests
import shutil
import PyPDF2

import xlwings
from datetime import datetime
from num2words import num2words
from decimal import Decimal, ROUND_DOWN

# from aliyunsdkcore.client import AcsClient
# from aliyunsdkcore.request import CommonRequest

from tornado.options import define, options


class Application(tornado.web.Application):
    def __init__(self):
        self.domain = "localhost"
        self.scapegoat = "https://www.google.com/"
        handlers = [
            (r'^/$', Index),
            (r'^/index$', Index),
            (r'^/input$', Input),
            (r'^/output$', Output),
            (r'^/loading$', Loading),
            (r'^/downloads$', Downloads),
            (r'.*', ErrorHandler),

        ]

        settings = {
            "static_path": os.path.join(os.path.dirname(__file__), "../static"),
            "template_path": os.path.join(os.path.dirname(__file__), "../templates"),
            "cookie_secret": "Cs7ceta5auXQyPgKtcLkFm2zPSFmMPLz1OgATuHgw=",
            "login_url": "/login",
            "xsrf_cookies": True
        }
        tornado.web.Application.__init__(self, handlers, **settings, debug=False)


class BaseHandler(tornado.web.RequestHandler):

    def get_current_user(self):
        # 以ip地址为基础注册用户
        ip = self.request.remote_ip
        headers = self.request.headers
        user_agent = headers['User-Agent']
        s = hashlib.sha1()
        to_sha1_string = (ip + user_agent).encode('utf-8')
        s.update(to_sha1_string)
        user_document = {
            "ip": ip,
            "headers": headers,
            "user_agent": user_agent,
            "user_id": s.hexdigest()
        }
        if user_document:
            return user_document
        else:
            return None

    def stamptodate(self, timestamp):
        time_local = time.localtime(timestamp)
        # 先变成时间数组
        local_now = time.strftime("%Y-%m-%d-%H:%M:%S", time_local)
        # 转换成新的时间格式(2016-05-05-20:28:54)
        return local_now

    def mislead(self):
        return "https://www.google.com/"

    def get_lpo_information(self, user_id):
        xlwings_app = xlwings.App(visible=False, add_book=False)
        lpo_file_path = "..//file//" + user_id + "//input//lpo//"
        file_list = os.listdir(lpo_file_path)
        lpo_file_name = file_list[0]
        workbook_1 = xlwings_app.books.open(lpo_file_path + lpo_file_name)
        sheet1_demo = workbook_1.sheets[0]
        lpo_number = sheet1_demo.range('H' + str(5)).value
        send_date = datetime.today().strftime("%d/%b/%Y")
        code_a21 = sheet1_demo.range('A21').value
        trn_code = code_a21[20:].split(".")[-1].replace(" ", "")
        project = sheet1_demo.range('B9').value
        supplier = sheet1_demo.range('B6').value
        res_scan = {
            "type": "Tax Invoice",
            "date": send_date,
            "invoice": "",
            "project": project,
            "lpo": lpo_number,
            "trn": trn_code,
            "supplier": supplier,
            "lpo_file_name": lpo_file_name
        }
        workbook_1.close()
        xlwings_app.quit()
        invoice_json_path = "..//file//" + user_id + "//input//data//"
        file_list = os.listdir(invoice_json_path)
        invoice_json = file_list[0]
        with open(invoice_json_path + invoice_json, 'r') as up:
            invoice_code = up.read()
        res_scan["invoice"] = invoice_code
        return res_scan

    def make_payment_log(self, output, supplier, only_year, name_date, project, lpo_total, send_date,
                         lpo_file_name_date):
        # 修改Payment log，不改的话就注释掉
        xlwings_app = xlwings.App(visible=False)
        new_line_payment_log = xlwings_app.books.add()
        paylog_sheet1 = new_line_payment_log.sheets[0]
        last_row = 1
        new_row = last_row + 1
        paylog_sheet1.range('A' + str(new_row)).value = [
            "Finance Form " + only_year + "hq-" + name_date + "-" + project,
            last_row]
        paylog_sheet1.range('D' + str(new_row)).value = lpo_total
        paylog_sheet1.range('E' + str(new_row)).value = [project, send_date, send_date, send_date, send_date,
                                                         "processing",
                                                         supplier, " ", " ", " "]
        paylog_sheet1.range('A' + str(new_row)).color = (124, 252, 0)
        paylog_sheet1.range('B' + str(new_row)).color = (124, 252, 0)
        paylog_sheet1.range('D' + str(new_row)).color = (124, 252, 0)
        cell_A = 'A' + str(new_row)
        cell_M = "M" + str(new_row)
        # 设置内边线
        paylog_sheet1.range(cell_A, cell_M).api.Borders(9).lineStyle = 1
        paylog_sheet1.range(cell_A, cell_M).api.Borders(10).lineStyle = 1
        paylog_sheet1.range(cell_A, cell_M).api.Borders(11).lineStyle = 1
        new_line_name = "pre_payment_log.xlsx"
        paylog_sheet1.used_range.autofit()
        new_line_payment_log.save(output + "//pre_payment_log//" + new_line_name)
        new_line_payment_log.close()
        xlwings_app.quit()

    def get_LPO(self, res_scan, user_id):
        xlwings_app_1 = xlwings.App(visible=False, add_book=False)
        xlwings_app_2_raw = xlwings.App(visible=False, add_book=False)
        xlwings_app_2 = xlwings.App(visible=False, add_book=False)
        only_year = datetime.today().strftime("%Y")
        supplier = res_scan["supplier"].capitalize()
        department = "HQ"
        project = res_scan["project"]
        invoice = res_scan["invoice"]
        lpo = res_scan["lpo"]
        due_date = res_scan["date"]
        trn = res_scan["trn"]
        per_time = "90 Days"
        lpo_file_name = res_scan["lpo_file_name"]
        # lpo_file_path = "Z:\\Shared\\HQ\\13 IT\\2-IT Procurement\\4.LPO\\" + only_year + "\\"
        # financial_form_path = "Z:\\Shared\\HQ\\13 IT\\2-IT Procurement\\5.Payment\\1.Payment Certificate\\HQ\\" + only_year + "\\"
        output_path = "..//file//" + user_id + "//output//"
        lpo_file_path = "..//file//" + user_id + "//input//lpo//"
        lpo_pdf_path = "..//file//" + user_id + "//input//lpo_pdf//"
        invoice_path = "..//file//" + user_id + "//input//tax_invoice//"
        financial_form_path = "..//file//" + user_id + "//output//total_xlsx//"
        # 各类格式的日期生成
        name_date = datetime.today().strftime("%d-%b")
        send_date = datetime.today().strftime("%d-%b-%Y")
        work_form_date = datetime.today().strftime("%b/%Y")
        lpo_file_name_date = datetime.today().strftime("%d%b")
        final_pc_name = ("Finance Form " + only_year + department.lower() + "-" + lpo_file_name_date +
                         "-" + supplier.replace(".", "") + ".xlsx")
        # workbook_1, sheet1_demo指的是LPO文件
        workbook_1 = xlwings_app_1.books.open(lpo_file_path + lpo_file_name)
        sheet1_demo = workbook_1.sheets[0]
        workbook_2_raw = xlwings_app_2_raw.books.open("..//file//PaymentCertificateDemo.xlsx")
        # file for handover
        handover_file = financial_form_path + final_pc_name
        workbook_2_raw.save(handover_file)
        workbook_2_raw.close()
        xlwings_app_2_raw.quit()
        workbook_2 = xlwings_app_2.books.open(handover_file)
        sheet1 = workbook_2.sheets["Attachment 1"]
        sheet2 = workbook_2.sheets["2Conf. of M. Received)"]
        # supplier = sheet1_demo.range('B16').value
        # for lpo_row in range(1, 30):
        #     discription = sheet1_demo.range('B' + str(lpo_row)).value
        #     print("row:", lpo_row, "discription:", discription)
        figure_total = Decimal(0)
        for i in range(16, 31):
            lpo_row = i - 4
            description = sheet1_demo.range('B' + str(lpo_row)).value
            quantity = sheet1_demo.range('F' + str(lpo_row)).value
            unit = sheet1_demo.range('G' + str(lpo_row)).value
            rate = sheet1_demo.range('H' + str(lpo_row)).value
            total = sheet1_demo.range('I' + str(lpo_row)).value
            if description is None or total is None:
                break
            sheet1.range('B' + str(i - 2)).value = str(i - 15)
            sheet1.range('C' + str(i - 2)).value = description
            sheet1.range('E' + str(i - 2)).value = quantity
            sheet1.range('F' + str(i - 2)).value = unit
            sheet1.range('G' + str(i - 2)).value = rate
            sheet1.range('H' + str(i - 2)).value = total
            sheet1.range('I' + str(i - 2)).value = total
            sheet1.range('J' + str(i - 2)).value = "100%"
            sheet1.range('K' + str(i - 2)).value = total
            # 下面是2Conf. of M. Received)
            sheet2.range('A' + str(i - 9)).value = str(i - 15)
            sheet2.range('C' + str(i - 9)).value = str(invoice)
            sheet2.range('E' + str(i - 9)).value = description
            sheet2.range('F' + str(i - 9)).value = description
            sheet2.range('G' + str(i - 9)).value = unit
            sheet2.range('H' + str(i - 9)).value = quantity
            sheet2.range('J' + str(i - 9)).value = total
            sheet2.range('I' + str(i - 9)).value = rate
            # 根据单价判断是什么类型的商品
            low_disc = description.lower()
            if "program" in low_disc or "license" in low_disc:
                goods_type = "software"
            else:
                if int(rate) < 2000:
                    goods_type = "office supplies"
                else:
                    goods_type = "office equipment"
            # an additional counter of total
            figure_total += Decimal(total)
            # title of sheet2
            sheet2.range('B' + str(i - 9)).value = goods_type
        sheet2.range('A' + str(
            4)).value = ("Headoffice/Division:                     " + department +
                         "                              Department or Project:       " + project +
                         "           Place of Receipt:       " + project + "           Date:    " + send_date)
        figure_total = figure_total * Decimal("1.05").quantize(Decimal("0.01"), rounding=ROUND_DOWN)
        words_total = num2words(Decimal(figure_total), to='currency')
        sheet2.range('C23').value = words_total.replace("euro", "dirham").replace("cent", "fils")
        # 根据表单上的公式得出总价格
        no_tax_total = sheet1.range('K' + str(30)).value
        lpo_total = no_tax_total * 1.05
        input_tax = no_tax_total * 0.05
        # 新建一个日志文件, 可用于复制粘贴
        self.make_payment_log(output_path, supplier, only_year, name_date, project, lpo_total, send_date,
                              lpo_file_name_date)
        # 此时再回到付款确认单第一页填写内容
        sheet0 = workbook_2.sheets["Payment certificate"]
        sheet0.range('F' + str(5)).value = "PO Ref.No:     " + str(lpo)
        sheet0.range('F' + str(9)).value = "Payment Due Date:     " + str(due_date)
        sheet0.range('F' + str(13)).value = ("Payment Certificate No:    " + str(lpo))
        sheet0.range('A' + str(9)).value = "Name of Subcontractor/Supplier:     " + str(supplier)
        sheet0.range('A' + str(10)).value = "Tax Registration Number:     " + str(trn)
        sheet0.range('F' + str(15)).value = "Period  of  work  from:    " + str(work_form_date)
        sheet0.range('A' + str(7)).value = "Project Name:     " + project + " / IT"
        sheet0.range('A' + str(8)).value = str(project)
        sheet1.range('B' + str(7)).value = "Department:     " + str(project)
        sheet1.range('E' + str(7)).value = "     Date   :     " + str(send_date)
        sheet1.range('J' + str(7)).value = "For Payment Certificate No :     " + str(lpo)
        # 下面是纯税务Attachment4
        sheet4 = workbook_2.sheets["Attachment4"]
        sheet4.range('A' + str(5)).value = "Department:     " + project + " / IT"
        sheet4.range('C' + str(5)).value = None
        sheet4.range('D' + str(5)).value = "Date        :     " + str(send_date)
        sheet4.range('E' + str(5)).value = "For Payment Certificate No :     " + str(lpo)
        sheet4.range('B' + str(9)).value = str(invoice)
        sheet4.range('D' + str(9)).value = str(input_tax)
        sheet4.range('E' + str(9)).value = str(input_tax)
        # 关闭所有项目
        workbook_1.close()
        xlwings_app_1.quit()
        # 写入总文件 total_xlsx
        workbook_2.save()
        # 删掉不打印的页面
        workbook_2.sheets[7].delete()
        workbook_2.sheets[0].delete()
        printing_path = "..//file//" + user_id + "//output//printing_xlsx//printing_excel.xlsx"
        workbook_2.save(printing_path)
        workbook_2.close()
        xlwings_app_2.quit()
        # webbrowser.open(lpo_file_path + pdf_lpo_name)
        # 合并成payment log里的链接文件
        file_list = os.listdir(invoice_path)
        invoice_filename = file_list[0]
        file_list = os.listdir(lpo_pdf_path)
        lpo_pdf_filename = file_list[0]
        # 一行代码，转换pdf和img
        # pdf2png(invoice_filename, "1")
        to_merge_list = [lpo_pdf_path + lpo_pdf_filename, invoice_path + invoice_filename]
        print(to_merge_list)
        file_merger = PyPDF2.PdfMerger()
        for file in to_merge_list:
            file_merger.append(file)
        paylog_name = "Finance Form " + only_year + project + "-" + name_date + "-" + project
        file_merger.write("..//file//" + user_id + "//output//merged_pdf//" + paylog_name + ".pdf")


class Index(BaseHandler):
    def get(self):
        self.render("index.html", title="Index")


class Input(BaseHandler):
    def post(self, *args, **kwargs):
        user_document = self.get_current_user()
        if user_document:
            user_id = user_document["user_id"]
            # 文件的暂存路径
            upload_path = "..//file//" + user_id
            exist_flag = os.path.exists(upload_path)  # 判断路径是否存在，存在则返回true
            if exist_flag is False:
                os.mkdir(upload_path)
                os.mkdir(upload_path + "//input")
                os.mkdir(upload_path + "//" + "input//lpo")
                os.mkdir(upload_path + "//" + "input//lpo_pdf")
                os.mkdir(upload_path + "//" + "input//tax_invoice")
                os.mkdir(upload_path + "//" + "input//data")
                os.mkdir(upload_path + "//output")
                os.mkdir(upload_path + "//output//merged_pdf")
                os.mkdir(upload_path + "//output//pre_payment_log")
                os.mkdir(upload_path + "//output//printing_xlsx")
                os.mkdir(upload_path + "//output//total_xlsx")
            else:
                shutil.rmtree(upload_path)
                os.mkdir(upload_path)
                os.mkdir(upload_path + "//input")
                os.mkdir(upload_path + "//input//lpo")
                os.mkdir(upload_path + "//input//lpo_pdf")
                os.mkdir(upload_path + "//input//tax_invoice")
                os.mkdir(upload_path + "//input//data")
                os.mkdir(upload_path + "//output")
                os.mkdir(upload_path + "//output//merged_pdf")
                os.mkdir(upload_path + "//output//pre_payment_log")
                os.mkdir(upload_path + "//output//printing_xlsx")
                os.mkdir(upload_path + "//output//total_xlsx")
            # 存储发票号
            invoice_code = str(self.get_argument("invoice_code", ""))
            with open(upload_path + "//input//data//" + invoice_code + ".json", 'wb') as up:
                up.write(invoice_code.encode("utf-8"))
            # 分别处理三个文件
            file_metas = self.request.files  # 提取表单中‘name’为‘file’的文件元数据
            if file_metas is None:
                print("not get xlsx now")
                pass
            else:
                meta = file_metas["lpo_xlsx"][0]
                file_path = upload_path + "//input//lpo//" + meta['filename']
                with open(file_path, 'wb') as up:
                    up.write(meta['body'])
                # 下面是pdf格式文件
                meta = file_metas["lpo_pdf"][0]
                file_path = upload_path + "//input//lpo_pdf//" + meta['filename']
                with open(file_path, 'wb') as up:
                    up.write(meta['body'])
                # 下面是tax_invoice文件
                meta = file_metas["tax_invoice"][0]
                file_path = upload_path + "//input//tax_invoice//" + meta['filename']
                with open(file_path, 'wb') as up:
                    up.write(meta['body'])
            self.redirect("/loading")


class Loading(BaseHandler):
    def get(self):
        self.render("loading.html")

    def post(self):
        user_document = self.get_current_user()
        if user_document:
            user_id = user_document["user_id"]
            # try:
            res_scan = self.get_lpo_information(user_id)
            self.get_LPO(res_scan, user_id)
            self.write({"loading_flag": 0})
            # except Exception as err:
            #     self.write({"loading_flag": 1})


class Output(BaseHandler):
    def get(self):
        self.render("output.html")


class Downloads(BaseHandler):
    def get(self):
        user_document = self.get_current_user()
        if user_document:
            user_id = user_document["user_id"]
            # 文件的暂存路径
            upload_path = "..//file//" + user_id + "//output//"
            file_name_code = int(self.get_argument("filename", None))
            name_list = os.listdir(upload_path)
            file_path = upload_path + name_list[file_name_code-1]
            file_list = os.listdir(file_path)
            file_name = file_list[0]
            # http头 浏览器自动识别为文件下载
            self.set_header('Content-Type', 'application/octet-stream')
            # 下载时显示的文件名称
            self.set_header('Content-Disposition', 'attachment; filename=%s' % parse.quote(file_name))
            try:
                with open(file_path + "//" + file_name, 'rb') as file:
                    while True:
                        data = file.read(1024)
                        if not data:
                            break
                        self.write(data)
            except Exception as err:
                print("Downloads: Error:", err)
                self.finish()


class ErrorHandler(BaseHandler):
    def get(self):
        self.redirect(self.mislead())


def main():

    web_port = 10000
    define("port", default=web_port, help="run on the port " + str(web_port), type=int)

    tornado.options.parse_command_line()
    options.log_file_prefix = "../log/main.log"
    options.log_rotate_mode = "time"
    options.log_rotate_when = "MIDNIGHT"
    options.log_rotate_interval = 30

    http_server = tornado.httpserver.HTTPServer(Application(), xheaders=False, max_buffer_size=2000000000)  # 2G
    http_server.listen(options.port)
    print("server is started.")
    tornado.ioloop.IOLoop.instance().start()


if __name__ == "__main__":
    main()
    # nohup sudo python3 sechat.py >/dev/null 2>&1 &
