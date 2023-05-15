import os
import json
import codecs
import datetime

from flask import Flask
from flask import request
from flask import send_file

from src.document import MyDocument
from win32_utils import docx_handle

app = Flask(__name__)

def get_time():
    d_time = datetime.datetime.now()
    d_time = "{}-{:0>2d}-{:0>2d}-{:0>2d}-{:0>2d}-{:0>2d}".format(d_time.year, d_time.month, d_time.day, d_time.hour, d_time.minute, d_time.second)
    return "%s.docx" %d_time


def dict_to_json(data_dict, json_file, indent=4, ensure_ascii=False):
    f = codecs.open(json_file, "w", encoding="utf-8")
    json.dump(data_dict, f, indent=indent, ensure_ascii=ensure_ascii)
    f.close()
    return True


@app.route("/json_to_docx", methods = ['POST'])
def json_to_docx():
    if request.method == 'POST':
        data = request.json
        now = get_time()
        out_dir = "docx"
        out_docx = os.path.join(out_dir, now + ".docx")
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)

        # 保存推送过来的数据
        dict_to_json(data, os.path.join(out_dir, now + ".json"))

        # 渲染docx文件
        document = MyDocument("default.docx")
        for section in document.document.sections:
            section.left_margin = 720000
            section.right_margin = 720000
        document.render(data, out_docx)
        for i in data:
            if not i:
                continue
            if i["type"] == "table_of_contents":
                docx_handle(out_docx, update_toc=True)
                break
        return send_file(out_docx, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/docx_to_pdf", methods = ['POST'])
def docx_to_pdf():
    if request.method == 'POST':
        data = request.files["file"]
        now = get_time()
        out_dir = "pdf"
        out_docx = os.path.abspath(os.path.join(out_dir, now + ".docx"))
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)

        data.save(out_docx)

        out_pdf = out_docx.replace(".docx", ".pdf")
        docx_handle(out_docx, update_toc=False, pdf_file=out_pdf)
        return send_file(out_pdf, mimetype="application/pdf")


if __name__ == '__main__':
    app.run(host='0.0.0.0',port=5556,debug=True)
