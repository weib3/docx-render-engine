import base64
import json
import codecs
import copy
import random
from io import BytesIO

# from lxml import etree
import bs4
from docx.shared import Pt
from PIL import Image


def json_to_dict(json_file, encoding="utf-8"):
    f = open(json_file, "r", encoding=encoding)
    data_dict = json.load(f)
    f.close()
    return data_dict


def dict_to_json(data_dict, json_file, indent=4, ensure_ascii=False):
    f = codecs.open(json_file, "w", encoding="utf-8")
    json.dump(data_dict, f, indent=indent, ensure_ascii=ensure_ascii)
    f.close()
    return True

def hex_to_rgb(color):
    """
    Usage:
        hex_to_rgb("#1F5C8B") will return (31, 92, 139).
    """
    color = color.strip("#")
    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


def rgb_to_hex(rgb):
    """
    Usage:
        rgb_to_hex((31, 92, 139)) will return "#1f5c8b".
    """
    return "#%02x%02x%02x" %rgb


def bool_str(instr):
    instr = str(instr)
    if instr.lower() in ["true"]:
        return True
    elif instr.lower() in ["false"]:
        return False
    else:
        raise Exception("str must be T[t]rue or F[f]alse, but input is %s" %instr)


def parse_from_xml_like_text(text):
    def parse_node(node, parent_attrib=None):
        return_list = []
        if parent_attrib is None:
            parent_attrib = {}
        if isinstance(node, bs4.element.NavigableString):
            return_list.append({
                "tag": "r",
                "text": str(node),
                "style": parent_attrib
            })
        else:
            if not node.contents:  #如果是空元素，保留取当前node tag
                current_attrs = copy.deepcopy(parent_attrib)
                current_attrs.update(dict(node.attrs))
                return_list.append({
                    "tag": node.name,
                    "text": "",
                    "style": current_attrs
                })
            else:
                for i in node.contents:
                    current_attrs = copy.deepcopy(parent_attrib)
                    current_attrs.update(dict(node.attrs))
                    return_list.extend(
                        parse_node(i, current_attrs)
                    )
        return return_list

    if not text.startswith(("<p ", "<p>")):  # 如果没有<p> 标签，手动加上
        text = "<p>%s</p>" % text
    node = bs4.BeautifulSoup(text, features="lxml")
    return parse_node(node)


# def parse_from_xml_like_text(text):
#     return_list = []
#     if not text.startswith(("<p ", "<p>")):  # 如果没有<p> 标签，手动加上
#         text = "<p>%s</p>" % text
#     #  字符串中有&符号会报错
#     tree = etree.fromstring(text)
#     for i in tree.iter():
#         child_attrib = dict(i.attrib)
#         tree_attrib = dict(tree.attrib)
#         [child_attrib.setdefault(_, tree_attrib[_]) for _ in tree_attrib]
#         return_list.append({
#             "tag": i.tag,
#             "text": i.text if i.text else "",
#             "style": child_attrib
#         })
#         if i.tail:
#             return_list.append({
#                 "tag": "r",
#                 "text": i.tail,
#                 "style": tree_attrib
#             })
#     return return_list

def point_unit_guess(value):
    """
    猜测单位是否是point
    """
    if -12700 < value < 12700:  #Pt(1)
        return Pt(value)
    return value


def is_digit(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


def str_to_numeric(in_str):
    if not isinstance(in_str, str):
        return in_str
    if "%" in in_str:
        return float(in_str.strip('%')) / 100
    elif "." in in_str:
        return float(in_str)
    else:
        return int(in_str)


def string_to_image(image_str, outfile, file_type="JPEG"):
    im = Image.open(BytesIO(base64.b64decode(image_str)))
    im.save(outfile, file_type)
    return True


def random_file_name():
    return "".join(random.sample("abcdefghijklmnopqrstuvwxyz", 6))
