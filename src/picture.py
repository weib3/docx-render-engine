import os
import tempfile

from docx.text.run import Run
from docx.oxml import parse_xml, register_element_cls, OxmlElement
from docx.oxml.ns import nsdecls
from docx.oxml.shape import CT_Picture
from docx.oxml.xmlchemy import BaseOxmlElement, OneAndOnlyOne, ZeroOrOneChoice
from lxml import etree

from .utils import point_unit_guess, str_to_numeric, string_to_image, random_file_name


class MyInlinePicture(object):
    def __init__(self, picture, parent: Run):
        self.picture = picture
        self.parent = parent

    @classmethod
    def create(cls, src: str, run: Run, **kwargs):
        width = None
        height = None
        if "width" in kwargs:
            width = point_unit_guess(str_to_numeric(kwargs["width"]))
        if "height" in kwargs:
            height = point_unit_guess(str_to_numeric(kwargs["height"]))

        # 处理base64，格式为：base64;png;xxxxxxxx
        if src.startswith("base64"):
            tmp = src.split(";")
            file_type = tmp[1]
            base64_string = ";".join(tmp[2:])
            tmp_file = os.path.join(
                tempfile.gettempdir(),
                random_file_name() + "." + file_type.lower()
            )
            string_to_image(base64_string, tmp_file, file_type=file_type.upper())
            src = tmp_file
            
        pic = run.add_picture(src, width=width, height=height)
        return cls(pic, run)

    def set_style(self, **kwargs):
        if "width" in kwargs:
            self.picture.width = point_unit_guess(str_to_numeric(kwargs["width"]))
        if "height" in kwargs:
            self.picture.height = point_unit_guess(str_to_numeric(kwargs["height"]))


class MyFloatPicture(object):
    """
    http://officeopenxml.com/drwPicFloating-position.php
    http://officeopenxml.com/drwPicFloating.php
    """
    def __init__(self, picture, parent: Run):
        self.picture = picture
        self.parent = parent

    @classmethod
    def create(cls, src: str, run: Run, **kwargs):
        width = kwargs.get("width", None)
        height = kwargs.get("height", None)
        if width is not None:
            width = point_unit_guess(str_to_numeric(kwargs["width"]))
        if height is not None:
            height = point_unit_guess(str_to_numeric(kwargs["height"]))
        pos_x = point_unit_guess(str_to_numeric(kwargs.get("pos_x", 0)))
        pos_y = point_unit_guess(str_to_numeric(kwargs.get("pos_y", 0)))
        anchor_x = kwargs.get("anchor_x", "page")
        anchor_y = kwargs.get("anchor_y", "page")
        text_wrap = kwargs.get("text_wrap", "wrapTopAndBottom")

        # 处理base64，格式为：base64;png;xxxxxxxx
        if src.startswith("base64"):
            tmp = src.split(";")
            file_type = tmp[1]
            base64_string = ";".join(tmp[2:])
            tmp_file = os.path.join(
                tempfile.gettempdir(),
                random_file_name() + "." + file_type.lower()
            )
            string_to_image(base64_string, tmp_file, file_type=file_type.upper())
            src = tmp_file

        rId, image = run.part.get_or_add_image(src)
        cx, cy = image.scaled_dimensions(width, height)
        shape_id, filename = run.part.next_id, image.filename
        anchor = CT_Anchor.new_pic_anchor(shape_id, rId, filename, cx, cy,
                                          pos_x, pos_y, anchor_x, anchor_y, text_wrap)
        run._r.add_drawing(anchor)
        return cls(FloatShape(anchor), run)

    def set_style(self, **kwargs):
        if "width" in kwargs:
            self.picture.width = point_unit_guess(str_to_numeric(kwargs["width"]))
        if "height" in kwargs:
            self.picture.height = point_unit_guess(str_to_numeric(kwargs["height"]))
        if "pos_x" in kwargs:
            self.picture.pos_x = point_unit_guess(str_to_numeric(kwargs["pos_x"]))
        if "pos_y" in kwargs:
            self.picture.pos_y = point_unit_guess(str_to_numeric(kwargs["pos_y"]))
        if "anchor_x" in kwargs:
            self.picture.anchor_x = kwargs["anchor_x"]
        if "anchor_y" in kwargs:
            self.picture.anchor_y = kwargs["anchor_y"]


class FloatShape(object):
    """
    Refer to docx.shape.InlineShape
    """
    def __init__(self, float):
        super(FloatShape, self).__init__()
        self._float = float

    @property
    def height(self):
        return self._float.extent.cy

    @height.setter
    def height(self, cy):
        self._float.extent.cy = cy
        self._float.graphic.graphicData.pic.spPr.cy = cy

    @property
    def width(self):
        return self._float.extent.cx

    @width.setter
    def width(self, cx):
        self._float.extent.cx = cx
        self._float.graphic.graphicData.pic.spPr.cx = cx

    @property
    def pos_x(self):
        return str_to_numeric(self._float.positionH.first_child_found_in("wp:posOffset").text)

    @pos_x.setter
    def pos_x(self, value):
        self._float.positionH.first_child_found_in("wp:posOffset").text = str(value)

    @property
    def pos_y(self):
        return str_to_numeric(self._float.positionV.first_child_found_in("wp:posOffset").text)

    @pos_y.setter
    def pos_y(self, value):
        self._float.positionV.first_child_found_in("wp:posOffset").text = str(value)

    @property
    def anchor_x(self):
        return self._float.positionH.attrib["relativeFrom"]

    @anchor_x.setter
    def anchor_x(self, value):
        self._float.positionH.attrib["relativeFrom"] = value

    @property
    def anchor_y(self):
        return self._float.positionV.attrib["relativeFrom"]

    @anchor_y.setter
    def anchor_y(self, value):
        self._float.positionV.attrib["relativeFrom"] = value

    @property
    def text_wrap(self):
        res = self._float.first_child_found_in(
            "wp:wrapNone", "wp:wrapSquare", "wp:wrapThrough", "wp:wrapTight", "wp:wrapTopAndBottom"
        )
        return etree.QName(res).localname

    @text_wrap.setter
    def text_wrap(self, value):
        old_element = self._float.first_child_found_in(
            "wp:wrapNone", "wp:wrapSquare", "wp:wrapThrough", "wp:wrapTight", "wp:wrapTopAndBottom"
        )
        new_element = OxmlElement("wp:%s" %value)
        self._float.replace(old_element, new_element)


# refered from https://github.com/dothinking/pdf2docx/issues/54
class CT_Anchor(BaseOxmlElement):
    """
    ``<w:anchor>`` element, container for a floating image.
    """
    extent = OneAndOnlyOne('wp:extent')
    docPr = OneAndOnlyOne('wp:docPr')
    graphic = OneAndOnlyOne('a:graphic')
    positionH = OneAndOnlyOne('wp:positionH')
    positionV = OneAndOnlyOne('wp:positionV')

    @classmethod
    def new(cls, cx, cy, shape_id, pic, pos_x, pos_y, anchor_x, anchor_y, text_wrap):
        """
        Return a new ``<wp:anchor>`` element populated with the values passed
        as parameters.
        """
        anchor = parse_xml(cls._anchor_xml(pos_x, pos_y, anchor_x, anchor_y, text_wrap))
        anchor.extent.cx = cx
        anchor.extent.cy = cy
        anchor.docPr.id = shape_id
        anchor.docPr.name = 'Picture %d' % shape_id
        anchor.graphic.graphicData.uri = (
            'http://schemas.openxmlformats.org/drawingml/2006/picture'
        )
        anchor.graphic.graphicData._insert_pic(pic)
        return anchor

    @classmethod
    def new_pic_anchor(cls, shape_id, rId, filename, cx, cy, pos_x, pos_y,
                       anchor_x="page", anchor_y="page", text_wrap="wrapTopAndBottom"):
        """
        Return a new `wp:anchor` element containing the `pic:pic` element
        specified by the argument values.
        """
        pic_id = 0  # Word doesn't seem to use this, but does not omit it
        pic = CT_Picture.new(pic_id, filename, rId, cx, cy)
        anchor = cls.new(cx, cy, shape_id, pic, pos_x, pos_y, anchor_x, anchor_y, text_wrap)
        return anchor

    @classmethod
    def _anchor_xml(cls, pos_x, pos_y, anchor_x, anchor_y, text_wrap):
        return (
            '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" \n'
            '           behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1" \n'
            '           %s>\n'
            '  <wp:simplePos x="0" y="0"/>\n'
            '  <wp:positionH relativeFrom="%s">\n'
            '    <wp:posOffset>%d</wp:posOffset>\n'
            '  </wp:positionH>\n'
            '  <wp:positionV relativeFrom="%s">\n'
            '    <wp:posOffset>%d</wp:posOffset>\n'
            '  </wp:positionV>\n'                    
            '  <wp:extent cx="7556500" cy="10693400"/>\n'
            '  <wp:%s />'
            '  <wp:docPr id="666" name="unnamed"/>\n'
            '  <wp:cNvGraphicFramePr>\n'
            '    <a:graphicFrameLocks noChangeAspect="1"/>\n'
            '  </wp:cNvGraphicFramePr>\n'
            '  <a:graphic>\n'
            '    <a:graphicData uri="URI not set"/>\n'
            '  </a:graphic>\n'
            '</wp:anchor>' % (nsdecls('wp', 'a', 'pic', 'r'),
                              anchor_x, str_to_numeric(pos_x),
                              anchor_y, str_to_numeric(pos_y),
                              text_wrap
                              )
        )


class CT_PostionH(BaseOxmlElement):
    value = OneAndOnlyOne('wp:positionH')

class CT_PostionV(BaseOxmlElement):
    value = OneAndOnlyOne('wp:positionV')


register_element_cls('wp:anchor', CT_Anchor)
register_element_cls('wp:positionH', CT_PostionH)
register_element_cls('wp:positionV', CT_PostionV)
