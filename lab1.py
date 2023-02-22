import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor

import mtk2

doc = docx.Document("C:\\Users\\maxbe\\Downloads\\variant10.docx")
paragraphs = doc.paragraphs


def encode():
    count_color = 0
    count_size = 0
    count_highlight = 0
    count_spacing = 0
    count_scale = 0

    def get_scale(run):
        return run._r.get_or_add_rPr().xpath("./w:w")

    def get_spacing(run):
        return run._r.get_or_add_rPr().xpath("./w:spacing")

    code = ""
    for paragraph in paragraphs:
        for run in paragraph.runs:
            font_color = run.font.color.rgb
            font_highlight_color = run.font.highlight_color
            font_size = run.font.size
            font_scale = get_scale(run)
            font_spacing = get_spacing(run)

            if (font_color != RGBColor(0, 0, 0) or
                    font_size.pt != 12.0 or
                    font_highlight_color != WD_COLOR_INDEX.WHITE or
                    font_spacing != [] or
                    font_scale != []):
                if font_color != RGBColor(0, 0, 0):
                    count_color += 1
                if font_size.pt != 12.0:
                    count_size += 1
                if font_highlight_color != WD_COLOR_INDEX.WHITE:
                    count_highlight += 1
                if font_spacing != []:
                    count_spacing += 1
                if font_scale != []:
                    count_scale += 1
                for i in range(len(run.text)):
                    code += '1'
            else:
                for i in range(len(run.text)):
                    code += '0'
    method = max(count_scale, count_spacing, count_highlight, count_size, count_color)
    
    if (method == count_size):
        print("Способ форматирования: по размеру шрифта")
    while len(code) % 8 != 0:
        code += "0"

    return code


class Charset:
    def __init__(self, _code):
        self.code = _code

    def mtk2_decode(self):
        return mtk2.MTK2_decode(self.code)

    def koi8_decode(self):
        return bytes.fromhex(hex(int(self.code, 2))[2:]).decode(encoding="koi8_r")

    def cp866_decode(self):
        return bytes.fromhex(hex(int(self.code, 2))[2:]).decode(encoding="cp866")

    def cp1251_decode(self):
        return bytes.fromhex(hex(int(self.code, 2))[2:]).decode(encoding="cp1251")


def main():
    charset = Charset(encode())
    print("            MTK2:\n\n", charset.mtk2_decode(), "\n\n")
    print("            KOI8-R:\n\n", charset.koi8_decode(), "\n\n")
    print("            CP866:\n\n", charset.cp866_decode(), "\n\n")
    print("            CP1251:\n\n", charset.cp1251_decode())


if __name__ == '__main__':
    main()
