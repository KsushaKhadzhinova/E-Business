from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.text.paragraph import Paragraph
import copy, io, os

base_dir = r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
labs = [
    ("ЛР1", "ЛАБОРАТОРНАЯ РАБОТА №1. АНАЛИЗ ПОИСКОВОГО СПРОСА И ФИКСАЦИЯ ГРАНИЦ РЫНКА"),
    ("ЛР2-3", "ЛАБОРАТОРНЫЕ РАБОТЫ №2–3. ОЦЕНКА ОБЪЕМА РЫНКА ЭЛЕКТРОННОГО БИЗНЕСА"),
    ("ЛР4", "ЛАБОРАТОРНАЯ РАБОТА №4. АНАЛИЗ ПЯТИ СИЛ ПОРТЕРА И БАРЬЕРОВ ВХОДА НА РЫНОК"),
    ("ЛР5", "ЛАБОРАТОРНАЯ РАБОТА №5. АНАЛИЗ КОНКУРЕНТОВ ПО САЙТАМ, МОДЕЛЯМ ЛЕВИТТА И КАНО"),
    ("ЛР6", "ЛАБОРАТОРНАЯ РАБОТА №6. ЦЕННОСТНОЕ ПРЕДЛОЖЕНИЕ, БИЗНЕС-МОДЕЛЬ КАНВАС, СЦЕНАРИИ, ЭКРАНЫ И РЕЛИЗЫ"),
    ("ЛР7", "ЛАБОРАТОРНАЯ РАБОТА №7. ИТОГОВЫЙ ОТЧЕТ ПО БИЗНЕС-АНАЛИЗУ ЦИФРОВОГО ТОВАРА"),
]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def src(lr):
    return Document(os.path.join(base_dir, lr, "ОТЧЕТ_%s_NotaCode.docx" % lr))


def style(k):
    s = k.xpath('./w:pPr/w:pStyle/@w:val')
    return s[0] if s else ''


def first_body_idx(ks):
    for i, k in enumerate(ks):
        if i > 15 and k.tag == qn('w:p') and style(k) == '1' and not k.xpath('.//w:fldChar|.//w:instrText'):
            return i
    raise SystemExit("no body start")


base = src("ЛР1")
body = base.element.body
kids = list(body.iterchildren())
b0 = first_body_idx(kids)
for k in kids[16:b0]:
    body.remove(k)
for k in kids[b0:]:
    if k.tag != qn('w:sectPr'):
        body.remove(k)
sect = body.find(qn('w:sectPr'))


def settext(el, text):
    p = Paragraph(el, None)
    runs = p.runs
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


settext(kids[7], "по лабораторным работам №1–7")
settext(kids[10], "на тему «Бизнес-анализ цифрового продукта NotaCode»")

toc_xml = (
    '<w:p xmlns:w="%s"><w:pPr><w:pStyle w:val="10"/></w:pPr>'
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    '<w:r><w:t>Правой кнопкой мыши – «Обновить поле», чтобы построить содержание.</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' % W
)
sect.addprevious(parse_xml(toc_xml))


def para(text, sty, pb=False):
    ppr = '<w:pPr><w:pStyle w:val="%s"/>%s</w:pPr>' % (sty, '<w:pageBreakBefore/>' if pb else '')
    return parse_xml('<w:p xmlns:w="%s">%s<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (W, ppr, text))


sect.addprevious(para("ВВЕДЕНИЕ", "1", True))
sect.addprevious(para(
    "Итоговый отчет объединяет отчеты по лабораторным работам №1–7 по дисциплине «Электронный бизнес» "
    "для проекта NotaCode. Каждая часть сохраняет собственную нумерацию разделов, таблиц, рисунков и "
    "приложений; данные частей согласованы с рабочими книгами Excel.", "FirstParagraph"))

docpr = [1000]


def clean(el):
    for t in ("w:bookmarkStart", "w:bookmarkEnd", "w:lastRenderedPageBreak"):
        for x in el.xpath(".//" + t):
            x.getparent().remove(x)
    for x in el.iter():
        for a in list(x.attrib):
            if a.endswith("}paraId") or a.endswith("}textId"):
                del x.attrib[a]


def remap_images(el, sdoc):
    for blip in el.xpath(".//a:blip"):
        rid = blip.get(qn("r:embed"))
        if rid:
            blob = sdoc.part.related_parts[rid].blob
            nrid, _ = base.part.get_or_add_image(io.BytesIO(blob))
            blip.set(qn("r:embed"), nrid)
    for dp in el.xpath(".//wp:docPr"):
        docpr[0] += 1
        dp.set("id", str(docpr[0]))


total = 0
for lr, title in labs:
    sd = src(lr)
    sk = list(sd.element.body.iterchildren())
    i0 = first_body_idx(sk)
    sect.addprevious(para(title, "1", True))
    n = 0
    for k in sk[i0:]:
        if k.tag == qn('w:sectPr'):
            continue
        e = copy.deepcopy(k)
        if e.tag == qn('w:p'):
            s = style(e)
            if s in ("1", "2", "3", "4", "5"):
                e.xpath('./w:pPr/w:pStyle')[0].set(qn('w:val'), str(int(s) + 1))
        clean(e)
        remap_images(e, sd)
        sect.addprevious(e)
        n += 1
    total += n
    print(lr, "elements", n)

st = base.settings.element
if st.find(qn("w:updateFields")) is None:
    u = OxmlElement("w:updateFields")
    u.set(qn("w:val"), "true")
    st.append(u)
out = os.path.join(base_dir, "Финальный_отчет", "ОТЧЕТ_по_лабораторным_работам_1-7_NotaCode.docx")
base.save(out)
print("saved", out, total)
