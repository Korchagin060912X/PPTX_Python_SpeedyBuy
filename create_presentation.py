from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import os

# --- НАСТРОЙКИ ЦВЕТОВ ---
BG_COLOR = RGBColor(13, 27, 74)       # Темно-синий фон
CARD_COLOR = RGBColor(37, 99, 235)    # Голубой цвет плашек
TEXT_COLOR = RGBColor(255, 255, 255)  # Белый текст
SC_COLOR = RGBColor(30, 58, 138)      # Цвет надписей SC на фоне
OUTLINE_COLOR = RGBColor(30, 58, 138) # Темно-голубая обводка текста

# --- СОЗДАНИЕ ПРЕЗЕНТАЦИИ ---
prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def set_text_transparency(run, transparency_percent):
    """Устанавливает прозрачность текста через lxml"""
    rPr = run._r.get_or_add_rPr()
    solidFill = rPr.find(qn('a:solidFill'))
    if solidFill is None:
        solidFill = etree.SubElement(rPr, qn('a:solidFill'))
    
    srgbClr = solidFill.find(qn('a:srgbClr'))
    if srgbClr is None:
        srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
        
    alpha_val = int((100 - transparency_percent) * 1000)
    alpha = srgbClr.find(qn('a:alpha'))
    if alpha is None:
        alpha = etree.SubElement(srgbClr, qn('a:alpha'))
    alpha.set('val', str(alpha_val))

def set_text_outline(run, color, width_pt):
    """Добавляет обводку к тексту через XML"""
    rPr = run._r.get_or_add_rPr()
    old_ln = rPr.find(qn('a:ln'))
    if old_ln is not None:
        rPr.remove(old_ln)
        
    ln = etree.SubElement(rPr, qn('a:ln'))
    ln.set('w', str(int(width_pt * 12700))) 
    
    solidFill = etree.SubElement(ln, qn('a:solidFill'))
    srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
    
    # Исправлено: RGBColor это кортеж, обращаемся по индексам
    srgbClr.set('val', f"{color[0]:02X}{color[1]:02X}{color[2]:02X}")

def add_sc_background(slide):
    """Добавляет фоновые надписи SC"""
    sc_positions = [
        (1, 0.5, 1.5), (10, 1, 2), (5, 5, 1.2), (11, 5.5, 1.8), 
        (3, 6, 1.5), (8, 0.5, 1), (0.5, 3, 1), (12, 3, 1.5),
        (6, 2, 2.5), (2, 2, 1.2)
    ]
    for left, top, size in sc_positions:
        txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(size), Inches(size * 0.6))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = "SC"
        p.font.size = Pt(int(size * 40))
        p.font.bold = True
        p.font.color.rgb = SC_COLOR
        p.alignment = PP_ALIGN.CENTER

def add_blue_card(slide, text, left, top, width, height, font_size=18):
    """Добавляет голубую плашку с закругленными краями и текстом"""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = CARD_COLOR
    shape.line.fill.background() 
    
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.2)
    tf.margin_bottom = Inches(0.2)
    
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = TEXT_COLOR
    p.alignment = PP_ALIGN.LEFT

def add_author_signature(slide):
    """Добавляет подпись с прозрачностью 80%"""
    txBox = slide.shapes.add_textbox(Inches(4.5), Inches(7.0), Inches(4.33), Inches(0.4))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    
    run = p.add_run()
    run.text = "Сделал Korchagin060912X"
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(255, 255, 255)
    run.font.italic = True
    set_text_transparency(run, 80)

def add_title(slide, text, left, top, width):
    """Добавляет большой заголовок"""
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(1))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = TEXT_COLOR
    p.alignment = PP_ALIGN.CENTER # Заголовки по центру

def set_slide_bg(slide):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

# ==========================================
# СЛАЙД 1
# ==========================================
slide1 = prs.slides.add_slide(blank_layout)
set_slide_bg(slide1)
add_sc_background(slide1)

if os.path.exists("banner.png"):
    slide1.shapes.add_picture("banner.png", Inches(0.5), Inches(1), width=Inches(6))
else:
    placeholder = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1), Inches(6), Inches(4))
    placeholder.fill.solid()
    placeholder.fill.fore_color.rgb = RGBColor(20, 40, 100)
    tf = placeholder.text_frame
    tf.text = "ВСТАВЬТЕ banner.png"
    tf.paragraphs[0].font.color.rgb = TEXT_COLOR
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

add_title(slide1, "SpeedyBuy - Цифровые товары", 7, 1.2, 5.5)
add_blue_card(slide1, "Твой удобный магазин цифровых товаров: ключи, игры, подписки и популярные аксессуары для геймеров.", 7, 3.0, 5.5, 1.8, font_size=18)
add_blue_card(slide1, "Сайт на Flask и Python!", 7, 5.0, 5.5, 1.2, font_size=18)
add_author_signature(slide1)

# ==========================================
# СЛАЙД 2
# ==========================================
slide2 = prs.slides.add_slide(blank_layout)
set_slide_bg(slide2)
add_sc_background(slide2)

add_title(slide2, "Технологический стек", 1, 1.2, 11.33)
tech_text = "Сайт был сделан на библиотеке Flask 3.0.3 и с Django админкой, дизайн на CSS и сайт на Языке программирования Python и html!"
add_blue_card(slide2, tech_text, 1, 3.0, 11.33, 3.0, font_size=22)
add_author_signature(slide2)

# ==========================================
# СЛАЙД 3
# ==========================================
slide3 = prs.slides.add_slide(blank_layout)
set_slide_bg(slide3)
add_sc_background(slide3)

for i, img_name in enumerate(["banner1.png", "banner2.png"]):
    y_pos = 1.0 if i == 0 else 4.0
    if os.path.exists(img_name):
        slide3.shapes.add_picture(img_name, Inches(0.5), Inches(y_pos), width=Inches(5.5))
    else:
        placeholder = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(y_pos), Inches(5.5), Inches(2.8))
        placeholder.fill.solid()
        placeholder.fill.fore_color.rgb = RGBColor(20, 40, 100)
        tf = placeholder.text_frame
        tf.text = f"ВСТАВЬТЕ {img_name}"
        tf.paragraphs[0].font.color.rgb = TEXT_COLOR
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER

txBox = slide3.shapes.add_textbox(Inches(7), Inches(2), Inches(5.5), Inches(3.5))
tf = txBox.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Архитектура и код проекта"
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = TEXT_COLOR
p.alignment = PP_ALIGN.CENTER
run1 = p.runs[0] if p.runs else p.add_run()
set_text_outline(run1, OUTLINE_COLOR, 2.5)

p2 = tf.add_paragraph()
p2.text = "Чистый и структурированный код на Python. Современные фреймворки обеспечивают высокую скорость работы, масштабируемость и удобство поддержки."
p2.font.size = Pt(22)
p2.font.color.rgb = TEXT_COLOR
p2.alignment = PP_ALIGN.CENTER
run2 = p2.runs[0] if p2.runs else p2.add_run()
set_text_outline(run2, OUTLINE_COLOR, 1.5)

add_author_signature(slide3)

# ==========================================
# СЛАЙД 4
# ==========================================
slide4 = prs.slides.add_slide(blank_layout)
set_slide_bg(slide4)
add_sc_background(slide4)

add_title(slide4, "Запуск и настройка сайта", 1, 1.2, 11.33)

text1 = "Сайт работает на виртуальном окружении Venv. Активация происходит в терминале, после чего необходимо загрузить все библиотеки через файл requirements.txt."
add_blue_card(slide4, text1, 1, 2.5, 11.33, 1.8, font_size=20)

text2 = "После активации Venv и установки библиотек, запустите сайт командой python app.py и откройте в браузере: http://127.0.0.1:5000"
add_blue_card(slide4, text2, 1, 4.6, 11.33, 1.8, font_size=20)

add_author_signature(slide4)

# ==========================================
# СЛАЙД 5
# ==========================================
slide5 = prs.slides.add_slide(blank_layout)
set_slide_bg(slide5)
add_sc_background(slide5)

add_title(slide5, "Функционал SpeedyBuy", 1, 1.2, 11.33)

list_text = (
    "Что реализовано:\n\n"
    "• Каталог товаров с фильтрацией по категориям\n"
    "• Добавление товара и просмотр корзины\n"
    "• Оформление заказа с формой клиента\n"
    "• Личный кабинет: история покупок и число заказов\n"
    "• Админ-панель: управление заказами, товарами, пользователями\n"
    "• Редактирование профиля: никнейм, пароль, настройки"
)
add_blue_card(slide5, list_text, 1, 2.5, 5.5, 4.0, font_size=18)

if os.path.exists("banner3.png"):
    slide5.shapes.add_picture("banner3.png", Inches(7), Inches(2), width=Inches(5.5))
else:
    placeholder = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7), Inches(2), Inches(5.5), Inches(4))
    placeholder.fill.solid()
    placeholder.fill.fore_color.rgb = RGBColor(20, 40, 100)
    tf = placeholder.text_frame
    tf.text = "ВСТАВЬТЕ banner3.png"
    tf.paragraphs[0].font.color.rgb = TEXT_COLOR
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

add_author_signature(slide5)

# ==========================================
# СЛАЙД 6 (НОВЫЙ: 4 чекбокса)
# ==========================================
slide6 = prs.slides.add_slide(blank_layout)
set_slide_bg(slide6)
add_sc_background(slide6)

add_title(slide6, "Что есть на сайте:", 1, 0.8, 11.33)

# Сетка 2x2 для чекбоксов
add_blue_card(slide6, "Есть переход на страницы Главной, FAQ, Маркетплейс, Поддержка и Профиль с корзиной!", 1.0, 2.2, 5.5, 1.8, font_size=18)
add_blue_card(slide6, "Также имеется Админ Панель, для добавления товаров, изменения товаров, информации заказов, редактирование пользователей и чаты с поддержкой!", 7.0, 2.2, 5.5, 1.8, font_size=18)
add_blue_card(slide6, "Есть чат с Поддержкой, где пользователь может задать вопрос и получить ответ на него!", 1.0, 4.3, 5.5, 1.8, font_size=18)
add_blue_card(slide6, "Также есть корзина через которую можно совершить заказ товара(или товаров)!", 7.0, 4.3, 5.5, 1.8, font_size=18)

add_author_signature(slide6)

# ==========================================
# СЛАЙД 7 (НОВЫЙ: Под конец)
# ==========================================
slide7 = prs.slides.add_slide(blank_layout)
set_slide_bg(slide7)
add_sc_background(slide7)

add_title(slide7, "Под конец.", 1, 0.8, 11.33)

# Левая часть: Картинка по центру левой половины
if os.path.exists("banner4.png"):
    slide7.shapes.add_picture("banner4.png", Inches(1.0), Inches(2.5), width=Inches(5.5))
else:
    placeholder = slide7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(2.5), Inches(5.5), Inches(3.5))
    placeholder.fill.solid()
    placeholder.fill.fore_color.rgb = RGBColor(20, 40, 100)
    tf = placeholder.text_frame
    tf.text = "ВСТАВЬТЕ banner4.png"
    tf.paragraphs[0].font.color.rgb = TEXT_COLOR
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# Правая часть: 2 чекбокса
add_blue_card(slide7, "В данном проекте в общем получилось около 2500-2700 строчек кода!", 7.0, 2.5, 5.5, 1.5, font_size=20)
add_blue_card(slide7, "Сайт был создан с участием Нейросетей, в качестве помощника.", 7.0, 4.3, 5.5, 1.5, font_size=20)

add_author_signature(slide7)

# ==========================================
# СЛАЙДЫ 8-10 (Финальные)
# ==========================================
slides_content = [
    ("Главные преимущества", "Покупай быстро и дешево. Моментальная доставка ключей на электронную почту."),
    ("Безопасность", "Надежная защита транзакций и конфиденциальность данных каждого пользователя."),
    ("Спасибо за внимание!", "Заходи на SpeedyBuy и оцени все преимущества цифрового шопинга!")
]

for title, text in slides_content:
    slide = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide)
    add_sc_background(slide)
    
    add_title(slide, title, 1, 1.5, 11.33)
    
    txBox = slide.shapes.add_textbox(Inches(1), Inches(3.0), Inches(11.33), Inches(3))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(24)
    p.font.color.rgb = TEXT_COLOR
    p.alignment = PP_ALIGN.CENTER
    
    run = p.runs[0] if p.runs else p.add_run()
    set_text_outline(run, OUTLINE_COLOR, 1.5)
    
    add_author_signature(slide)

# --- СОХРАНЕНИЕ ---
prs.save('SpeedyBuy_Presentation.pptx')
print("Презентация успешно создана: SpeedyBuy_Presentation.pptx")