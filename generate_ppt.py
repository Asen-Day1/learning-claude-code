"""
Generate a professional PPT for Claude Code installation tutorial.
Includes embedded images with proper scaling and clean layout.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import os

# Constants
IMG_DIR = r"D:\Sunday\Learning\A 安装教程"
OUTPUT = os.path.join(IMG_DIR, "Claude_Code_安装教程.pptx")
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# Color palette
BG_DARK   = RGBColor(0x0D, 0x11, 0x17)   # near-black
BG_CARD   = RGBColor(0x16, 0x1B, 0x22)   # dark card
BG_CODE   = RGBColor(0x1E, 0x24, 0x2C)   # code block bg
ACCENT    = RGBColor(0xE9, 0x54, 0x3E)    # warm red
ACCENT2   = RGBColor(0x38, 0xA1, 0x69)    # green
ACCENT3   = RGBColor(0x3B, 0x82, 0xF6)    # blue
GOLD      = RGBColor(0xF5, 0x9E, 0x0B)    # amber/gold
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
GRAY1     = RGBColor(0xE0, 0xE0, 0xE0)    # light gray
GRAY2     = RGBColor(0x94, 0xA3, 0xB8)    # mid gray
GRAY3     = RGBColor(0x64, 0x74, 0x8B)    # darker gray
BORDER    = RGBColor(0x2D, 0x33, 0x3B)    # subtle border
SUBTLE    = RGBColor(0x1A, 0x20, 0x28)    # slightly lighter than bg

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H


# ── Helper functions ───────────────────────────────────────────

def bg(slide, color=BG_DARK):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def rect(slide, l, t, w, h, color, radius=None):
    """Add a filled rectangle, optionally rounded."""
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                               l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s

def accent_line(slide, l, t, w, h=Inches(0.05), color=ACCENT):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = color; s.line.fill.background()
    return s


def tb(slide, l, t, w, h, text, size=16, color=WHITE, bold=False, align=PP_ALIGN.LEFT,
       font='Microsoft YaHei', v_anchor=MSO_ANCHOR.TOP, line_spacing=1.2):
    """Add a single-line or multi-line text box. Returns the shape."""
    box = slide.shapes.add_textbox(l, t, w, h)
    box.text_frame.word_wrap = True
    box.text_frame.auto_size = None
    tf = box.text_frame
    tf.paragraphs[0].alignment = align
    # line spacing
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.font.name = font
        p.alignment = align
        p.space_after = Pt(2)
        if line_spacing != 1.0:
            p.line_spacing = Pt(size * line_spacing)
    return box


def bullets(slide, l, t, w, h, items, size=16, color=GRAY1, bullet='▸', spacing=Pt(6)):
    """Add bullet-point text."""
    box = slide.shapes.add_textbox(l, t, w, h)
    box.text_frame.word_wrap = True
    tf = box.text_frame
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f'{bullet} {item}'
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.name = 'Microsoft YaHei'
        p.space_after = spacing
        p.line_spacing = Pt(size * 1.35)
    return box


def code_block(slide, l, t, w, h, text, size=14):
    """Styled code block with dark bg."""
    rect(slide, l, t, w, h, BG_CODE)
    box = slide.shapes.add_textbox(l + Inches(0.25), t + Inches(0.15),
                                   w - Inches(0.5), h - Inches(0.3))
    box.text_frame.word_wrap = True
    tf = box.text_frame
    for i, line in enumerate(text.strip().split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(size)
        p.font.color.rgb = ACCENT2
        p.font.name = 'Consolas'
        p.line_spacing = Pt(size * 1.35)
    return box


def step_circle(slide, l, t, num, size=0.55):
    """Numbered circle."""
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, Inches(size), Inches(size))
    c.fill.solid(); c.fill.fore_color.rgb = ACCENT; c.line.fill.background()
    p = c.text_frame.paragraphs[0]
    p.text = str(num); p.font.size = Pt(int(size * 33)); p.font.bold = True
    p.font.color.rgb = WHITE; p.font.name = 'Microsoft YaHei'; p.alignment = PP_ALIGN.CENTER
    return c


def fit_image(slide, path, target_l, target_t, target_w, target_h):
    """Place an image scaled to fit within target box, preserving aspect ratio.
    Returns the placed shape."""
    if not os.path.exists(path):
        tb(slide, target_l, target_t + target_h/4, target_w, Inches(0.4),
           f'[图片未找到]\n{os.path.basename(path)}', size=11, color=GRAY3, align=PP_ALIGN.CENTER)
        return None
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(target_w / iw, target_h / ih)
    new_w = int(iw * scale)
    new_h = int(ih * scale)
    # center within target box
    cx = target_l + (target_w - new_w) / 2
    cy = target_t + (target_h - new_h) / 2
    return slide.shapes.add_picture(path, cx, cy, new_w, new_h)


def card_with_shadow(slide, l, t, w, h, color):
    """Card with a subtle darker 'shadow' offset."""
    shadow = rect(slide, l + Inches(0.03), t + Inches(0.03), w, h, RGBColor(0x08, 0x0C, 0x12))
    main = rect(slide, l, t, w, h, color)
    return main


# ── Slide 1 · Cover ────────────────────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0), Inches(2.6), SLIDE_W, color=ACCENT)
accent_line(s, Inches(0), Inches(5.2), SLIDE_W, color=ACCENT)
tb(s, Inches(1.5), Inches(2.9), Inches(10.3), Inches(1.2),
   'Claude Code 安装教程', size=54, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
tb(s, Inches(1.5), Inches(4.0), Inches(10.3), Inches(0.7),
   '从零开始 · 轻松上手 AI 编程助手', size=24, color=GRAY2, align=PP_ALIGN.CENTER)
tb(s, Inches(1.5), Inches(5.5), Inches(10.3), Inches(0.5),
   'Node.js  |  Claude Code CLI  |  Git  |  CC-Switch  |  DeepSeek  |  VS Code',
   size=15, color=GRAY3, align=PP_ALIGN.CENTER)
tb(s, Inches(1.5), Inches(6.3), Inches(10.3), Inches(0.4),
   '2026 / 06', size=13, color=GRAY3, align=PP_ALIGN.CENTER)


# ── Slide 2 · Overview ─────────────────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
tb(s, Inches(1.0), Inches(0.35), Inches(5), Inches(0.7),
   '安装步骤总览', size=34, color=WHITE, bold=True)
tb(s, Inches(1.0), Inches(0.95), Inches(5), Inches(0.4),
   '6 步完成 Claude Code 安装与配置', size=14, color=GRAY2)

steps = [
    ('01', 'Node.js', '运行环境，Claude Code 的基石'),
    ('02', 'Claude Code CLI', 'npm 全局安装命令行工具'),
    ('03', 'Git', '版本控制工具，必须组件'),
    ('04', '配置文件', '修改 claude.json 跳过登录'),
    ('05', 'CC-Switch + DeepSeek', '接入国产大模型，低成本高性价比'),
    ('06', 'VS Code 集成', '可视化编辑器，效率翻倍'),
]
for i, (num, title, desc) in enumerate(steps):
    y = Inches(1.8) + Inches(0.88) * i
    rect(s, Inches(1.0), y, Inches(7.5), Inches(0.7), BG_CARD)
    step_circle(s, Inches(1.15), y + Inches(0.08), num, 0.48)
    tb(s, Inches(1.85), y + Inches(0.05), Inches(2.5), Inches(0.32),
       title, size=17, color=WHITE, bold=True)
    tb(s, Inches(1.85), y + Inches(0.36), Inches(5.5), Inches(0.28),
       desc, size=12, color=GRAY2)

# Right panel
rect(s, Inches(9.3), Inches(1.8), Inches(3.5), Inches(5.0), BG_CARD)
tb(s, Inches(9.6), Inches(2.1), Inches(3.0), Inches(0.4),
   '📌 为什么用 CC-Switch?', size=16, color=GOLD, bold=True)
bullets(s, Inches(9.6), Inches(2.6), Inches(3.0), Inches(3.8), [
    'Claude 是国外产品',
    '需要科学上网',
    '账号经常被封',
    'API 价格昂贵',
    'CC-Switch 一键切换',
    '接入国产 DeepSeek',
    '无需翻墙，成本更低',
], size=12, color=GRAY2, bullet='•', spacing=Pt(4))


# ── Slide 3 · Node.js ──────────────────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
step_circle(s, Inches(0.9), Inches(0.35), '1', 0.5)
tb(s, Inches(1.7), Inches(0.35), Inches(6), Inches(0.65),
   '安装 Node.js —— 运行环境', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(6), Inches(0.35),
   'Claude Code 依赖 Node.js，必须先安装', size=13, color=GRAY2)

# Left: text content
left_x = Inches(0.9)
y0 = Inches(1.7)
tb(s, left_x, y0, Inches(5.5), Inches(0.4), '📥 下载安装', size=22, color=GOLD, bold=True)
bullets(s, left_x, y0 + Inches(0.55), Inches(5.8), Inches(0.8), [
    '访问官网：https://nodejs.org/zh-cn/download',
    '推荐选择 LTS（长期支持）版本，更加稳定',
], size=14, color=GRAY1)

code_block(s, left_x, y0 + Inches(1.5), Inches(5.8), Inches(0.5),
            'https://nodejs.org/zh-cn/download', size=13)

tb(s, left_x, y0 + Inches(2.3), Inches(5.5), Inches(0.4),
   '✅ 验证安装', size=22, color=ACCENT2, bold=True)
bullets(s, left_x, y0 + Inches(2.8), Inches(5.8), Inches(0.6), [
    'Win + R → 输入 cmd → 回车打开命令行',
], size=14, color=GRAY1)
code_block(s, left_x, y0 + Inches(3.35), Inches(4.0), Inches(0.45),
            'node -v', size=14)
tb(s, left_x, y0 + Inches(4.0), Inches(5.8), Inches(0.4),
   '出现版本号（如 v20.x.x）即表示安装成功', size=14, color=ACCENT2, bold=True)

# Right: image
img_x = Inches(7.3); img_w = Inches(5.5); img_h = Inches(5.2)
rect(s, img_x - Inches(0.1), Inches(1.6), img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image.png'), img_x, Inches(1.8), img_w, img_h)


# ── Slide 4 · Claude Code CLI ──────────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
step_circle(s, Inches(0.9), Inches(0.35), '2', 0.5)
tb(s, Inches(1.7), Inches(0.35), Inches(6), Inches(0.65),
   '安装 Claude Code CLI', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(6), Inches(0.35),
   '通过 npm 全局安装 Claude Code 命令行工具', size=13, color=GRAY2)

left_x = Inches(0.9)
y0 = Inches(1.7)
tb(s, left_x, y0, Inches(5.5), Inches(0.4), '⚠️ 重要提示', size=20, color=ACCENT, bold=True)
bullets(s, left_x, y0 + Inches(0.5), Inches(5.8), Inches(0.6), [
    '使用 Windows PowerShell（右键以管理员身份运行）',
    '或直接在 cmd 中执行（同样需要管理员权限）',
], size=13, color=GRAY1)

tb(s, left_x, y0 + Inches(1.3), Inches(5.5), Inches(0.4), '📦 安装命令', size=22, color=GOLD, bold=True)
code_block(s, left_x, y0 + Inches(1.8), Inches(6.0), Inches(0.5),
            'npm install -g @anthropic-ai/claude-code', size=14)
tb(s, left_x, y0 + Inches(2.5), Inches(5.8), Inches(0.5),
   '-g 表示全局安装，安装后在任何目录都可以直接使用 claude 命令',
   size=12, color=GRAY2)

tb(s, left_x, y0 + Inches(3.2), Inches(5.8), Inches(0.4), '💡 加速技巧', size=20, color=ACCENT3, bold=True)
rect(s, left_x, y0 + Inches(3.65), Inches(6.0), Inches(0.7), BG_CARD)
tb(s, left_x + Inches(0.2), y0 + Inches(3.75), Inches(5.6), Inches(0.55),
   '如安装速度慢，可切换到国内镜像源：\n'
   'npm config set registry https://registry.npmmirror.com',
   size=12, color=ACCENT2)

# Right: image
img_x = Inches(7.3); img_w = Inches(5.5); img_h = Inches(5.2)
rect(s, img_x - Inches(0.1), Inches(1.6), img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-1.png'), img_x, Inches(1.8), img_w, img_h)


# ── Slide 5 · Git ──────────────────────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
step_circle(s, Inches(0.9), Inches(0.35), '3', 0.5)
tb(s, Inches(1.7), Inches(0.35), Inches(6), Inches(0.65),
   '安装 Git —— 版本控制工具', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(6), Inches(0.35),
   'Claude Code 运行的必要组件，不安装会报错', size=13, color=GRAY2)

left_x = Inches(0.9)
y0 = Inches(1.7)
tb(s, left_x, y0, Inches(5.5), Inches(0.4), '📥 下载与安装', size=22, color=GOLD, bold=True)
bullets(s, left_x, y0 + Inches(0.55), Inches(6.2), Inches(1.2), [
    '访问 Git 官网：https://git-scm.com/downloads',
    '下载对应 Windows 版本的安装包',
    '双击运行安装程序，保持默认选项一路 Next',
    '安装完成，在 cmd 中输入 git --version 验证',
], size=14, color=GRAY1)

tb(s, left_x, y0 + Inches(2.5), Inches(6.0), Inches(0.4), '⚠️ 警示', size=20, color=ACCENT, bold=True)
bullets(s, left_x, y0 + Inches(3.0), Inches(6.2), Inches(0.8), [
    '如果未安装 Git，运行 Claude Code 会直接报错',
    'Git 是免费开源软件，安全可靠，请放心安装',
], size=14, color=GRAY1)

rect(s, left_x, y0 + Inches(3.9), Inches(6.0), Inches(0.55), BG_CARD)
tb(s, left_x + Inches(0.2), y0 + Inches(4.0), Inches(5.6), Inches(0.4),
   '✅ Git 安装成功后，Claude Code 已经可以运行！\n    但还需要修改配置跳过登录验证……',
   size=13, color=ACCENT2, bold=True)

# Right: image
img_x = Inches(7.5); img_w = Inches(5.3); img_h = Inches(5.2)
rect(s, img_x - Inches(0.1), Inches(1.6), img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-2.png'), img_x, Inches(1.8), img_w, img_h)


# ── Slide 6 · Config claude.json ──────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
step_circle(s, Inches(0.9), Inches(0.35), '4', 0.5)
tb(s, Inches(1.7), Inches(0.35), Inches(7), Inches(0.65),
   '配置 Claude Code —— 修改 claude.json', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(7), Inches(0.35),
   '绕过首次登录引导，为接入 DeepSeek 做准备', size=13, color=GRAY2)

# Two images side by side in the main area
# Left image
img_y = Inches(1.6)
img1_w = Inches(5.8); img1_h = Inches(3.0)
rect(s, Inches(0.8), img_y, img1_w + Inches(0.2), img1_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-3.png'), Inches(0.9), img_y + Inches(0.1), img1_w, img1_h)

# Right image
img2_x = Inches(7.2)
img2_w = Inches(5.8); img2_h = Inches(3.0)
rect(s, img2_x, img_y, img2_w + Inches(0.2), img2_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-4.png'), img2_x + Inches(0.1), img_y + Inches(0.1), img2_w, img2_h)

# Text explanation below
y_bot = Inches(5.0)
tb(s, Inches(0.9), y_bot, Inches(12), Inches(0.4),
   '🔧 关键配置', size=20, color=GOLD, bold=True)
bullets(s, Inches(0.9), y_bot + Inches(0.45), Inches(5.5), Inches(0.8), [
    '找到 claude.json 配置文件',
    '添加 "hasCompletedOnboarding": true',
], size=14, color=GRAY1)
code_block(s, Inches(0.9), y_bot + Inches(1.15), Inches(5.5), Inches(0.75),
            '{\n  "hasCompletedOnboarding": true\n}', size=14)
tb(s, Inches(7.2), y_bot + Inches(0.4), Inches(5.5), Inches(1.6),
   '设置完成后，Claude Code 可以运行，\n但会提示需要登录 Claude 账号。\n\n💡 Claude 是国外大模型，需要付\n    费买 token，账号经常被封。\n\n👉 解决方案：接入国产 DeepSeek！',
   size=13, color=GRAY1)


# ── Slide 7 · Login & CC-Switch intro ─────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
tb(s, Inches(1.7), Inches(0.35), Inches(7), Inches(0.65),
   '运行 Claude Code —— 登录提示', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(7), Inches(0.35),
   '首次运行会提示登录，我们需要绕过它', size=13, color=GRAY2)

# Two images
img_y = Inches(1.6)
img_w = Inches(6.0); img_h = Inches(2.8)
rect(s, Inches(0.7), img_y, img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-5.png'), Inches(0.8), img_y + Inches(0.1), img_w, img_h)

rect(s, Inches(7.2), img_y, img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-6.png'), Inches(7.3), img_y + Inches(0.1), img_w, img_h)

y_bot = Inches(4.8)
tb(s, Inches(0.9), y_bot, Inches(12), Inches(0.4),
   '🔑 解决方案：CC-Switch', size=22, color=GOLD, bold=True)
bullets(s, Inches(0.9), y_bot + Inches(0.5), Inches(12), Inches(1.8), [
    'CC-Switch 是一个开源工具，可一键切换 Claude Code 底层的大模型 API',
    '将 Claude Code 对接 DeepSeek（国产大模型），无需科学上网，成本大幅降低',
    'GitHub 地址：https://github.com/farion1231/cc-switch',
    '下载安装后，配置 DeepSeek API Key 即可使用',
], size=14, color=GRAY1)


# ── Slide 8 · CC-Switch download ──────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
step_circle(s, Inches(0.9), Inches(0.35), '5', 0.5)
tb(s, Inches(1.7), Inches(0.35), Inches(7), Inches(0.65),
   '安装 CC-Switch', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(7), Inches(0.35),
   '在 GitHub 下载 CC-Switch 工具', size=13, color=GRAY2)

# Two images stacked or side by side - these are tall web page screenshots
img_y = Inches(1.6)
img_w = Inches(6.0); img_h = Inches(5.0)
rect(s, Inches(0.7), img_y, img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-7.png'), Inches(0.8), img_y + Inches(0.1), img_w, img_h)

rect(s, Inches(7.2), img_y, img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-8.png'), Inches(7.3), img_y + Inches(0.1), img_w, img_h)


# ── Slide 9 · DeepSeek API Key ────────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
step_circle(s, Inches(0.9), Inches(0.35), '5', 0.5)
tb(s, Inches(1.7), Inches(0.35), Inches(7), Inches(0.65),
   '获取 DeepSeek API 密钥', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(7), Inches(0.35),
   '在 DeepSeek 开放平台创建 API Key', size=13, color=GRAY2)

# Two images
img_y = Inches(1.6)
img_w = Inches(6.0); img_h = Inches(3.0)
rect(s, Inches(0.7), img_y, img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-9.png'), Inches(0.8), img_y + Inches(0.1), img_w, img_h)

rect(s, Inches(7.2), img_y, img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-10.png'), Inches(7.3), img_y + Inches(0.1), img_w, img_h)

# Text below
y_bot = Inches(5.0)
tb(s, Inches(0.9), y_bot, Inches(12), Inches(0.4),
   '📋 操作步骤', size=20, color=GOLD, bold=True)
bullets(s, Inches(0.9), y_bot + Inches(0.45), Inches(12), Inches(1.5), [
    '打开 DeepSeek API 开放平台：platform.deepseek.com',
    '注册 / 登录账号 → 进入「API Keys」管理页面',
    '点击「创建 API Key」→ 复制生成的密钥（格式为 sk-xxxxx）',
    '⚠️ 密钥仅显示一次，请务必妥善保存！',
], size=14, color=GRAY1)


# ── Slide 10 · CC-Switch Config ───────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
tb(s, Inches(1.7), Inches(0.35), Inches(7), Inches(0.65),
   '配置 CC-Switch 接入 DeepSeek', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(7), Inches(0.35),
   '添加模型，粘贴密钥，完成最终配置', size=13, color=GRAY2)

# Two images
img_y = Inches(1.6)
img_w = Inches(5.8); img_h = Inches(3.0)
rect(s, Inches(0.8), img_y, img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-11.png'), Inches(0.9), img_y + Inches(0.1), img_w, img_h)

rect(s, Inches(7.2), img_y, img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-12.png'), Inches(7.3), img_y + Inches(0.1), img_w, img_h)

y_bot = Inches(5.0)
tb(s, Inches(0.9), y_bot, Inches(12), Inches(0.4),
   '🎉 配置完成', size=20, color=ACCENT2, bold=True)
bullets(s, Inches(0.9), y_bot + Inches(0.45), Inches(12), Inches(1.5), [
    '打开 CC-Switch → 添加模型 → 选择 DeepSeek → 粘贴 API Key',
    '保存配置后，重新进入 cmd 运行 claude',
    '不再提示登录，Claude Code 已成功对接 DeepSeek！',
], size=14, color=GRAY1)


# ── Slide 11 · DeepSeek Recharge ──────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
tb(s, Inches(1.7), Inches(0.35), Inches(7), Inches(0.65),
   'DeepSeek 账号充值与使用', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(7), Inches(0.35),
   '完成实名认证和充值后即可畅享 AI 编程', size=13, color=GRAY2)

# Three images in a row
img_y = Inches(1.6)
img_w = Inches(3.9); img_h = Inches(2.8)
positions = [Inches(0.6), Inches(4.7), Inches(8.8)]
img_files = ['image-13.png', 'image-14.png', 'image-15.png']
for pos, fname in zip(positions, img_files):
    rect(s, pos, img_y, img_w + Inches(0.15), img_h + Inches(0.2), BG_CARD)
    fit_image(s, os.path.join(IMG_DIR, fname), pos + Inches(0.07), img_y + Inches(0.1), img_w, img_h)

y_bot = Inches(4.8)
tb(s, Inches(0.9), y_bot, Inches(12), Inches(0.4),
   '💰 实名认证 & 充值', size=20, color=GOLD, bold=True)
bullets(s, Inches(0.9), y_bot + Inches(0.45), Inches(6.0), Inches(1.5), [
    'DeepSeek 需要实名认证后才能充值使用',
    '按照平台指引完成身份验证',
    '认证通过后选择合适的金额充值',
    '充值后即可无限制调用 API',
], size=13, color=GRAY1)

# Price comparison on right
rect(s, Inches(7.5), y_bot + Inches(0.4), Inches(5.3), Inches(2.2), BG_CARD)
tb(s, Inches(7.8), y_bot + Inches(0.5), Inches(4.5), Inches(0.35),
   '💡 DeepSeek vs Claude 价格参考', size=14, color=GOLD, bold=True)
tb(s, Inches(7.8), y_bot + Inches(0.95), Inches(4.5), Inches(0.7),
   'DeepSeek-V3\n  输入：¥1 / 百万 tokens\n  输出：¥2 / 百万 tokens\n\nClaude Sonnet 4\n  输入：$3 / 百万 tokens\n  输出：$15 / 百万 tokens',
   size=12, color=GRAY1)
tb(s, Inches(7.8), y_bot + Inches(2.3), Inches(4.5), Inches(0.25),
   '👉 DeepSeek 性价比远超 Claude，便宜数十倍！', size=12, color=ACCENT2, bold=True)


# ── Slide 12 · VS Code ─────────────────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0.7), Inches(0.55), Inches(0.07), Inches(0.45))
step_circle(s, Inches(0.9), Inches(0.35), '6', 0.5)
tb(s, Inches(1.7), Inches(0.35), Inches(7), Inches(0.65),
   'VS Code 集成 Claude Code', size=30, color=WHITE, bold=True)
tb(s, Inches(1.7), Inches(0.95), Inches(7), Inches(0.35),
   '在可视化编辑器中运行，告别纯命令行', size=13, color=GRAY2)

left_x = Inches(0.9)
y0 = Inches(1.7)
tb(s, left_x, y0, Inches(5.5), Inches(0.4),
   '🌟 为什么选择 VS Code？', size=20, color=GOLD, bold=True)
bullets(s, left_x, y0 + Inches(0.5), Inches(5.8), Inches(1.5), [
    '微软出品的免费编辑器，轻量且功能强大',
    '可视化操作界面，告别纯命令行',
    '可直观查看代码变更、对比差异',
    '丰富插件生态，无限扩展可能',
    '免费！无需付费即可使用全部功能',
], size=13, color=GRAY1)

tb(s, left_x, y0 + Inches(2.4), Inches(5.5), Inches(0.4),
   '🔌 安装 Claude Code 插件', size=20, color=ACCENT3, bold=True)
bullets(s, left_x, y0 + Inches(2.9), Inches(5.8), Inches(1.2), [
    '打开 VS Code → 点击左侧扩展图标（Ctrl+Shift+X）',
    '搜索 "Claude Code" → 找到 Anthropic 官方插件',
    '点击 Install → 等待安装完成',
    '重启 VS Code → 在终端（Ctrl+`）输入 claude',
], size=13, color=GRAY1)

rect(s, left_x, y0 + Inches(4.3), Inches(5.8), Inches(0.7), BG_CARD)
tb(s, left_x + Inches(0.2), y0 + Inches(4.4), Inches(5.4), Inches(0.5),
   '🎉 恭喜！现在你可以在 VS Code 中享受\n    可视化 AI 编程体验了！',
   size=14, color=ACCENT2, bold=True)

# Right: image
img_x = Inches(7.3); img_w = Inches(5.5); img_h = Inches(5.2)
rect(s, img_x - Inches(0.1), Inches(1.6), img_w + Inches(0.2), img_h + Inches(0.2), BG_CARD)
fit_image(s, os.path.join(IMG_DIR, 'image-16.png'), img_x, Inches(1.8), img_w, img_h)


# ── Slide 13 · Summary ─────────────────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6]); bg(s)
accent_line(s, Inches(0), Inches(1.8), SLIDE_W, color=ACCENT)

tb(s, Inches(1.5), Inches(0.6), Inches(10.3), Inches(0.8),
   '🎯 总结', size=44, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
tb(s, Inches(1.5), Inches(1.3), Inches(10.3), Inches(0.4),
   '安装完成后，在终端输入 claude 即可开启 AI 编程之旅', size=15, color=GRAY2, align=PP_ALIGN.CENTER)

cards = [
    ('Node.js', '运行环境', 'nodejs.org', ACCENT),
    ('Claude Code', 'AI 编程 CLI', 'npm install -g', ACCENT3),
    ('Git', '版本控制', 'git-scm.com', GOLD),
    ('CC-Switch', '模型切换', 'GitHub', ACCENT2),
    ('DeepSeek', '国产大模型', 'platform.deepseek', ACCENT),
    ('VS Code', '可视化编辑', '插件市场', ACCENT3),
]
card_w = Inches(1.9); card_h = Inches(3.2)
start_x = Inches(0.6); gap = Inches(0.15)
for i, (name, desc, link, clr) in enumerate(cards):
    cx = start_x + (card_w + gap) * i
    cy = Inches(2.4)
    rect(s, cx, cy, card_w, card_h, BG_CARD)
    accent_line(s, cx, cy, card_w, color=clr)
    tb(s, cx + Inches(0.1), cy + Inches(0.3), Inches(1.7), Inches(0.4),
       name, size=17, color=clr, bold=True, align=PP_ALIGN.CENTER)
    tb(s, cx + Inches(0.1), cy + Inches(1.0), Inches(1.7), Inches(0.4),
       desc, size=12, color=GRAY2, align=PP_ALIGN.CENTER)
    tb(s, cx + Inches(0.1), cy + Inches(1.8), Inches(1.7), Inches(0.4),
       link, size=10, color=GRAY3, align=PP_ALIGN.CENTER)

tb(s, Inches(1.5), Inches(6.0), Inches(10.3), Inches(0.8),
   'Claude Code + DeepSeek = 高效 · 低成本 · 无限制的 AI 编程体验',
   size=18, color=GRAY1, align=PP_ALIGN.CENTER)


# ── Save ───────────────────────────────────────────────────────
prs.save(OUTPUT)
print(f'PPT saved to: {OUTPUT}')
print(f'Total slides: {len(prs.slides)}')
