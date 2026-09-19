# -*- coding: utf-8 -*-
"""类型级本体构建进度汇报 —— PPTX 生成（与 index.html 同内容）
运行： python3 build_pptx.py
输出： 本体构建进度汇报.pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BG   = RGBColor(0x0a, 0x0e, 0x1a)
CARD = RGBColor(0x16, 0x1f, 0x2e)
TXT  = RGBColor(0xf1, 0xf5, 0xf9)
BODY = RGBColor(0xcb, 0xd5, 0xe1)
DIM  = RGBColor(0x94, 0xa3, 0xb8)
AMB  = RGBColor(0xfb, 0xbf, 0x24)
AMB2 = RGBColor(0xf5, 0x9e, 0x0b)
SKY  = RGBColor(0x38, 0xbd, 0xf8)
GRN  = RGBColor(0x34, 0xd3, 0x99)
RED  = RGBColor(0xf8, 0x71, 0x71)
PUR  = RGBColor(0xa7, 0x8b, 0xfa)

W, H = 13.333, 7.5
ML, CW = 0.5, 12.333

prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
BLANK = prs.slide_layouts[6]


def new_slide():
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    return s


def tb(s, l, t, w, h, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.paragraphs[0].alignment = align
    return tf


def put(tf, text, size=11, color=BODY, bold=False, align=None, first=False, space_before=0):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    if align is not None:
        p.alignment = align
    if space_before:
        p.space_before = Pt(space_before)
    return p


def title(s, text, lead=None, top=0.42):
    tf = tb(s, ML, top, CW, 0.72, align=PP_ALIGN.CENTER)
    put(tf, text, size=25, color=AMB, bold=True, align=PP_ALIGN.CENTER, first=True)
    y = top + 0.78
    if lead:
        tf2 = tb(s, ML, y, CW, 0.34, align=PP_ALIGN.CENTER)
        put(tf2, lead, size=13, color=BODY, align=PP_ALIGN.CENTER, first=True)
        y += 0.5
    return y


def box(s, l, t, w, h, fill=CARD, border=None, bar=None):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if border:
        sh.line.color.rgb = border
        sh.line.width = Pt(1)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    sh.adjustments[0] = 0.08
    if bar:
        b = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t + 0.06), Inches(0.055), Inches(h - 0.12))
        b.fill.solid()
        b.fill.fore_color.rgb = bar
        b.line.fill.background()
        b.shadow.inherit = False
    return sh


def size_for(text, base=10.5, per=95):
    n = len(text)
    if n > 2 * per:
        return base - 1.5
    if n > per:
        return base - 0.7
    return base


def cards(s, items, y, ncols=None, h=None, color=AMB):
    """items: [(label, text)] -> 返回该块总高度"""
    n = len(items)
    ncols = ncols or n
    gap = 0.16
    cw = (CW - gap * (ncols - 1)) / ncols
    hs = []
    for lab, txt in items:
        if h:
            hs.append(h)
        else:
            lines = max(2, int(len(txt) / 42) + 1)
            hs.append(max(0.62, 0.50 + lines * 0.225))
    rows = (n + ncols - 1) // ncols
    row_h, ypos, cy = [], [], y
    for r in range(rows):
        seg = hs[r * ncols:(r + 1) * ncols]
        rh = max(seg) if seg else 0.6
        row_h.append(rh)
        ypos.append(cy)
        cy += rh + gap
    for i, (lab, txt) in enumerate(items):
        r, c = divmod(i, ncols)
        l = ML + c * (cw + gap)
        t = ypos[r]
        ch2 = row_h[r]
        box(s, l, t, cw, ch2, fill=CARD)
        tf = tb(s, l + 0.16, t + 0.09, cw - 0.32, ch2 - 0.18)
        put(tf, lab, size=11.5, color=color, bold=True, first=True)
        put(tf, txt, size=size_for(txt), color=BODY, space_before=2)
    return cy - gap - y


def block(s, items, y, h, color=AMB):
    """单块（如 q / 说明块）：一条文字"""
    box(s, ML, y, CW, h, fill=CARD, border=color, bar=color)
    tf = tb(s, ML + 0.2, y + 0.1, CW - 0.4, h - 0.2)
    first = True
    for t in items:
        put(tf, t[0], size=t[1], color=t[2], bold=t[3], first=first)
        first = False
    return h


def quote(s, text, y, h=0.80):
    box(s, ML, y, CW, h, fill=RGBColor(0x1a, 0x1d, 0x18), border=AMB2, bar=AMB)
    tf = tb(s, ML + 0.2, y + 0.08, CW - 0.4, h - 0.16, anchor=MSO_ANCHOR.MIDDLE)
    put(tf, text, size=size_for(text, 10.5), color=BODY, first=True)
    return h


def nums(s, items, y, h=0.95):
    n = len(items)
    gap = 0.16
    cw = (CW - gap * (n - 1)) / n
    for i, (num, sub, col) in enumerate(items):
        l = ML + i * (cw + gap)
        box(s, l, y, cw, h, fill=CARD)
        tf = tb(s, l + 0.1, y + 0.1, cw - 0.2, h - 0.2, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        put(tf, num, size=22, color=col, bold=True, align=PP_ALIGN.CENTER, first=True)
        put(tf, sub, size=9.5, color=DIM, align=PP_ALIGN.CENTER, space_before=3)
    return h


def flow(s, items, y, h=0.62):
    n = len(items)
    arrow = 0.3
    cw = (CW - arrow * (n - 1)) / n
    cols = [GRN, SKY, PUR, AMB]
    for i, txt in enumerate(items):
        l = ML + i * (cw + arrow)
        box(s, l, y, cw, h, fill=CARD, border=cols[i % 4])
        tf = tb(s, l + 0.06, y + 0.06, cw - 0.12, h - 0.12, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        put(tf, txt, size=10, color=cols[i % 4], bold=True, align=PP_ALIGN.CENTER, first=True)
        if i < n - 1:
            atf = tb(s, l + cw, y + 0.06, arrow, h - 0.12, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            put(atf, '\u2192', size=13, color=DIM, align=PP_ALIGN.CENTER, first=True)
    return h


def table(s, header, rows, y, col_w=None, fs=10, hrow=0.34):
    nr, nc = len(rows) + 1, len(header)
    h = hrow + nr * 0.30
    shp = s.shapes.add_table(nr, nc, Inches(ML), Inches(y), Inches(CW), Inches(h))
    tbl = shp.table
    tbl.first_row = True
    if col_w:
        tot = sum(col_w)
        for i, cwv in enumerate(col_w):
            tbl.columns[i].width = Inches(CW * cwv / tot)
    def fill_cell(cell, text, size, color, bold=False, align=PP_ALIGN.CENTER):
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x1b, 0x24, 0x35)
        cell.margin_left = cell.margin_right = Inches(0.05)
        cell.margin_top = cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.alignment = align
    for j, htxt in enumerate(header):
        fill_cell(tbl.cell(0, j), htxt, fs + 0.5, AMB, bold=True)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            col = BODY
            bold = (j == 1)
            txt = val
            if isinstance(val, tuple):
                txt, col = val
            fill_cell(tbl.cell(i, j), txt, fs, col, bold=bold)
    return h


# ============================ 内容 ============================
# 1 封面
s = new_slide()
tf = tb(s, 0.8, 2.35, W - 1.6, 1.2, align=PP_ALIGN.CENTER)
put(tf, '类型级本体构建', size=44, color=AMB, bold=True, align=PP_ALIGN.CENTER, first=True)
tf = tb(s, 1.4, 3.5, W - 2.8, 0.5, align=PP_ALIGN.CENTER)
put(tf, '从大规模文本中学习 Schema：LLM 抽取 → 统计收敛 → 增量拼接', size=16, color=TXT, align=PP_ALIGN.CENTER, first=True)
tf = tb(s, 1.8, 4.05, W - 3.6, 0.4, align=PP_ALIGN.CENTER)
put(tf, '阶段进度汇报 · 三域实证（法律 / 欧盟法 / 计算机科学）', size=12.5, color=DIM, align=PP_ALIGN.CENTER, first=True)
box(s, 3.7, 4.7, 5.93, 0.5, fill=RGBColor(0x10, 0x1c, 0x28), border=RGBColor(0x1e, 0x4a, 0x5e))
tf = tb(s, 3.8, 4.78, 5.73, 0.34, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
put(tf, '类型级（TBox）· 统计收敛作归纳偏置 · 单数据集域内自适应', size=11.5, color=SKY, align=PP_ALIGN.CENTER, first=True)

# 2 研究问题
s = new_slide()
y = title(s, '这项研究在解决什么问题', '把"读文本 → 写本体"这件本该由人做的 Schema 设计工作，交给机器做，并且做得可验证')
cards(s, [('目标', '从大规模领域文本中自动学习类型级本体（TBox：类型 + 上下位关系），而不是只做逐实例的 ABox 填充。本体是一套 Schema——它可以反过来约束、校验、复用在下游的图构建与推理上。'),
          ('为什么难', '① 抽取本身有噪声，且每次运行结果都在变；② 同一概念在文本里有多种叫法（异词同义），越高频越碎；③ 层级关系在文本里是隐含的，不能靠名字相似度猜；④ 缺少该规模下的人工标注真值，评价口径本身要设计。')],
      y, ncols=2, color=PUR)
quote(s, '核心思路：单个文档里"谁是上位类型"不可靠，但跨文档的统计频次可靠 —— 把类型级统计分布当作归纳偏置：频次决定"谁值得进本体、谁和谁能合并、谁是谁的上位"，LLM 只负责它擅长的语义判断。',
      y + 1.9, h=0.85)

# 3 方法主线
s = new_slide()
y = title(s, '方法主线：三阶段流水线', '每一阶段都有明确的机器可检验产物')
y += flow(s, ['① LLM 抽取（类型 / 关系 / 三层）', '② 类型级统计收敛（频次即归纳偏置）', '③ 增量拼接（ConceptCard + 局部操作）', '本体树（类型 + 上下位）'], y) + 0.2
cards(s, [('设计 A · 三层类型作"坐标"', '抽取同时产出 direct（最细，作叶子）/ parent（中层，作聚合与匹配层）/ grandparent（粗类，作分区坐标）。实测：L1 不可匹配（巨分量仅 61–92%），L2 恒 100% 连通 → 聚合层必须放在 L2。'),
          ('设计 B · star 结构（关系通配符化）', '三元组把关系换成通配符，碎片变成类型对 (A) —*→ (B)，权重由抽取统计量决定。让"关系层噪声"不再污染结构层，同时保留统计信号。')],
      y, ncols=2)
quote(s, '老师口径：不做全泛化 OL，限定单个数据集 —— 因此顶层设计走"域内自适应"，跨域只做可比的锚点层。', y + 1.5, h=0.6)

# 4 数据与规模
s = new_slide()
y = title(s, '实验规模：三个域、两套口径', '三域语料规模从数千篇到近万篇，覆盖法律、多语法律、学术论文三类语体')
y += table(s, ['域', '语料规模（文档 / 三元组）', '可用外部本体（TBox）', '本体树节点'],
           [['CFR（美国联邦法规）', '6,190 / 97,118', 'LKIF-Core（203 节点）', '3,231'],
            ['MultiEURLEX（欧盟法）', '8,262 / 217,645', 'EuroVoc（7,402 标签）', '4,381'],
            ['arXiv（计算机科学）', '2,173 / 175,573', 'ACM CCS（13 类）', '495']],
           y, col_w=[1.5, 1.5, 1.4, 1.0]) + 0.18
cards(s, [('有 TBox 真值', '4 套标准本体可用于类型级对齐（LKIF / EuroVoc / US Law / ACM CCS）。'),
          ('无 ABox 标注', '没有该语料规模下的人工实例标注 → 评价主要靠结构指标 + 类型级对齐 + 人工抽检。'),
          ('口径说明', 'ACM CCS 是主题分类而非类型本体，不可直接作为质量判据 —— 已在报告中显式标注。')], y, ncols=3)

# 5 抽取层
s = new_slide()
y = title(s, '阶段进展 ①：抽取层已定稿', '经过版本迭代，三层抽取 prompt 收敛到可用版本，并用两套模型交叉验证')
y += cards(s, [('类型层 · 定稿 v8c', '27B 本地模型已达更大云端模型的收敛水平（类型数 191 ≈ 194）。内容准确率 99.6%。'),
               ('关系层 · 定稿 v9.1', '句子锚定率 72% 完整 / 25% 部分 / 3% 无锚点；三元组约 20.0 条/篇。短板：schema 缺 value / qualifier 槽位。'),
               ('三层类型 · 定稿 v13.3', '22 类粗类 + concept 兜底，兜底占比 4.1–6.4%。direct / parent / grandparent 一次产出。')],
           y, ncols=3) + 0.18
cards(s, [('稳定性（关键结论）', '重复抽取的表达层波动大（轮间三元组 Jaccard 0.59，逐篇 0.11–1.00），但概念层 1–2 轮就饱和——重复 5–10 次不再产生新类型。→ 报告必须取 3 轮以上的收敛值，单轮结果不能当结论。'),
          ('抽取质量', '逐条人工/自动判定：明显错误率 < 1%；语义段按频次加权约 2–4%。两个后端抽出的三元组集合差异很大（Jaccard 0–0.45），但各自抽出的都基本正确 —— 属于互补而非冲突。')],
      y, ncols=2, color=GRN)

# 6 star 结构
s = new_slide()
y = title(s, '阶段进展 ②：star 结构的效果', '关系通配符化把"关系层噪声"从结构层剥离，同时显著改善长尾与复现性')
y += nums(s, [('3.6–7.3×', '碎片压缩倍数（同义/近义碎片被归并）', GRN),
              ('60–64% ↓ 41–44%', '长尾碎片占比（长尾显著收敛）', SKY),
              ('+25–40%', '劈半复现性（数据切一半，结论复现率）', AMB),
              ('不变', '节点数与覆盖率（不是靠删数据换来的）', PUR)], y, h=1.0) + 0.18
y += cards(s, [('为什么复现性会提升', '原口径下每个关系的命名变体都会生成一个新的碎片；star 化后同一类型对的多个关系共享一个碎片，统计量随之合并，低频噪声不再各自成片。'),
               ('为什么这很重要', '评审最常质疑的是"LLM 抽取不稳定、结论不可复现"。劈半复现性正是对这个质疑的直接回应指标 —— 而 star 化让它提升了 25–40%。')],
           y, ncols=2) + 0.16
quote(s, '方法学结论：结构层（类型对 + 统计权重）比关系层（具体关系名）更稳定 —— 这也解释了为什么后续建树完全基于类型层。', y, h=0.62)

# 7 v4 建树
s = new_slide()
y = title(s, '阶段进展 ③：从"碎片森林"到单连通本体树', '同一份抽取结果，只改建层级机制，结构指标全线翻转（CFR，同口径质检）')
y += table(s, ['结构指标', '改机制前', '改机制后', '说明'],
           [['有父节点比例', ('25.60%', RED), ('99.97%', GRN), '此前 74–77% 节点是孤立碎片'],
            ['连通分量 / 单节点树', ('1,891 / 1,683', RED), ('1 / 0', GRN), '从"碎片森林"变为单棵连通树'],
            ['最大连通树占比', ('5.55%', RED), ('100%', GRN), '全部节点接入同一棵树'],
            ['多父冲突（DAG 冲突）', ('137', RED), ('0', GRN), '每个节点唯一父'],
            ['上下位边数 / 平均深度', '869 / 0.35', ('3,230 / 2.68', GRN), '深度封顶 3 层'],
            ['对齐标准本体（严格口径）', '84（3.09%）', ('123（3.81%）', GRN), '绝对命中数提升 46%'],
            ['标准本体侧覆盖率', '18.05%', ('23.41%', GRN), '被覆盖的标准本体节点占比']],
           y, col_w=[2.0, 1.2, 1.2, 2.2]) + 0.16
quote(s, '关键机制发现：抽取阶段早就输出了上位类型字段（覆盖率 100%），但旧流水线一处都没引用，层级靠"向量召回 + LLM 重猜"重建 —— 这是覆盖率长期被钉在 22–33% 的真正原因。', y, h=0.72)

# 8 三域一致
s = new_slide()
y = title(s, '三域一致，且"变好"不是靠删数据')
y += table(s, ['域', '节点', '上下位边', '有父比例', '连通分量', '真环', '粗类数'],
           [['CFR', '3,231', '3,230', '99.97%', '1', '0', '21'],
            ['MultiEURLEX', '4,381', '4,380', '99.98%', '1', '0', '20'],
            ['arXiv', '495', '494', '99.80%', '1', '0', '20']],
           y, col_w=[1.6, 1, 1, 1, 1, 0.8, 1]) + 0.16
y += cards(s, [('完整链路抽检 14 条', '12 条明确正确、2 条宽松可接受、0 条错误。例：statute ⊑ legislation ⊑ document ⊑ root。'),
               ('粗类判定抽检 25 条', '10 条明确正确、5 条宽松可接受、0 条明显错误。'),
               ('节点是否被砍', '没有：2,719 → 3,231（增加的是中间层与粗类），语义边完整保留。')],
           y, ncols=3, color=GRN) + 0.16
quote(s, '诚实记录（两条弯路）：① 用"名字向量聚类"归并顶层 → 472 个不相关的根被塞进同一个类（抽象名词的名字向量区分度太低）；② 允许上位类型继续向上串联 → 出现 job_role ⊑ research_activity 这类链式漂移。结论：第一跳可信，多跳不可信，改为"压平到粗层"。', y, h=0.88)

# 9 规模实验
s = new_slide()
y = title(s, '重要发现：加数据量不能自动改善质量', '同流水线、同配置，只换数据量（CFR 1× / 2.5× / 10×）')
y += table(s, ['指标', '1×（619 篇）', '2.5×', '10×（6,190 篇）', '趋势'],
           [['有父节点比例', '24.8%', '33.1%', '25.6%', ('非单调，无系统性改善', RED)],
            ['对齐标准本体（严格）', '4.34%', '3.90%', '3.09%', ('单调下降', RED)],
            ['同义簇内节点占比', '25.0%', '28.0%', '33.0%', ('单调恶化', RED)],
            ['LLM 判定采纳率', '4.9%', '6.0%', '9.9%', ('唯一真实收益', GRN)],
            ['每类型摊到的候选数', '14.78', '14.75', '6.90', ('10× 档腰斩', RED)]],
           y, col_w=[1.9, 1.1, 0.8, 1.3, 1.6]) + 0.16
y += cards(s, [('① 分母涨、分子同比例涨', '无父节点中 61% 是"有候选但被否"，真"无候选"仅 7–15%；各频次桶的有父率几乎相同（19–29%）。'),
               ('② 候选召回是固定预算', '候选来自固定 Top-K 向量召回，数据量越大每类型摊到的候选反而越少 → 覆盖率上限被钉死。'),
               ('③ 合并门槛与数据量无关', '"实例重叠 ≥20%"的比值门槛拦掉了 CFR 91.8% / ME 74.4% 的真同义对 —— 模型根本没机会判断。')],
           y, ncols=3, color=RED) + 0.14
quote(s, '归档警示：在 6–9.4× 语料上顺势放宽阈值（覆盖率 0.8）→ 本体塌缩到 14 个节点。数据量上去后放宽阈值是危险操作。', y, h=0.6)

# 10 顶层设计
s = new_slide()
y = title(s, '顶层设计：域内自适应 + 跨域锚点层', '发现固定类表跨域会错配后，顶层改为按域内分布归纳（依据领域专家评审意见）')
y += cards(s, [('改之前：一套固定类表跨域套用', '把法律域的 22 类固定表套到计算机科学域，产生明显错配：theoretical_guarantee（理论保证）被判成"权利 right"、filtering_rule 被判成"文书"。根因：固定类表在多义抽象名词上必然误吸。'),
               ('改之后：域内自适应顶层类', '取覆盖 85% 权重的第二层节点 → 让模型提候选顶层类（含定义 / 边界 / 正例 / 反例）→ 约束剪枝 → 对全部第二层节点重新归类。结果：13 个域内自适应类，法律类完全消失；兜底权重 4.3%（<5%）、单类最大 21.2%、覆盖 100%。')],
           y, ncols=2) + 0.16
y += cards(s, [('设计选择 1', '顶层类域内自适应，允许不同域有不同顶层。'),
               ('设计选择 2', '类表开放可增长：任何作匹配层的层，类数必须随语料增长。'),
               ('设计选择 3', '跨域比较只在锚点层做 —— 该层不参与域内建树。')], y, ncols=3, h=0.8) + 0.14
quote(s, '为什么必须"随语料增长"：固定 22 类在全量下覆盖率达到 100.0% 却只覆盖 48.5–60.8% 的权重 —— 一个全连通的分类层等于没有分类（退化为完全图）。', y, h=0.66)

# 11 CAL
s = new_slide()
y = title(s, '跨域可比性：锚点层（CAL）的实证结果', '11 个元类，用同一把尺子量三个域 —— 论文里可直接使用的跨域陈述')
y += table(s, ['锚点元类', 'CFR（美国法规）', 'MultiEURLEX（欧盟法）', 'arXiv（计算机科学）'],
           [['文书 / 信息载体', ('35.6%', TXT), ('47.9%', TXT), ('1.6%', RED)],
            ['方法 / 模型', ('2.0%', RED), ('2.3%', RED), ('36.3%', TXT)],
            ['主体 / 行为者（Agent）', ('26.3%', TXT), ('31.9%', TXT), ('24.9%', TXT)],
            ['数据 / 资源', '7.5%', '2.3%', ('24.5%', TXT)],
            ['规范性地位', '6.1%', '2.4%', ('0%', RED)],
            ['兜底（Concept）', '9.9%', '2.9%', '4.3%']],
           y, col_w=[1.7, 1.3, 1.5, 1.5]) + 0.18
cards(s, [('两个法律域', '由文书 + 规范性地位主导，结构相似（35.6/47.9% 与 6.1/2.4%）。'),
          ('计算机科学域', '由方法模型 + 数据资源主导，文书类几乎为 0 —— 域特征被锚点层干净地捕捉到。'),
          ('唯一跨域稳定元类', '主体 / 行为者（Agent）：三域均为 25–32%。这是三域唯一真正稳定的锚点，也是一个可被独立验证的实证结论。')],
      y, ncols=3, color=GRN)

# 12 异词同义合并
s = new_slide()
y = title(s, '当前 P0：异词同义合并（跨父）', '老师指定的优先方向 —— 也是质量报告点名的最大短板（同一概念 5 种叫法并存）')
y += cards(s, [('旧口径：三条证据里最弱的那条被当成硬门槛', '合并候选要求"实例重叠 ≥ 20%"。实测该门槛把 73% 的规则通过候选挡在模型判断之前（例：statutory_exemption 与 statutory_exclusion 名称相似度 0.933、邻域 0.888，但实例重叠为 0）。'),
               ('新口径：三证据 + 模型仲裁 + 上下位护栏', '① 名称语义 ≥0.90；② 邻域频次向量余弦（替代稀疏 Jaccard）；③ 实例重叠降级为支持证据。三条一起送模型判定，另加护栏：名称严格真包含的关系一律判为上下位、禁止合并。')],
           y, ncols=2, color=RED) + 0.18
y += nums(s, [('2,324', '类型总数', SKY), ('426 → 107', '名称候选 → 规则通过', SKY),
              ('73%', '其中实例重叠 = 0（旧门槛全部拦掉）', RED), ('12 对', '最终合并（前后指标无回退）', GRN)], y, h=1.0) + 0.16
quote(s, '本轮最重要的发现：两套模型的判定一致率 95%（102/107），且严格程度相当 —— 瓶颈已从"过滤器"转移到"判定框架"，换更大的模型没有收益。两套模型都拒掉了强证据对（如 structural_requirement + structural_design_requirement，邻域相似度高达 0.949）：它们把"限定词不同的变体"判成了不同概念。下一轮的杠杆在提示词框架，不在模型规模。', y, h=1.0)

# 13 评价口径
s = new_slide()
y = title(s, '评价口径：正在收紧（论文主表的基础）', '目前所有指标都是"分母 = 全部预测节点 / 全部节点"，不使用只算成功映射的虚高口径')
y += cards(s, [('核心指标：Verified Coverage（可验证覆盖率）', '用人工抽检验证过的区间来定义覆盖率，而不是把机器相似度高的部分当成全部正确。实测教训：相似度 0.85 以上的段人工验证 15/15 可接受，但 0.70–0.85 段只有约 28% 正确、32% 是错的 —— 该段只能作为"候选"，不能进结论。'),
               ('要一并报告的指标', '① 精确率 / 召回率 / F1（分母含全部预测）；② 四级匹配：Exact / Subsumption / Partial / Wrong；③ 标注一致性 κ（人工判定的一致性）；④ 保留自动指标时必须给出 PR 曲线，不能只报单点。')],
           y, ncols=2, color=GRN) + 0.18
cards(s, [('已排除的口径风险', '曾出现"只算成功映射节点"的精确率虚高至 0.71，真实值仅 0.10 —— 分母口径已改为全部预测节点。'),
          ('已修正的脚本误判', '曾把 DAG 多父节点被重复遍历判为"存在上下位环" —— 用无环检测复核后：三域真环数均为 0。'),
          ('可比的跨域口径', '域内顶层类允许不同，跨域比较只在锚点层做 → 保证跨域结论与域内细节不互相污染。')],
      y, ncols=3, color=GRN)

# 14 下一步
s = new_slide()
y = title(s, '下一步的工作与可能的学术方向')
y += table(s, ['#', '工作项', '现状', '核心问题'],
           [['1', '异词同义合并（跨父）', ('进行中', GRN), '改判定框架 + 强证据带直并，推广到另两个域'],
            ['2', '评价口径落地为论文主表', '进行中', 'Verified Coverage + PR 曲线 + 四级匹配 + κ'],
            ['3', '顶层设计的消融实验', '待开展', '固定类表 / 域内固定不增长 / 无锚点 / 有锚点 / 无顶层'],
            ['4', '增量拼接（阶段③）的规模化验证', '待开展', '流式追加数据时，本体的稳定性与漂移控制'],
            ['5', '三层抽取与 star 化接入主树', '待接入', '把已定稿机制统一到同一条主线']],
           y, col_w=[0.4, 3.0, 1.0, 4.2]) + 0.16
y += quote(s, '可能的投稿方向：方法论主线（统计收敛 + 域内自适应顶层 + 锚点式跨域）→ KDD / ACL 一级；本体工程与语义网方向 → ISWC / VLDB。具体定位待评价口径定稿后再定。', y, h=0.62) + 0.14
cards(s, [('可讲的故事 1', '单个文档里的上位关系不可靠，跨文档统计频次才是可靠信号 —— 统计收敛作归纳偏置。'),
          ('可讲的故事 2', '"加数据量不改善质量"的机制诊断：固定预算 + 比值门槛 是三个可复现的失败模式。'),
          ('可讲的故事 3', '域内自适应顶层 + 跨域锚点层：让 LLM 本体学习在单数据集上闭环，同时保持跨域可比。')],
      y, ncols=3)

# 15 小结
s = new_slide()
y = title(s, '一页小结')
y += cards(s, [('已经拿到的', '① 三域本体树全部为单连通、无环、有父率 >99.8% 的规整三层结构（此前 74–77% 是碎片）；② 抽取三层 prompt 定稿，明显错误率 <1%；③ star 化让碎片压缩 3.6–7.3×、劈半复现性 +25–40%；④ 跨域锚点层给出三域可比的实证分布（Agent 是唯一稳定元类）。'),
               ('正在攻的', '① 异词同义合并（老师指定的优先级）：门槛问题已证实并被绕过，瓶颈已定位到判定框架；② 评价口径：把"可验证覆盖率 + PR 曲线 + 四级匹配"做成论文主表；③ 顶层设计的五组消融与三层抽取接入主树。')],
           y, ncols=2, color=GRN) + 0.2
quote(s, '一句话定位：把"学本体"还原成"学 Schema" —— 用类型级统计收敛把 LLM 的语义判断锚定在可复现的频次证据上，并在单数据集内闭环、在锚点层上跨域可比。', y, h=0.75)

prs.save('本体构建进度汇报.pptx')
print('slides:', len(prs.slides.__iter__.__self__._sldIdLst))
print('saved: 本体构建进度汇报.pptx')

# 几何检查：底部不得超出 7.3"
bad = []
for i, s in enumerate(prs.slides, 1):
    for sh in s.shapes:
        if sh.top is None or sh.height is None:
            continue
        b = (sh.top + sh.height) / 914400
        if b > 7.34:
            bad.append((i, sh.shape_type, round(b, 2)))
print('overflow shapes:', bad if bad else 'none')
