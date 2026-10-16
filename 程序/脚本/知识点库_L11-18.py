# -*- coding: utf-8 -*-
r"""NCE1 Lesson 11-18 知识点库（四年级）。

全局题号：一 1-5 ｜ 二 6-25 ｜ 三 26-30 ｜ 四 31-35 ｜ 五 36-40
每条知识点含：标题、要记的内容（孩子看得懂的话）、对应全局题号。
"""

# ---------- 知识点 ----------
# key -> (标题, [要记的条目], [全局题号])
KP = {    # 一 名词复数
    "复数规则": (
        "名词变复数：一般直接加 s",
        ["pen → pens", "boy → boys", "book → books"],
        [1, 3],
    ),
    "复数es": (
        "名词变复数：以 s / x / ch / sh 结尾要加 es",
        ["dress → dresses", "bus → buses", "box → boxes", "watch → watches"],
        [2],
    ),
    "复数不规则": (
        "名词变复数：两个特殊的，必须背下来",
        ["woman → women（女人）", "man → men（男人）", "sheep → sheep（绵羊，单复数一样）",
         "child → children（孩子）"],
        [4, 5, 27],
    ),
    # 二 单项选择
    "be动词": (
        "be 动词怎么选：看前面的主语",
        ["I → am", "You → are", "He / She / It / 单数名词 → is",
         "We / You / They / 两个以上 → are",
         "例：My mother is a doctor.　Tom and Lucy are at school."],
        [11, 12, 13, 14, 15, 16],   # 全局 11-16（卷面 6-11）
    ),
    "物主代词": (
        "物主代词：表示\"谁的\"",
        ["my 我的　your 你的　his 他的　her 她的　its 它的",
         "It's = It is（它是）　Its = 它的　—— 这两个千万别混"],
        [7, 16, 29, 34],            # 全局 7（卷面2）、16（卷面11）、29（选词4）、34（改写4）
    ),
    "名词所有格": (
        "名词所有格：表示\"谁的\"，人名后面加 's",
        ["Kate → Kate's（凯特的）", "Tim → Tim's（蒂姆的）",
         "Whose is this shirt? ＝ Whose shirt is this?（这是谁的衬衫？）"],
        [8, 33],
    ),
    "冠词": (
        "冠词 a / an 怎么用",
        ["辅音音开头用 a：a teacher, a nurse, a car",
         "元音音开头用 an：an English girl, an apple, an old man"],
        [25],
    ),
    "疑问词": (
        "四个常用疑问词，问什么用什么",
        ["Whose 谁的　→ Whose is this book?",
         "What make 什么牌子　→ What make is this car?",
         "What nationality 哪国人　→ What nationality are you?",
         "What ... job 做什么工作　→ What is his job?"],
        [8, 21, 22, 23, 24, 33, 34, 35],
    ),
    "交际用语": (
        "见面打招呼、递东西，这几句要背熟",
        ["Nice to meet you. → Nice to meet you, too.",
         "How are you? → I'm fine, thank you.",
         "Give me some oranges, please. → Here you are.",
         "Are you a teacher? → Yes, I am. / No, I'm not.",
         "Is this your dress? → Yes, it is. / No, it isn't.",
         "Is Tom Japanese? → Yes, he is. / No, he isn't."],
        [6, 10, 13, 17, 18, 19, 20],
    ),
    "look at": (
        "“看……”的固定说法",
        ["look at the picture（看这张图）",
         "look 后面要跟 at 才能接东西；see 是“看见”，不加 at"],
        [9],
    ),
    # 三 选词填空
    "选词Chinese": (
        "国籍词",
        ["Chinese 中国人／中文（单复数一样，不加 s）",
         "问国籍：What nationality are you? — I am Chinese."],
        [26, 37],
    ),
    "选词代词": (
        "人称代词：主格和所有格别混",
        ["he 他（做主语）　it 它（做主语）",
         "your 你的（后面要跟名词）　you 你（做主语）",
         "例：Is this your umbrella?　It is a German car."],
        [28, 29, 30],
    ),
    # 四 改写句子
    "改写复数句": (
        "单数句子变复数：三个地方一起变",
        ["This → These　　is → are　　my → our",
         "例：This is my coat. → These are my coats.",
         "反过来：These are my coats. → This is my coat."],
        [31, 32],
    ),
    "改写提问": (
        "对划线部分提问（改写句子最后一类）",
        ["问“谁的”：Whose is this shirt?",
         "问“什么牌子”：What make is it?",
         "问“做什么工作”：What is his job? / What does he do?"],
        [33, 34, 35],
    ),
    # 五 阅读判断
    "职业词汇": (
        "家庭成员的职业，这几个词要认得住",
        ["policeman 警察　nurse 护士　housewife 家庭主妇",
         "keyboard operator 打字员　teacher 老师　doctor 医生",
         "family 家庭　grandma 奶奶　brother 兄弟"],
        [38, 39, 40],
    ),
}

# 全局题号 -> 知识点 key（一个题号可归入多个）
Q2KP = {}
for _k, (_t, _items, _qs) in KP.items():
    for _q in _qs:
        Q2KP.setdefault(_q, []).append(_k)

# 大题 -> 全局题号区间
SEC_RANGE = {"一": (1, 5), "二": (6, 25), "三": (26, 30),
             "四": (31, 35), "五": (36, 40)}
SEC_NAME = {"一": "名词复数", "二": "单项选择", "三": "选词填空",
            "四": "按要求改写句子", "五": "阅读理解"}


def kps_for_wrong(wrong_qs):
    """错题号 -> 需要补的知识点 key 列表（按出现顺序去重）。"""
    out = []
    for q in sorted(wrong_qs):
        for k in Q2KP.get(q, []):
            if k not in out:
                out.append(k)
    return out


def kps_for_section(sec_key):
    """没有逐题数据时，按大题给出该大题相关的全部知识点。"""
    a, b = SEC_RANGE[sec_key]
    out = []
    for q in range(a, b + 1):
        for k in Q2KP.get(q, []):
            if k not in out:
                out.append(k)
    return out

# ---------- 短名（写进话术用）与一句话说明 ----------
SHORT = {
    "复数规则": "名词复数",
    "复数es": "名词复数（加 es）",
    "复数不规则": "名词复数（不规则变化）",
    "be动词": "be 动词（am / is / are）",
    "物主代词": "物主代词",
    "名词所有格": "名词所有格",
    "冠词": "冠词 a / an",
    "疑问词": "四个疑问词",
    "交际用语": "日常交际用语",
    "look at": "look at 的用法",
    "选词Chinese": "Chinese 的用法",
    "选词代词": "人称代词",
    "改写复数句": "单复数句转换",
    "改写提问": "对划线部分提问",
    "职业词汇": "职业词汇",
}

TIP = {
    "复数规则": "一般名词直接加 s，这个最基础",
    "复数es": "以 s、x、ch、sh 结尾的要加 es，很容易漏",
    "复数不规则": "woman 变 women、sheep 单复数一样，只能单独背",
    "be动词": "am / is / are 要跟着主语走，单数用 is、复数用 are，这里一混就连着错",
    "物主代词": "my / your / his / her 这几个，还有 It's（它是）和 Its（它的）特别容易混",
    "名词所有格": "人名后面加 's，写成 Kates 或 Kates' 就错了",
    "冠词": "元音开头要用 an，孩子常常一律写 a",
    "疑问词": "Whose、What make、What nationality、What job，问什么就得用什么",
    "交际用语": "见面问候和回答的固定说法，背下来就能拿分",
    "look at": "look 后面要跟 at 才能接东西，常把 at 漏掉",
    "选词Chinese": "Chinese 单复数一样，不能加 s",
    "选词代词": "he / it / your / you 这几个的位置和作用容易搞混",
    "改写复数句": "This→These、is→are、my→our 要三处一起变，孩子常常只改一处",
    "改写提问": "对划线部分提问，先想清楚问的是什么，再选疑问词",
    "职业词汇": "policeman、nurse、housewife 这些词得先认得，阅读判断才做得下去",
}
