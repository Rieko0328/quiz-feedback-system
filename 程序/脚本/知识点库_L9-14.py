# -*- coding: utf-8 -*-
r"""NCE1 Lesson 9-14 知识点库（四年级 · 基础偏弱班）。

全局题号：一 单选 1-20 ｜ 二 选词填空 21-25 ｜ 三 改写句子 26-30（每题2分）｜ 四 阅读判断 31-35
"""
KP = {
    # ---- 一 单项选择 1-20 ----
    "物主代词": (
        "物主代词：表示\"谁的\"",
        ["my 我的　your 你的　his 他的　her 她的　its 它的",
         "It's = It is（它是）　Its = 它的　—— 这两个千万别混"],
        [2, 11, 24, 25],
    ),
    "疑问词": (
        "几个常用疑问词，问什么用什么",
        ["Whose 谁的　→ Whose is this book?",
         "What make 什么牌子　→ What make is this car?",
         "What nationality 哪国人　→ What nationality are you?",
         "What ... job 做什么工作　→ What is his job?"],
        [3, 16, 18, 19, 28, 29, 30],
    ),
    "look at": (
        "“看……”的固定说法",
        ["look at the picture（看这张图）",
         "look 后面要跟 at 才能接东西；see 是“看见”，不加 at"],
        [4],
    ),
    "交际用语": (
        "见面打招呼、递东西，这几句要背熟",
        ["Nice to meet you. → Nice to meet you, too.",
         "How are you? → I'm fine, thank you.",
         "Give me some oranges, please. → Here you are.",
         "Are you a teacher? → Yes, I am. / No, I'm not.",
         "Is this your dress? → Yes, it is. / No, it isn't.",
         "Is Tom Japanese? → Yes, he is. / No, he isn't."],
        [1, 5, 8, 9, 14, 15],
    ),
    "be动词": (
        "be 动词怎么选：看前面的主语",
        ["I → am", "You → are", "He / She / It / 单数名词 → is",
         "We / You / They / 两个以上 → are",
         "例：My mother is a doctor.　Tom and Lucy are at school."],
        [6, 7, 10],
    ),
    "名词所有格": (
        "名词所有格：表示\"谁的\"，人名后面加 's",
        ["Kate → Kate's（凯特的）　Tim → Tim's（蒂姆的）",
         "Whose shirt is this? ＝ Whose is this shirt?"],
        [13, 28],
    ),
    "冠词": (
        "冠词 a / an 怎么用",
        ["辅音音开头用 a：a teacher, a nurse, a car",
         "元音音开头用 an：an English girl, an apple"],
        [20],
    ),
    "职业与国籍": (
        "职业和国籍的说法",
        ["What's your job? — I'm a teacher.（问工作）",
         "What nationality are you? — I am Chinese.（问国籍）",
         "Chinese 中国人／中文，单复数一样，不加 s"],
        [17, 18, 19, 21, 32],
    ),
    # ---- 二 选词填空 21-25 ----
    "选词Chinese": (
        "China 和 Chinese 别混",
        ["China 中国（国家名）　Chinese 中国人／中文（语言、国籍）",
         "例：Is he Chinese?　I am Chinese."],
        [21],
    ),
    "选词job": (
        "job 和 work 别混",
        ["job 是可数名词：What's your job?（你做什么工作）",
         "work 是不可数动词/名词：I work in a school."],
        [22],
    ),
    "选词代词": (
        "人称代词：主格和所有格别混",
        ["he 他（做主语）　it 它（做主语）",
         "your 你的（后面要跟名词）　you 你（做主语）",
         "例：Is this your umbrella?　It is a German car."],
        [23, 24, 25],
    ),
    # ---- 三 改写句子 26-30 ----
    "变否定句": (
        "改否定句：be 动词后面加 not，实义动词用 don't / didn't",
        ["This is my coat. → This is not / isn't my coat.",
         "He is French. → He is not / isn't French.",
         "She has a book. → She doesn't have a book."],
        [26],
    ),
    "变一般疑问句": (
        "改一般疑问句：be 动词提到句首，首字母大写",
        ["She is French. → Is she French?",
         "This is my coat. → Is this my coat?",
         "回答：Yes, she is. / No, she isn't."],
        [27],
    ),
    "划线提问": (
        "对划线部分提问：先想清楚问的是什么，再选疑问词",
        ["问“谁的”：Whose shirt is this?",
         "问“什么牌子”：What make is it?",
         "问“做什么工作”：What is your job? / What do you do?"],
        [28, 29, 30],
    ),
    # ---- 四 阅读判断 31-35 ----
    "职业词汇": (
        "家庭成员的职业，这几个词要认得住",
        ["policeman 警察　nurse 护士　housewife 家庭主妇",
         "keyboard operator 打字员　teacher 老师　doctor 医生",
         "family 家庭　grandma 奶奶　brother 兄弟"],
        [33, 34, 35],
    ),
}

Q2KP = {}
for _k, (_t, _items, _qs) in KP.items():
    for _q in _qs:
        Q2KP.setdefault(_q, []).append(_k)

SEC_RANGE = {"一": (1, 20), "二": (21, 25), "三": (26, 30), "四": (31, 35)}
SEC_NAME = {"一": "单项选择", "二": "选词填空", "三": "按要求改写句子", "四": "阅读理解"}


def kps_for_wrong(wrong_qs):
    out = []
    for q in sorted(wrong_qs):
        for k in Q2KP.get(q, []):
            if k not in out:
                out.append(k)
    return out


def kps_for_section(sec_key):
    a, b = SEC_RANGE[sec_key]
    out = []
    for q in range(a, b + 1):
        for k in Q2KP.get(q, []):
            if k not in out:
                out.append(k)
    return out


# ---------- 话术里的短名与一句话说明 ----------
SHORT = {
    "物主代词": "物主代词",
    "疑问词": "四个疑问词",
    "look at": "look at 的用法",
    "交际用语": "日常交际用语",
    "be动词": "be 动词（am / is / are）",
    "名词所有格": "名词所有格",
    "冠词": "冠词 a / an",
    "职业与国籍": "职业和国籍的说法",
    "选词Chinese": "China 和 Chinese",
    "选词job": "job 和 work",
    "选词代词": "人称代词",
    "变否定句": "改否定句",
    "变一般疑问句": "改一般疑问句",
    "划线提问": "对划线部分提问",
    "职业词汇": "职业词汇",
}

TIP = {
    "物主代词": "my / your / his / her 这几个，还有 It's（它是）和 Its（它的）特别容易混",
    "疑问词": "Whose、What make、What nationality、What job，问什么就得用什么",
    "look at": "look 后面要跟 at 才能接东西，常把 at 漏掉",
    "交际用语": "见面问候和回答的固定说法，背下来就能拿分",
    "be动词": "am / is / are 要跟着主语走，单数用 is、复数用 are",
    "名词所有格": "人名后面加 's，写成 Kates 就错了",
    "冠词": "元音开头要用 an，孩子常常一律写 a",
    "职业与国籍": "问工作和问国籍的句型不一样，别弄反",
    "选词Chinese": "China 是国家，Chinese 是中国人和中文，括号里二选一别选错",
    "选词job": "job 是可数名词，work 不可数 —— What's your job? 用 job",
    "选词代词": "he / it / your / you 这几个的位置和作用容易搞混",
    "变否定句": "be 动词后面加 not 就行，位置别放错",
    "变一般疑问句": "be 动词提到句子最前面，第一个字母改成大写",
    "划线提问": "先想清楚问的是什么，再选疑问词，别直接照抄原句",
    "职业词汇": "policeman、nurse、housewife 这些词得先认得，阅读判断才做得下去",
}
