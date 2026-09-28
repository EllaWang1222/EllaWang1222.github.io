from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
CONTENT_DIR = ROOT / "content"
PREFERRED_SOURCE = CONTENT_DIR / "个人简历.xlsx"
ASSETS = ROOT / "content" / "assets"
OUTPUT = ROOT / "dist" / "data" / "site-content.json"
PUBLIC_ASSETS = ROOT / "dist" / "assets"
ROOT_DATA = ROOT / "data" / "site-content.json"
ROOT_ASSETS = ROOT / "assets"

IMAGE_PATTERN = re.compile(r"\.(?:avif|gif|jpe?g|png|svg|webp)$", re.IGNORECASE)
SPLIT_PATTERN = re.compile(r"\s*\|\|\s*")


def find_source() -> Path:
    if PREFERRED_SOURCE.exists():
        return PREFERRED_SOURCE

    workbooks = sorted(
        path for path in CONTENT_DIR.glob("*.xlsx")
        if not path.name.startswith("~$")
    )
    if len(workbooks) == 1:
        return workbooks[0]
    if not workbooks:
        raise FileNotFoundError(f"No .xlsx workbook found in: {CONTENT_DIR}")
    names = ", ".join(path.name for path in workbooks)
    raise FileNotFoundError(
        f"Multiple .xlsx workbooks found ({names}). Rename the website workbook to 个人简历.xlsx."
    )


SOURCE = find_source()

TRANSLATIONS = {
    "Sheet1": "Home",
    "首页": "Home",
    "基本信息": "Profile",
    "教育经历": "Education",
    "校园经历": "Campus Experience",
    "工具技能": "Creative Toolkit",
    "语言能力": "Languages",
    "自我介绍": "About Me",
    "王佳辰(Wang Jiachen)": "Wang Jiachen",
    "2025.9-2029.6，上海交通大学，外国语学院德语专业，学士": "Sep 2025 – Jun 2029 · B.A. in German, School of Foreign Languages, Shanghai Jiao Tong University",
    "1、AIGC互动解谜游戏制作": "1. AIGC Interactive Puzzle Game",
    "与2名同学合作，利用AIGC工具及Mediapipe技术设计并完成一款简易互动解谜游戏。": "Collaborated with two classmates to design and complete an interactive puzzle game using AIGC tools and MediaPipe.",
    "负责游戏剧情构思、AI生成内容及对视频素材进行剪辑与整体调优，确保叙事连贯性与可玩性。": "Developed the story, created AI-assisted content, and edited and refined video assets to strengthen narrative flow and playability.",
    "借助AI生成角色对白、场景描述及部分美术素材，探索低成本快速原型开发流程。": "Used AI to create character dialogue, scene descriptions, and selected art assets while exploring a fast, low-cost prototyping workflow.",
    "最终产出可运行的互动游戏Demo，团队内部测试反馈良好，锻炼了创意落地与跨角色协作能力。": "Delivered a working interactive game demo that received positive internal feedback and strengthened cross-functional collaboration skills.",
    "2、校学生会 宣传部成员": "2. Communications Team · Student Union",
    "为学生会官方号设计过小红书推文，使用可画完成内容选题、图文创作与排版，用于活动宣传。": "Created Xiaohongshu posts for the Student Union, covering topic selection, visual production, and layout for event promotion.",
    "在学校游园会活动中担任现场采访者，撰写采访提纲、现场记录，为后续宣传提供素材。": "Worked as an on-site interviewer at the university fair, preparing interview outlines and gathering material for follow-up communications.",
    "3、班级宣传委员": "3. Class Communications Representative",
    "负责班级及团支部宣发工作，独立完成团日活动公众号推文的撰写与排版，并获得辅导员肯定。": "Led communications for the class and Youth League branch, independently writing and designing a WeChat article recognized by the academic advisor.",
    "熟练使用秀米进行公众号推文的撰写与排版，熟悉从文稿撰写到图文美化、格式调整的全流程。": "Proficient in Xiumi for writing and laying out WeChat articles, from drafting and visual polish to final formatting.",
    "能够运用可画设计小红书推文及活动海报，擅长根据校园传播场景输出符合品牌调性的视觉内容。": "Skilled in Canva-style visual design with Kawa, creating Xiaohongshu posts and event posters suited to campus communication contexts.",
    "掌握剪映的基础剪辑功能，包括剪辑、字幕添加、配乐等，可独立完成短视频的后期制作。": "Comfortable with Jianying/CapCut editing, including cutting, subtitles, and music, and able to complete short-form video post-production independently.",
    "熟悉可灵、即梦、豆包等AI生成工具，可辅助完成文本创作、图像生成等内容的快速迭代。": "Familiar with generative AI tools including Kling, Jimeng, and Doubao for rapid iteration of written and visual content.",
    "具备线下采访和素材收集能力，能完成采访提纲设计、现场记录以及后续宣传内容输出。": "Experienced in field interviews and material gathering, from interview planning and on-site notes to final promotional content.",
    "大学英语四级（CET-4）614分，大学英语六级（CET-6）630分，具备良好的读写与听说能力": "CET-4: 614 · CET-6: 630 · Strong reading, writing, listening, and speaking skills",
    "待填写": "To be added",
}


def bilingual(value: object, *, fallback: str | None = None) -> dict[str, str]:
    text = str(value).strip() if value is not None else ""
    parts = SPLIT_PATTERN.split(text, maxsplit=1)
    if len(parts) == 2:
        return {"zh": parts[0], "en": parts[1]}
    return {"zh": text, "en": TRANSLATIONS.get(text, fallback or text)}


def is_image(value: object) -> bool:
    return bool(value and IMAGE_PATTERN.search(str(value).strip()))


def build() -> None:
    if not SOURCE.exists():
        raise FileNotFoundError(f"Missing workbook: {SOURCE}")

    workbook = load_workbook(SOURCE, data_only=True)
    pages: list[dict[str, object]] = []

    for page_index, sheet in enumerate(workbook.worksheets):
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            continue

        sheet_name = sheet.title
        display_name = "首页" if page_index == 0 and sheet_name == "Sheet1" else sheet_name
        page = {
            "id": f"page-{page_index + 1}",
            "sourceSheet": sheet_name,
            "title": bilingual(display_name, fallback="Home" if page_index == 0 else display_name),
            "modules": [],
        }

        headers = rows[0]
        for column_index, header in enumerate(headers):
            if header is None or not str(header).strip():
                continue
            items = []
            for row in rows[1:]:
                value = row[column_index] if column_index < len(row) else None
                if value is None or not str(value).strip():
                    continue
                text = str(value).strip()
                if is_image(text):
                    items.append({"type": "image", "src": f"assets/{Path(text).name}", "alt": bilingual(str(header))})
                else:
                    items.append({"type": "text", "value": bilingual(text)})
            page["modules"].append({
                "id": f"module-{column_index + 1}",
                "title": bilingual(header),
                "items": items,
            })
        pages.append(page)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    PUBLIC_ASSETS.mkdir(parents=True, exist_ok=True)
    ROOT_DATA.parent.mkdir(parents=True, exist_ok=True)
    ROOT_ASSETS.mkdir(parents=True, exist_ok=True)
    payload = json.dumps({"pages": pages}, ensure_ascii=False, indent=2)
    OUTPUT.write_text(payload, encoding="utf-8")
    ROOT_DATA.write_text(payload, encoding="utf-8")

    if ASSETS.exists():
        for asset in ASSETS.iterdir():
            if asset.is_file():
                shutil.copy2(asset, PUBLIC_ASSETS / asset.name)
                shutil.copy2(asset, ROOT_ASSETS / asset.name)

    for filename in ("index.html", "styles.css", "app.js"):
        shutil.copy2(ROOT / "dist" / filename, ROOT / filename)

    print(f"Built {len(pages)} page(s) from {SOURCE.name}")


if __name__ == "__main__":
    build()
