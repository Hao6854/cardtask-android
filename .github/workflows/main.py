# -*- coding: utf-8 -*-
"""
储值卡与任务管理  v1.0
-------------------------------------------------
纯 Python (Kivy) 编写，兼容 Android 10 (API 29) 及以上。

功能：
  1. 储值卡管理：平台、卡号、卡密、金额；标记售出并记录售出时间；
     一键导出全部数据为 Word(.docx) 文档。
  2. 任务清单：添加任务、打勾完成、删除、自动统计完成数。

数据本地持久化（JSON，保存在应用私有目录，无需存储权限）。
"""

import io
import os
import json
import zipfile
import datetime

from kivy.app import App
from kivy.lang import Builder
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.properties import StringProperty, BooleanProperty
from kivy.utils import platform

APP_VERSION = "1.0"
DATA_FILE = "cards_tasks_data.json"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


# ---------------------------------------------------------------------------
# 通用提示弹窗
# ---------------------------------------------------------------------------
class MsgBox(BoxLayout):
    message = StringProperty("")


def show_message(text, title="提示"):
    box = MsgBox(message=str(text))
    popup = Popup(title=title, content=box, size_hint=(0.88, 0.42))
    box.ids.ok_btn.bind(on_press=popup.dismiss)
    popup.open()


# ---------------------------------------------------------------------------
# 列表行
# ---------------------------------------------------------------------------
class CardRow(BoxLayout):
    card_id = StringProperty("")
    platform_name = StringProperty("")
    card_no = StringProperty("")
    card_secret = StringProperty("")
    amount = StringProperty("")
    sold_text = StringProperty("")

    def mark_sold(self):
        App.get_running_app().root.mark_card_sold(self.card_id)

    def delete(self):
        App.get_running_app().root.delete_card(self.card_id)


class TaskRow(BoxLayout):
    task_id = StringProperty("")
    task_text = StringProperty("")
    done = BooleanProperty(False)

    def on_check(self, checkbox, value):
        App.get_running_app().root.toggle_task(self.task_id, value)

    def delete(self):
        App.get_running_app().root.delete_task(self.task_id)


# ---------------------------------------------------------------------------
# 界面布局（KV）
# ---------------------------------------------------------------------------
KV = r"""
<MsgBox>:
    orientation: 'vertical'
    padding: '16dp'
    spacing: '12dp'
    Label:
        text: root.message
        text_size: self.width - dp(16), None
        size_hint_y: None
        height: self.texture_size[1] + dp(16)
        halign: 'center'
        valign: 'middle'
    Button:
        id: ok_btn
        text: '确定'
        size_hint_y: None
        height: '44dp'

<CardRow>:
    orientation: 'horizontal'
    size_hint_y: None
    height: '128dp'
    padding: '6dp'
    spacing: '6dp'
    canvas.before:
        Color:
            rgba: (0.92, 0.95, 1, 1) if root.sold_text.find('已售出') == 0 else (0.98, 0.98, 0.98, 1)
        Rectangle:
            pos: self.pos
            size: self.size
    BoxLayout:
        orientation: 'vertical'
        size_hint_x: 0.60
        spacing: '2dp'
        Label:
            text: '平台：' + root.platform_name
            halign: 'left'
            text_size: self.width, None
            size_hint_y: None
            height: '24dp'
            font_size: '15sp'
        Label:
            text: '卡号：' + root.card_no
            halign: 'left'
            text_size: self.width, None
            size_hint_y: None
            height: '24dp'
            font_size: '15sp'
        Label:
            text: '卡密：' + root.card_secret
            halign: 'left'
            text_size: self.width, None
            size_hint_y: None
            height: '24dp'
            font_size: '15sp'
        Label:
            text: '金额：¥' + root.amount + '　' + root.sold_text
            halign: 'left'
            text_size: self.width, None
            size_hint_y: None
            height: '28dp'
            font_size: '15sp'
    BoxLayout:
        orientation: 'vertical'
        size_hint_x: 0.40
        spacing: '6dp'
        padding: '4dp'
        Button:
            text: '已售出' if root.sold_text.find('已售出') == 0 else '标记售出'
            disabled: root.sold_text.find('已售出') == 0
            on_press: root.mark_sold()
            size_hint_y: None
            height: '46dp'
            font_size: '15sp'
        Button:
            text: '删除'
            on_press: root.delete()
            size_hint_y: None
            height: '46dp'
            font_size: '15sp'

<TaskRow>:
    orientation: 'horizontal'
    size_hint_y: None
    height: '58dp'
    spacing: '6dp'
    padding: '4dp'
    canvas.before:
        Color:
            rgba: (0.95, 0.95, 0.95, 1)
        Rectangle:
            pos: self.pos
            size: self.size
    CheckBox:
        active: root.done
        on_active: root.on_check(self, self.active)
        size_hint_x: 0.14
    Label:
        text: root.task_text
        halign: 'left'
        valign: 'middle'
        text_size: self.width, None
        size_hint_x: 0.66
        color: (0.55, 0.55, 0.55, 1) if root.done else (0.10, 0.10, 0.10, 1)
        font_size: '16sp'
    Button:
        text: '删除'
        size_hint_x: 0.20
        on_press: root.delete()
        font_size: '14sp'

<RootWidget>:
    orientation: 'vertical'
    TabbedPanel:
        id: panel
        do_default_tab: False
        tab_width: '140dp'
        TabbedPanelItem:
            text: '储值卡管理'
            BoxLayout:
                orientation: 'vertical'
                padding: '10dp'
                spacing: '8dp'
                BoxLayout:
                    orientation: 'vertical'
                    size_hint_y: None
                    height: '224dp'
                    spacing: '6dp'
                    TextInput:
                        id: in_platform
                        hint_text: '平台（如：某某游戏平台）'
                        size_hint_y: None
                        height: '38dp'
                        font_size: '16sp'
                    TextInput:
                        id: in_card_no
                        hint_text: '卡号'
                        size_hint_y: None
                        height: '38dp'
                        font_size: '16sp'
                    TextInput:
                        id: in_card_secret
                        hint_text: '卡密'
                        size_hint_y: None
                        height: '38dp'
                        font_size: '16sp'
                    TextInput:
                        id: in_amount
                        hint_text: '金额（元）'
                        input_type: 'decimal'
                        size_hint_y: None
                        height: '38dp'
                        font_size: '16sp'
                    Button:
                        text: '添加储值卡'
                        size_hint_y: None
                        height: '44dp'
                        on_press: root.add_card()
                        font_size: '16sp'
                Label:
                    id: card_summary
                    text: ''
                    size_hint_y: None
                    height: '26dp'
                    font_size: '14sp'
                    halign: 'left'
                ScrollView:
                    GridLayout:
                        id: card_list
                        cols: 1
                        size_hint_y: None
                        height: self.minimum_height
                        spacing: '6dp'
                Button:
                    text: '导出全部数据为 DOC 文档'
                    size_hint_y: None
                    height: '50dp'
                    on_press: root.export_doc()
                    font_size: '16sp'
        TabbedPanelItem:
            text: '任务清单'
            BoxLayout:
                orientation: 'vertical'
                padding: '10dp'
                spacing: '8dp'
                BoxLayout:
                    orientation: 'horizontal'
                    size_hint_y: None
                    height: '48dp'
                    spacing: '6dp'
                    TextInput:
                        id: in_task
                        hint_text: '输入新任务'
                        size_hint_x: 0.74
                        font_size: '16sp'
                    Button:
                        text: '添加'
                        size_hint_x: 0.26
                        on_press: root.add_task()
                        font_size: '16sp'
                Label:
                    id: task_summary
                    text: ''
                    size_hint_y: None
                    height: '26dp'
                    font_size: '14sp'
                    halign: 'left'
                ScrollView:
                    GridLayout:
                        id: task_list
                        cols: 1
                        size_hint_y: None
                        height: self.minimum_height
                        spacing: '6dp'
"""


# ---------------------------------------------------------------------------
# 主界面逻辑
# ---------------------------------------------------------------------------
class RootWidget(BoxLayout):

    # ---------------------- 数据存取 ----------------------
    @property
    def app(self):
        return App.get_running_app()

    def save(self):
        self.app.save_data()

    @staticmethod
    def new_id():
        return datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")

    # ---------------------- 储值卡 ----------------------
    def add_card(self):
        platform = self.ids.in_platform.text.strip()
        card_no = self.ids.in_card_no.text.strip()
        card_secret = self.ids.in_card_secret.text.strip()
        amount_raw = self.ids.in_amount.text.strip()

        if not platform:
            show_message("请填写平台名称")
            return
        if not card_no:
            show_message("请填写卡号")
            return
        try:
            amount = float(amount_raw)
        except (TypeError, ValueError):
            show_message("金额必须是数字")
            return

        self.app.data["cards"].insert(0, {
            "id": self.new_id(),
            "platform": platform,
            "card_no": card_no,
            "card_secret": card_secret,
            "amount": round(amount, 2),
            "sold": False,
            "sold_time": "",
        })
        self.save()
        for key in ("in_platform", "in_card_no", "in_card_secret", "in_amount"):
            self.ids[key].text = ""
        self.refresh_cards()
        show_message("储值卡添加成功", "成功")

    def mark_card_sold(self, card_id):
        for card in self.app.data["cards"]:
            if card["id"] == card_id and not card["sold"]:
                card["sold"] = True
                card["sold_time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.save()
                self.refresh_cards()
                show_message("已标记为售出，并记录售出时间", "成功")
                return

    def delete_card(self, card_id):
        self.app.data["cards"] = [c for c in self.app.data["cards"] if c["id"] != card_id]
        self.save()
        self.refresh_cards()

    def refresh_cards(self):
        container = self.ids.card_list
        container.clear_widgets()
        cards = self.app.data.get("cards", [])
        total = sum(float(c.get("amount", 0)) for c in cards)
        sold_count = sum(1 for c in cards if c.get("sold"))
        self.ids.card_summary.text = (
            "共 {} 张卡　已售出 {} 张　总金额 ¥{:,.2f}".format(len(cards), sold_count, total)
        )
        for card in cards:
            sold_text = (
                "已售出　售出时间：" + card.get("sold_time", "")
                if card.get("sold") else "未售出"
            )
            container.add_widget(CardRow(
                card_id=card["id"],
                platform_name=card.get("platform", ""),
                card_no=card.get("card_no", ""),
                card_secret=card.get("card_secret", ""),
                amount="{:,.2f}".format(float(card.get("amount", 0))),
                sold_text=sold_text,
            ))

    # ---------------------- 任务 ----------------------
    def add_task(self):
        text = self.ids.in_task.text.strip()
        if not text:
            show_message("请输入任务内容")
            return
        self.app.data["tasks"].insert(0, {
            "id": self.new_id(),
            "text": text,
            "done": False,
            "created": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        self.save()
        self.ids.in_task.text = ""
        self.refresh_tasks()

    def toggle_task(self, task_id, value):
        for task in self.app.data["tasks"]:
            if task["id"] == task_id:
                task["done"] = bool(value)
                self.save()
                self.refresh_tasks()
                return

    def delete_task(self, task_id):
        self.app.data["tasks"] = [t for t in self.app.data["tasks"] if t["id"] != task_id]
        self.save()
        self.refresh_tasks()

    def refresh_tasks(self):
        container = self.ids.task_list
        container.clear_widgets()
        tasks = self.app.data.get("tasks", [])
        done_count = sum(1 for t in tasks if t.get("done"))
        self.ids.task_summary.text = "共 {} 项任务　已完成 {} 项".format(len(tasks), done_count)
        for task in tasks:
            container.add_widget(TaskRow(
                task_id=task["id"],
                task_text=task.get("text", ""),
                done=bool(task.get("done", False)),
            ))

    # ---------------------- 导出 DOCX ----------------------
    def export_doc(self):
        cards = self.app.data.get("cards", [])
        try:
            docx_bytes = build_docx(cards)
            filename = "储值卡数据_{}.docx".format(
                datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            )
            ok, message = save_export_file(docx_bytes, filename)
            show_message(message, "导出结果" if ok else "导出失败")
        except Exception as exc:  # noqa: BLE001 - 尽量给用户友好提示
            show_message("导出失败：{}".format(exc), "导出失败")


# ---------------------------------------------------------------------------
# DOCX 生成（纯标准库，无需 python-docx，方便打包）
# ---------------------------------------------------------------------------
def xml_escape(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Calibri" w:eastAsia="SimSun" w:hAnsi="Calibri"/>
<w:sz w:val="20"/><w:szCs w:val="20"/>
</w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>"""

BORDERS = """<w:tblBorders>
<w:top w:val="single" w:sz="4" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:color="auto"/>
</w:tblBorders>"""


def build_docx(cards):
    """把储值卡列表生成一个真正的 .docx（Word/WPS 可直接打开）。"""
    headers = ["平台", "卡号", "卡密", "金额(元)", "状态", "售出时间"]
    widths = [1700, 1900, 1900, 1200, 1000, 2300]

    def cell(text, width, header=False):
        jc = '<w:jc w:val="center"/>' if header else ""
        rpr = '<w:rPr><w:b/><w:sz w:val="20"/></w:rPr>' if header else '<w:rPr><w:sz w:val="20"/></w:rPr>'
        return (
            '<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>'
            '<w:p><w:pPr>{jc}</w:pPr><w:r>{rpr}'
            '<w:t xml:space="preserve">{t}</w:t></w:r></w:p></w:tc>'
        ).format(w=width, jc=jc, rpr=rpr, t=xml_escape(text))

    rows = ["<w:tr>" + "".join(cell(h, w, True) for h, w in zip(headers, widths)) + "</w:tr>"]
    for card in cards:
        values = [
            card.get("platform", ""),
            card.get("card_no", ""),
            card.get("card_secret", ""),
            "{:.2f}".format(float(card.get("amount", 0))),
            "已售出" if card.get("sold") else "未售出",
            card.get("sold_time", "") or "-",
        ]
        rows.append("<w:tr>" + "".join(cell(v, w) for v, w in zip(values, widths)) + "</w:tr>")

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>'
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr>'
        '<w:t>储值卡数据导出</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>导出时间：{now}　　共 {count} 张储值卡</w:t></w:r></w:p>'
        '<w:tbl><w:tblPr><w:tblW w:w="10000" w:type="dxa"/>{borders}</w:tblPr>'
        '<w:tblGrid>{grid}</w:tblGrid>'
        '{rows}</w:tbl>'
        '</w:body></w:document>'
    ).format(
        now=xml_escape(now),
        count=len(cards),
        borders=BORDERS,
        grid="".join('<w:gridCol w:w="{}"/>'.format(w) for w in widths),
        rows="".join(rows),
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        zf.writestr("word/styles.xml", STYLES)
    return buf.getvalue()


def save_export_file(data, filename):
    """桌面端写到当前目录；Android 端写入系统“下载”目录。"""
    if platform == "android":
        try:
            from jnius import autoclass  # 由 buildozer 自动打包

            PythonActivity = autoclass("org.kivy.android.PythonActivity")
            activity = PythonActivity.mActivity
            resolver = activity.getContentResolver()
            MediaStore = autoclass("android.provider.MediaStore")
            Downloads = autoclass("android.provider.MediaStore$Downloads")
            ContentValues = autoclass("android.content.ContentValues")
            Environment = autoclass("android.os.Environment")

            values = ContentValues()
            values.put("_display_name", filename)
            values.put("_mime_type", DOCX_MIME)
            values.put("relative_path", Environment.DIRECTORY_DOWNLOADS)
            values.put("is_pending", 0)
            uri = resolver.insert(Downloads.EXTERNAL_CONTENT_URI, values)
            out = resolver.openOutputStream(uri)
            out.write(bytearray(data))
            out.flush()
            out.close()
            return True, "已导出到手机“下载/Download”目录：\n" + filename
        except Exception as exc:  # noqa: BLE001
            path = os.path.join(App.get_running_app().user_data_dir, filename)
            with open(path, "wb") as fh:
                fh.write(data)
            return False, (
                "无法写入下载目录，已保存到应用内部目录：\n" + path + "\n原因：" + str(exc)
            )
    else:
        try:
            path = os.path.join(os.getcwd(), filename)
            with open(path, "wb") as fh:
                fh.write(data)
            return True, "已导出到：\n" + path
        except Exception:
            path = os.path.join(App.get_running_app().user_data_dir, filename)
            with open(path, "wb") as fh:
                fh.write(data)
            return True, "已导出到应用目录：\n" + path


# ---------------------------------------------------------------------------
# 应用入口
# ---------------------------------------------------------------------------
class CardTaskApp(App):
    title = "储值卡与任务管理 v" + APP_VERSION

    def build(self):
        self.data = self.load_data()
        Builder.load_string(KV)
        root = RootWidget()
        root.refresh_cards()
        root.refresh_tasks()
        return root

    @property
    def data_path(self):
        return os.path.join(self.user_data_dir, DATA_FILE)

    def load_data(self):
        try:
            with open(self.data_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            if not isinstance(data, dict):
                raise ValueError("bad data")
            data.setdefault("cards", [])
            data.setdefault("tasks", [])
            return data
        except Exception:
            return {"cards": [], "tasks": []}

    def save_data(self):
        os.makedirs(self.user_data_dir, exist_ok=True)
        with open(self.data_path, "w", encoding="utf-8") as fh:
            json.dump(self.data, fh, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    CardTaskApp().run()
