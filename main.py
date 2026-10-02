from threading import Thread
import re

from kivy.app import App
from kivy.clock import Clock
from kivy.graphics import Color, Line
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.progressbar import ProgressBar
from kivy.uix.scrollview import ScrollView
from kivy.uix.widget import Widget

import whatsapp_export


PROCESS_NAMES = [
    "Find WhatsApp ZIP",
    "Extract WhatsApp ZIP",
    "Find chat TXT",
    "Read chat messages",
    "Build media index",
    "Process messages",
    "Copy media files",
    "Create Excel previews",
    "Save Excel report",
    "Create Report ZIP",
    "Clean temporary files",
    "Completed",
]


class StatusIcon(Widget):
    """
    Real drawn circle / spinner / tick.
    This avoids the square-box problem caused by unsupported
    Unicode characters on some Android fonts.
    """

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.state = "waiting"
        self.angle = 0
        self.bind(pos=self.redraw, size=self.redraw)
        Clock.schedule_interval(self.animate, 1 / 20)

    def set_state(self, state):
        self.state = state
        self.redraw()

    def animate(self, dt):
        if self.state == "active":
            self.angle = (self.angle + 15) % 360
            self.redraw()

    def redraw(self, *args):
        self.canvas.clear()

        with self.canvas:
            if self.state == "complete":
                # Green circle
                Color(0.20, 0.85, 0.40, 1)
                Line(
                    circle=(
                        self.center_x,
                        self.center_y,
                        dp(10)
                    ),
                    width=2.2
                )

                # Tick
                Line(
                    points=[
                        self.center_x - dp(5),
                        self.center_y,
                        self.center_x - dp(1),
                        self.center_y - dp(5),
                        self.center_x + dp(7),
                        self.center_y + dp(6),
                    ],
                    width=2.6,
                    cap="round"
                )

            elif self.state == "active":
                # Animated blue loading circle
                Color(0.20, 0.65, 1, 1)
                Line(
                    circle=(
                        self.center_x,
                        self.center_y,
                        dp(10),
                        self.angle,
                        self.angle + 270
                    ),
                    width=3
                )

            else:
                # Waiting circle
                Color(0.55, 0.55, 0.55, 1)
                Line(
                    circle=(
                        self.center_x,
                        self.center_y,
                        dp(9)
                    ),
                    width=1.7
                )


class ProcessRow(BoxLayout):

    def __init__(self, number, name, **kwargs):
        super().__init__(
            orientation="vertical",
            size_hint_y=None,
            height=dp(57),
            spacing=dp(2),
            **kwargs
        )

        self.number = number
        self.name = name
        self.percent = 0

        top = BoxLayout(
            orientation="horizontal",
            size_hint_y=None,
            height=dp(30),
            spacing=dp(7)
        )

        self.icon = StatusIcon(
            size_hint_x=None,
            width=dp(28)
        )

        self.label = Label(
            text=f"{number}. {name}",
            font_size="14sp",
            halign="left",
            valign="middle"
        )
        self.label.bind(
            size=lambda i, v:
            setattr(i, "text_size", v)
        )

        self.percent_label = Label(
            text="0%",
            font_size="13sp",
            size_hint_x=None,
            width=dp(45),
            halign="right",
            valign="middle"
        )
        self.percent_label.bind(
            size=lambda i, v:
            setattr(i, "text_size", v)
        )

        top.add_widget(self.icon)
        top.add_widget(self.label)
        top.add_widget(self.percent_label)

        self.bar = ProgressBar(
            max=100,
            value=0,
            size_hint_y=None,
            height=dp(5)
        )

        self.add_widget(top)
        self.add_widget(self.bar)

    def set_state(self, state, percent):
        self.percent = max(0, min(100, int(percent)))
        self.bar.value = self.percent
        self.percent_label.text = f"{self.percent}%"
        self.icon.set_state(state)

        if state == "active":
            self.label.text = f"{self.number}. {self.name}  -  Working..."
        else:
            self.label.text = f"{self.number}. {self.name}"


class WhatsAppExporterApp(App):

    def build(self):
        self.title = "WhatsApp Chat Exporter"

        root = BoxLayout(
            orientation="vertical",
            padding=dp(12),
            spacing=dp(5)
        )

        title = Label(
            text="WHATSAPP CHAT EXPORTER",
            font_size="22sp",
            bold=True,
            size_hint_y=None,
            height=dp(38)
        )

        self.status = Label(
            text="Ready",
            font_size="15sp",
            size_hint_y=None,
            height=dp(28)
        )

        self.overall_bar = ProgressBar(
            max=100,
            value=0,
            size_hint_y=None,
            height=dp(8)
        )

        self.overall_label = Label(
            text="Overall progress: 0%",
            font_size="13sp",
            size_hint_y=None,
            height=dp(22)
        )

        process_box = BoxLayout(
            orientation="vertical",
            size_hint_y=None,
            spacing=dp(2)
        )

        self.rows = []

        for number, name in enumerate(PROCESS_NAMES, 1):
            row = ProcessRow(number, name)
            self.rows.append(row)
            process_box.add_widget(row)

        process_box.height = sum(
            row.height for row in self.rows
        )

        scroll = ScrollView(
            do_scroll_x=False
        )
        scroll.add_widget(process_box)

        self.export_button = Button(
            text="START EXPORT",
            font_size="18sp",
            size_hint_y=None,
            height=dp(56)
        )
        self.export_button.bind(
            on_press=self.start_export
        )

        root.add_widget(title)
        root.add_widget(self.status)
        root.add_widget(self.overall_bar)
        root.add_widget(self.overall_label)
        root.add_widget(scroll)
        root.add_widget(self.export_button)

        return root

    def set_process(self, number, state, percent):
        self.rows[number - 1].set_state(
            state,
            percent
        )
        self.update_overall()

    def update_overall(self):
        total = sum(
            row.percent for row in self.rows
        )

        overall = int(
            total / len(self.rows)
        )

        self.overall_bar.value = overall
        self.overall_label.text = (
            f"Overall progress: {overall}%"
        )

    def reset(self):
        for row in self.rows:
            row.set_state("waiting", 0)

        self.overall_bar.value = 0
        self.overall_label.text = (
            "Overall progress: 0%"
        )

    def start_export(self, instance):
        if self.export_button.disabled:
            return

        self.reset()

        self.export_button.disabled = True
        self.export_button.text = "EXPORTING..."
        self.status.text = "Starting export..."

        self.set_process(1, "active", 0)

        Thread(
            target=self.export_worker,
            daemon=True
        ).start()

    def export_worker(self):
        try:
            result = whatsapp_export.run_export(
                progress_callback=self.set_status
            )

            Clock.schedule_once(
                lambda dt:
                self.export_success(result)
            )

        except Exception as error:
            Clock.schedule_once(
                lambda dt:
                self.export_failed(str(error))
            )

    def set_status(self, text):
        Clock.schedule_once(
            lambda dt:
            self.handle_status(text)
        )

    def handle_status(self, text):

        self.status.text = text

        t = text.lower()

        # ----------------------------------------------------
        # 1. FIND ZIP
        # ----------------------------------------------------

        if "searching for whatsapp zip" in t:
            self.set_process(1, "active", 0)

        elif "original whatsapp zip found" in t:
            self.set_process(1, "complete", 100)
            self.set_process(2, "active", 0)

        # ----------------------------------------------------
        # 2. EXTRACT
        # ----------------------------------------------------

        elif "extracting whatsapp zip" in t:
            self.set_process(2, "active", 25)

        elif "extraction completed" in t:
            self.set_process(2, "complete", 100)
            self.set_process(3, "active", 0)

        # ----------------------------------------------------
        # 3. FIND CHAT TXT
        # ----------------------------------------------------

        elif "finding chat txt" in t:
            self.set_process(3, "active", 50)

        elif "chat txt found" in t:
            self.set_process(3, "complete", 100)
            self.set_process(4, "active", 0)

        # ----------------------------------------------------
        # 4. READ CHAT
        # ----------------------------------------------------

        elif "reading chat" in t:
            self.set_process(4, "active", 25)

        elif "chat lines loaded" in t:
            self.set_process(4, "complete", 100)
            self.set_process(5, "active", 0)

        # ----------------------------------------------------
        # 5. MEDIA INDEX
        # ----------------------------------------------------

        elif "media files indexed" in t:
            self.set_process(5, "complete", 100)
            self.set_process(6, "active", 0)

        # ----------------------------------------------------
        # 6. PROCESS MESSAGES
        # ----------------------------------------------------

        elif "processed" in t and "messages" in t:

            match = re.search(
                r"processed\s+(\d+)\s+messages",
                t
            )

            if match:
                count = int(match.group(1))

                # Your current exporter reports every 100 messages.
                # This gives visible movement while processing.
                # It is completed at the Excel-save stage.
                percent = count % 100

                if percent == 0:
                    percent = 1

                self.set_process(
                    6,
                    "active",
                    percent
                )

                # Keep media/previews visually active too because
                # they are generated while each message is written.
                self.set_process(
                    7,
                    "active",
                    min(99, percent)
                )

                self.set_process(
                    8,
                    "active",
                    min(99, percent)
                )

        # ----------------------------------------------------
        # 9. SAVE EXCEL
        # ----------------------------------------------------

        elif "saving excel report" in t:

            self.set_process(
                6, "complete", 100
            )

            self.set_process(
                7, "complete", 100
            )

            self.set_process(
                8, "complete", 100
            )

            self.set_process(
                9, "active", 50
            )

        # ----------------------------------------------------
        # 10. REPORT ZIP
        # ----------------------------------------------------

        elif "creating report zip" in t:

            self.set_process(
                9, "complete", 100
            )

            self.set_process(
                10, "active", 25
            )

        # ----------------------------------------------------
        # 11. CLEANUP
        # ----------------------------------------------------

        # The current engine performs cleanup immediately after
        # ZIP creation. It does not send a status message for it,
        # so completion is handled when the whole export succeeds.

    def export_success(self, result):

        # At this point every operation has really completed.
        # Therefore show 100% for every process.
        for row in self.rows:
            row.set_state(
                "complete",
                100
            )

        self.overall_bar.value = 100
        self.overall_label.text = (
            "Overall progress: 100%"
        )

        self.status.text = (
            "EXPORT COMPLETED"
        )

        self.export_button.disabled = False
        self.export_button.text = (
            "START EXPORT"
        )

    def export_failed(self, error):

        self.status.text = (
            "EXPORT FAILED"
        )

        self.export_button.disabled = False
        self.export_button.text = (
            "TRY AGAIN"
        )

        print(
            "Export error:",
            error
        )


if __name__ == "__main__":
    WhatsAppExporterApp().run()
