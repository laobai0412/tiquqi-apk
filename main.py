import os
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from openpyxl import load_workbook

class TiquqiApp(App):
    def build(self):
        self.title = "提取器apk"
        layout = BoxLayout(orientation='vertical', padding=20, spacing=20)
        self.label = Label(text="试题提取工具已启动\n请确保 input.xlsx 在“试题提取工具”文件夹中", halign='center')
        btn = Button(text="开始提取", size_hint=(1, 0.2), background_color=(0.9, 0, 0, 1))
        btn.bind(on_release=self.start_process)
        layout.add_widget(self.label)
        layout.add_widget(btn)
        return layout

    def start_process(self, instance):
        # 简化版逻辑，确保不闪退
        self.label.text = "正在处理中..."
        # ... (这里放你的提取逻辑)

if __name__ == "__main__":
    TiquqiApp().run()
