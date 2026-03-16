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
        self.label = Label(text="试题提取工具 (Android)\n请确保 input.xlsx 在“试题提取工具”文件夹中", halign='center')
        btn = Button(text="开始提取", size_hint=(1, 0.3), background_color=(0.9, 0, 0, 1))
        btn.bind(on_release=self.start_process)
        layout.add_widget(self.label)
        layout.add_widget(btn)
        return layout

    def start_process(self, instance):
        target_path = "/storage/emulated/0/试题提取工具/input.xlsx"
        if not os.path.exists(target_path):
            self.label.text = "错误：找不到 /试题提取工具/input.xlsx"
            return
        self.label.text = "正在处理 Excel..."
        # 你的提取逻辑写在这里...
        self.label.text = "处理完成！"

if __name__ == "__main__":
    TiquqiApp().run()
