from kivy.app import App
from kivy.uix.label import Label

class TestApp(App):
    def build(self):
        return Label(text='打包成功！请放入 input.xlsx')

if __name__ == '__main__':
    TestApp().run()
