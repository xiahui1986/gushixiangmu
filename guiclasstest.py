from dearpygui import core,simple
simple.show_debug()


def directory_picker(sender, data):

    core.select_directory_dialog(callback=apply_selected_directory)


def apply_selected_directory(sender, data):
    core.set_value("path", f"{data[0]}\\{data[1]}")


with simple.window("主窗口", height=600, width=400, no_background=True):
    core.add_additional_font('SourceHanSansCN-Regular-2.otf', 18, glyph_ranges='chinese_simplified_common')
    core.add_text("运行进度")
    core.add_button("Run", )
    core.add_button("选择文件夹",callback=directory_picker )
    core.add_input_text("path", default_value="")
    core.add_input_text("t1", default_value="")

"""
class main:
    def __init__(self):
        self.core=core
        self.simple=simple
        with self.simple.window("主窗口", height=600, width=400,no_background=True):
            self.core.add_additional_font('SourceHanSansCN-Regular-2.otf', 12, glyph_ranges='chinese_simplified_common')
            self.core.add_text("运行进度")
            self.core.add_button("Run",callback=directory_picker)
            self.core.add_button("选择文件夹", )
            self.core.add_input_text("path", default_value="")
            self.core.add_input_text("t1",default_value="")
        self.core.set_value("t1",10)

"""
if __name__=="__main__":
    #a=main()
    core.start_dearpygui(primary_window="主窗口")