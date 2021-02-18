from dearpygui.core import *
from dearpygui.simple import *


with window("主窗口",height=600,width=400):
    add_additional_font('SourceHanSansCN-Regular-2.otf',12, glyph_ranges='chinese_simplified_common')
    add_text("运行进度")
    add_button("Run")
    add_input_text("path",default_value="")

if __name__=="__main__":
    start_dearpygui(primary_window="主窗口")