from dearpygui.core import *
from dearpygui.simple import *

add_additional_font('SourceHanSansCN-Regular-2.otf', 18, glyph_ranges='chinese_simplified_common')

def directory_picker(sender, data):
    select_directory_dialog(callback=apply_selected_directory)

def apply_selected_directory(sender, data):
    log_debug(data)
    directory = data[0]
    folder = data[1]
    set_value("目录", directory)
    set_value("文件夹", folder)
    set_value("文件夹路径", f"{directory}\\{folder}")

show_logger()

with window("Tutorial"):
    add_button("目录选择器", callback=directory_picker)
    add_text("目录路径: ")
    add_same_line()
    add_label_text("##dir", source="目录", color=[255, 0, 0])
    add_text("文件夹: ")
    add_same_line()
    add_label_text("##folder", source="文件夹", color=[255, 0, 0])
    add_text("文件夹路径: ")
    add_same_line()
    add_label_text("##folderpath", source="文件夹路径", color=[255, 0, 0])

start_dearpygui()