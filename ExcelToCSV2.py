import dearpygui.dearpygui as dpg
import math
import pandas as pd
import os
import sys


cwd = os.path.join(os.path.expanduser("~"), "Desktop")
os.chdir(cwd)

def print_input_text():
    # print(dpg.get_value("__input_text"))
    
    fileInput=dpg.get_value("__input_text")
    # Append .xlsx to file name to save time
    fileInput_ext = fileInput + ".xlsx"

    # Number of rows per csv file
    chunksize = 20

    # Used for naming output files
    count = -1
    
    # Read given excel file
    df = pd.read_excel(fileInput_ext)

    # Calculate number of files to be created
    file_count = math.ceil(len(df) / chunksize)

    for i in range(file_count):
        start = i * chunksize
        end = start + chunksize
        chunk = df.iloc[start:end]
        fileOutput = os.path.join(cwd, f"{fileInput}_{count+1}.csv")
        print(fileOutput)
        chunk.to_csv(fileOutput, index=None, header=None)
        count += 1
    # print(dpg.get_value("__input_text"))

dpg.create_context()

# Change banner bar color
with dpg.theme() as theme_id:
    with dpg.theme_component(dpg.mvAll):
        dpg.add_theme_color(dpg.mvThemeCol_TitleBg, (53, 143, 53))  # Change to desired RGB color
        dpg.add_theme_color(dpg.mvThemeCol_TitleBgActive, (53, 143, 53))  # Active title bar color
        dpg.add_theme_color(dpg.mvThemeCol_TitleBgCollapsed, (53, 143, 53))  # Collapsed title bar color

dpg.create_viewport(title="Blind Enterprises",width=600,height=300)
dpg.setup_dearpygui()


# pyinstaller loads images into a temporary folder
# handles if temp folder is active or not
if getattr(sys, "frozen", False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

image_path = os.path.join(base_dir, "BEO.JPG")
width, height, channels, data = dpg.load_image(image_path)
with dpg.texture_registry(show=False):
    dpg.add_static_texture(width=width, height=height, default_value=data, tag="texture_tag")

with dpg.font_registry():
    arial_font = dpg.add_font("C:/Windows/Fonts/arial.ttf", 20)  # Adjust the path and size as needed

with dpg.window(label="Excel to CSV",width=601,height=301, no_close=True, no_collapse=True) as main_window:
    dpg.set_item_theme(main_window,theme_id)
    dpg.bind_font(arial_font)
    dpg.add_text("Input the file name (no extension), then press Enter or the button.")
    dpg.add_input_text(tag="__input_text", hint="File name", on_enter=True, callback=print_input_text)
    dpg.add_button(arrow=True,direction=1,callback=print_input_text)
    dpg.add_text("Output file name(s) wil be InputFileName_x.csv")
    dpg.add_image("texture_tag",pos=(225,200))

dpg.show_viewport()

dpg.start_dearpygui()
dpg.destroy_context()



# def save_callback():
#     print("Save Clicked")
# os.chdir("C:\\Users\\tami\\OneDrive\\Desktop")
# os.chdir("C:\\Users\\tanne\\Desktop")