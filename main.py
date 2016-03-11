import barcode
import openpyxl
from openpyxl.drawing.image import Image as OpenPyXlImage
from PIL import Image as pil_Image
import ImageOps as pil_ImageOps
import os
import math
from barcode.writer import ImageWriter
from Tkinter import *
from ttk import *
from tkFileDialog import *


root_window = Tk()

root_window.title("Barcode Insert Utility")

old_workbook_path = ""
new_workbook_path = ""


def select_folder_old_new_wrapper(selection):
    global old_workbook_path
    global new_workbook_path
    if selection is "old":
        old_workbook_path = askopenfilename()
        if os.path.exists(old_workbook_path):
            old_workbook_label.configure(text=old_workbook_path)
    else:
        new_workbook_path = asksaveasfilename()
        if os.path.exists(os.path.dirname(new_workbook_path)):
            new_workbook_label.configure(text=new_workbook_path)
    if os.path.exists(old_workbook_path) and os.path.exists(os.path.dirname(new_workbook_path)):
        process_workbook_button.configure(state=NORMAL, text="Process Workbook")


def do_process_workbook():
    wb = openpyxl.load_workbook(old_workbook_path)
    ws = wb.worksheets[0]
    count = 1
    save_counter = 1
    progress_bar.configure(maximum=ws.max_row, value=count)
    list_of_temp_images = []
    border_size = int(border_spinbox.get())

    for _ in ws.iter_rows():  # iterate over all rows in current worksheet
        try:
            # get code from column "B", on current row, add a zero to the end to make seven digits
            upc_barcode_number = ws["B" + str(count)].value + "0"
            # select barcode type, specify barcode, and select image writer to save as png
            ean = barcode.get('ean8', upc_barcode_number, writer=ImageWriter())
            # select output image size via dpi. internally, pybarcode renders as svg, then renders that as a png file.
            # dpi is the conversion from svg image size in mm, to what the image writer thinks is inches.
            ean.default_writer_options['dpi'] = int(dpi_spinbox.get())
            # module height is the barcode bar height in mm
            ean.default_writer_options['module_height'] = float(height_spinbox.get())
            # text distance is the distance between the bottom of the barcode, and the top of the text in mm
            ean.default_writer_options['text_distance'] = 1
            # font size is the text size in pt
            ean.default_writer_options['font_size'] = 6
            # quiet zone is the distance from the ends of the barcode to the ends of the image in mm
            ean.default_writer_options['quiet_zone'] = 2
            filename = ean.save("barcode " + str(upc_barcode_number))  # save barcode image with generated filename
            list_of_temp_images.append(str(filename))  # add image to list of files to remove after run
            barcode_image = pil_Image.open(str(filename))  # open image as pil object
            img_save = pil_ImageOps.expand(barcode_image, border=border_size, fill='white')  # add border around image
            width, height = img_save.size  # get image size of barcode with border
            # resize cell to size of image
            ws.column_dimensions['A'].width = int(math.ceil(float(width) * .15))
            ws.row_dimensions[count].height = int(math.ceil(float(height) * .75))
            # write out image to file
            img_save.save("barcode " + str(upc_barcode_number) + 'BORDER' + '.png')
            # add image to list of files to remove after run
            list_of_temp_images.append("barcode " + str(upc_barcode_number) + 'BORDER' + '.png')
            # open image with as openpyxl image object
            img = OpenPyXlImage("barcode " + str(upc_barcode_number) + 'BORDER' + '.png')
            # attach image to cell
            img.anchor(ws.cell('A' + str(count)), anchortype='oneCell')
            # add image to cell
            ws.add_image(img)
            # This save in the loop frees references to the barcode images,
            #  so that python's garbage collector can clear them
            if save_counter == 500:
                    # noinspection PyBroadException
                try:
                    wb.save(new_workbook_path)
                except:
                    print("Cannot write to output file")
                save_counter = 1
            save_counter += 1
        except Exception, error:
            print(error)
        finally:
            count += 1
            progress_bar.configure(value=count)
            progress_bar_frame.update()
    progress_bar.configure(value=0)
    progress_bar_frame.update()
    # noinspection PyBroadException
    try:
        wb.save(new_workbook_path)
    except:
        print("Cannot write to output file")
    finally:
        progress_bar.configure(maximum=len(list_of_temp_images), value=count)
        count = 0
        for line in list_of_temp_images:
            # noinspection PyBroadException
            try:
                os.remove(line)
            except:
                print(line + " missing")
            finally:
                count += 1
                progress_bar.configure(value=count)
                progress_bar_frame.update()
        progress_bar.configure(value=0)
        progress_bar_frame.update()


def process_workbook_command_wrapper():
    global new_workbook_path
    process_workbook_button.configure(state=DISABLED, text="Processing Workbook")
    do_process_workbook()
    new_workbook_path = ""
    new_workbook_label.configure(text="No File Selected")
    process_workbook_button.configure(text="Done Processing Workbook")


both_workbook_frame = Frame(root_window)
old_workbook_file_frame = Frame(both_workbook_frame)
new_workbook_file_frame = Frame(both_workbook_frame)
go_and_progress_frame = Frame(root_window)
go_button_frame = Frame(go_and_progress_frame)
progress_bar_frame = Frame(go_and_progress_frame)
size_spinbox_frame = Frame(root_window)

dpi_spinbox = Spinbox(size_spinbox_frame, from_=120, to=400, width=3, justify=RIGHT)
height_spinbox = Spinbox(size_spinbox_frame, from_=5, to_=50, width=3, justify=RIGHT)
border_spinbox = Spinbox(size_spinbox_frame, from_=0, to_=25, width=2, justify=RIGHT)

old_workbook_selection_button = Button(master=old_workbook_file_frame, text="Select Original Workbook",
                                       command=lambda: select_folder_old_new_wrapper("old")).pack(anchor='w')

new_workbook_selection_button = Button(master=new_workbook_file_frame, text="Select New Workbook",
                                       command=lambda: select_folder_old_new_wrapper("new")).pack(anchor='w')

old_workbook_label = Label(master=old_workbook_file_frame, text="No File Selected")
new_workbook_label = Label(master=new_workbook_file_frame, text="No File Selected")
old_workbook_label.pack(anchor='w')
new_workbook_label.pack(anchor='w')
size_spinbox_dpi_label = Label(master=size_spinbox_frame, text="Barcode DPI:", )
size_spinbox_dpi_label.grid(row=0, column=0, sticky=W + E)
size_spinbox_dpi_label.columnconfigure(0, weight=1)  # make this stretch to fill available space
size_spinbox_height_label = Label(master=size_spinbox_frame, text="Barcode Height:")
size_spinbox_height_label.grid(row=1, column=0, sticky=W + E)
size_spinbox_height_label.columnconfigure(0, weight=1)  # make this stretch to fill available space
border_spinbox_label = Label(master=size_spinbox_frame, text="Barcode Border:")
border_spinbox_label.grid(row=2, column=0, sticky=W + E)
border_spinbox_label.columnconfigure(0, weight=1)  # make this stretch to fill available space
dpi_spinbox.grid(row=0, column=1, sticky=E)
height_spinbox.grid(row=1, column=1, sticky=E)
border_spinbox.grid(row=2, column=1, sticky=E)

process_workbook_button = Button(master=go_button_frame, text="Select Workbooks",
                                 command=process_workbook_command_wrapper)

process_workbook_button.configure(state=DISABLED)

process_workbook_button.pack()

progress_bar = Progressbar(master=progress_bar_frame)
progress_bar.pack()

old_workbook_file_frame.pack(anchor='w')
new_workbook_file_frame.pack(anchor='w')
both_workbook_frame.grid(row=0, column=0, sticky=W)
both_workbook_frame.columnconfigure(0, weight=1)
size_spinbox_frame.grid(row=0, column=1, sticky=E + N)
size_spinbox_frame.columnconfigure(0, weight=1)
go_button_frame.pack(side=LEFT, anchor='w')
progress_bar_frame.pack(side=RIGHT, anchor='e')
go_and_progress_frame.grid(row=1, column=0, columnspan=2, sticky=W + E)
go_and_progress_frame.columnconfigure(0, weight=1)
root_window.columnconfigure(0, weight=1)

root_window.minsize(400, root_window.winfo_height())
root_window.resizable(width=FALSE, height=FALSE)

root_window.mainloop()
