import errno
import os
from PIL import Image as img
from wand.image import Image
import re
import shutil
import openpyxl
import time
import autoit
from selenium import webdriver


def pdftoimg(pdf, temp):
    for file in os.listdir(pdf):
        f = os.path.join(pdf, file)
        with(Image(filename=f, resolution=300)) as source:
            images = source.sequence
            newfilename = os.path.join(temp, F"{file.split('.pdf')[0]}.jpeg")
            Image(images[0]).save(filename=newfilename)


def crop(temp, crop):
    for file in os.listdir(temp):
        f = os.path.join(temp, file)
        im = img.open(f)
        if "-ag" in file.lower():
            ftype = "1"
        elif "bajaj" in file.lower():
            ftype = "2"
        elif "satyam" in file.lower():
            ftype = "3"
        else:
            continue

        # Setting the points for cropped image
        if ftype == "1":
            left = 605
            top = 1000
            right = 1550
            bottom = 1800
            im1 = im.crop((left, top, right, bottom))
            im1.save(os.path.join(crop, file))
        elif ftype == "2":
            left = 0
            top = 1200
            right = 750
            bottom = 2200
            im1 = im.crop((left, top, right, bottom))
            im1 = im1.rotate(90)
            # im1 = im1.crop((50,450,850,850))
            im1.save(os.path.join(crop, file))

        elif ftype == "3":
            left = 2155
            top = 50
            right = 2750
            bottom = 350
            im1 = im.crop((left, top, right, bottom))
            im1.save(os.path.join(crop, file))


def read_data(path):
    driver = webdriver.Chrome(os.path.join(chrm_path,'chromedriver.exe'))
    # driver.maximize_window()
    try:
        driver.minimize_window()
        driver.get('http://peculiarventures.github.io/js-zxing-pdf417/examples/zxing-pdf417-example1.html')
        for file in os.listdir(path)[:]:
            driver.find_element_by_id("file").click()
            time.sleep(1)
            autoit.control_focus("Open", "Edit1")
            fname = os.path.join(path, file)
            autoit.control_set_text("Open", "Edit1", fname)
            autoit.control_click("Open", "Button1")
            time.sleep(1)
            data = driver.find_element_by_class_name("decodedText").text
            if data == "[]":
                print("No data found for pdf:",str(file).replace(".jpeg",""))
                continue
            data = data.split('"Text": "')[1].split('"RawBytes"')[0]
            dd = re.sub(r',\n', "", re.sub(r'"', "", re.sub(r' ', "", data)))
            dd = dd.split("\\t")
            while "" in dd:
                dd.remove("")
            base = dd[:11]
            lft = dd[11:]
            splitlist = []
            for x in range(4, len(lft), 4):
                splitlist.append(x)
            res = [lft[i: j] for i, j in zip([0] + splitlist, splitlist + [None])]
            finaldata = []
            for x in res:
                tmp = []
                tmp.extend(base)
                tmp.extend(x)
                finaldata.append(tmp)
            create_excel(finaldata, os.path.join(excelpath, "Barcode_data.xlsx"))
    except:
        driver.quit()

def create_excel(records, fpath):
    if not os.path.isfile(fpath):
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.append(["vendor code", "Po No", "Invocie No.", "Date", "GSTIN", "total invocie amount", "total value",
                      "vehicle no", "SGST", "IGST", "CGST", "Description of Goods", "HSN", "quantity", "basic amount"])
    else:
        book = openpyxl.load_workbook(fpath)
        sheet = book.active

    for lst in records:
        sheet.append(lst)
    book.save(fpath)


def remove_files(*path):
    for file_path in path:
        for f in os.listdir(file_path):
            path = os.path.join(file_path, f)
            os.remove(path)
        os.removedirs(file_path)


def move_files(source, destination, file_list):
    print("Moving Files...")
    if not os.path.isdir(destination):
        make_dir(destination)
    for f in file_list:
        shutil.move(str(os.path.join(source, f)), destination)


def make_dir(*paths):
    for pt in paths:
        if not (os.path.isdir(pt)):
            try:
                os.makedirs(pt, mode=0o777, exist_ok=True)
            except OSError as exception:
                if exception.errno != errno.EEXIST:
                    raise


if __name__ == "__main__":
    temp_path = os.path.join(os.getcwd(), "Files", "Temp")
    chrm_path = os.path.join(os.getcwd(), "Files", "Chrome driver")
    crop_path = os.path.join(os.getcwd(), "Files", "Cropped")
    pdf_path = os.path.join(os.getcwd(), "Files", "PDF")
    excelpath = os.path.join(os.getcwd(), "Files", "Excel")
    make_dir(temp_path, crop_path, pdf_path, excelpath,chrm_path)
    print("Please wait reading pdf...")
    pdftoimg(pdf_path, temp_path)
    crop(temp_path, crop_path)
    print("Starting Data Extraction")
    read_data(crop_path)
    remove_files(temp_path, crop_path)
    print("Data Extraction Completed.")