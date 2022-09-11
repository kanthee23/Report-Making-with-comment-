from PIL import Image, ImageEnhance, ImageFilter, ImageMath, PSDraw, ImageDraw, ImageFont
import os
from openpyxl import Workbook, load_workbook


size = (512, 512)
path = './images/'
outpath = './out/'

# input_file = 'img2.jpg'
out_file = 'output.pdf'
temp_file_Front = 'tempFile_front.jpeg'
temp_file_join = 'join file.jpeg'

excel_filename = 'detail.xlsx'
white = (255, 255, 255)
black = (0, 0, 0)
yellow = (255, 255, 0)
body_font_colour = yellow


def makeThambnailAll():
    try:
        for filename in os.listdir(path):
            # print(filename)
            with Image.open(path + filename) as im:
                im.thumbnail(size)
                im.save(outpath + filename, "JPEG")
    except OSError:
        pass

def makeThambnailImage(path, fname, size):
    try:
        # print(path + fname)
        with Image.open(path + fname) as im:
            im.thumbnail(size)
            im.save(outpath + fname, "JPEG")
    except OSError:
        print('file not found')


def joinImages(width_n, height_n):
    im_w = Image.open(outpath+ os.listdir(outpath)[0]).size[0]
    im_h = Image.open(outpath+ os.listdir(outpath)[0]).size[1]
    w = im_w * width_n
    h = im_h * height_n

    newImage = Image.new("RGB", (w, h))

    files = os.listdir(outpath)
    k = 0
    for i in range(height_n):
        for j in range(width_n):
            im = Image.open(outpath + files[k])
            newImage.paste(im, (im_w * i, im_h * j))
            k += 1

    newImage.save(outpath + temp_file_join, "JPEG")


def read_excel_data(fileName):
    image_array = []
    wb = load_workbook(fileName)
    sheet = wb['Sheet1']

    heading_1 = sheet['A1'].value
    heading_1_details = sheet['B1'].value
    heading_2 = sheet['A2'].value
    heading_2_details = sheet['B2'].value
    heading_3 = sheet['A3'].value
    heading_3_details = sheet['B3'].value
    arial_h1 = ImageFont.FreeTypeFont('C:/windows/Fonts/arial.ttf', size = 16)
    arial_h2 = ImageFont.FreeTypeFont('C:/windows/Fonts/arial.ttf', size=14)
    arial_h3 = ImageFont.FreeTypeFont('C:/windows/Fonts/arial.ttf', size=12)
    arial_body = ImageFont.FreeTypeFont('C:/windows/Fonts/arial.ttf', size=12)

    newImage = Image.new("RGB", (1024, 1024), color = white)
    if heading_1_details is not None:
        id = ImageDraw.Draw(newImage)
        id.text((512, 10),heading_1 + " : " + heading_1_details,font = arial_h1, fill = black, align = "center", anchor='mm')

    if heading_2_details is not None:
        id = ImageDraw.Draw(newImage)
        id.text((512, 25), heading_2 + " : " + heading_2_details, font=arial_h2, fill = black, align="center",
                anchor='mm')

    if heading_3_details is not None:
        id = ImageDraw.Draw(newImage)
        id.text((512, 40), heading_3 + " : " + heading_3_details, font=arial_h3, fill = black, align="center",
                anchor='mm')
    loop = True
    num = 5
    page = 1
    row = 0
    column = 0
    while loop:
        img_name = sheet['A' + str(num)].value
        if img_name is not None:
            comment = sheet['B' + str(num)].value
            newComment = ''
            c_loop = True
            x = len(comment)
            y = 0
            while c_loop:
                if x - y >= 50:
                    newComment += comment[y : y + 50] + "\n"

                    y += 50
                else:
                    newComment += comment[y: ] + "\n"
                    c_loop = False

            comment = newComment

            if os.path.isfile(path + img_name):

                makeThambnailImage(path, img_name,(512, 512))
                im = Image.open(outpath + img_name)
                pos = (column * 512 + 10 , row * 512 + 50)
                text_pos = (column * 512 + 50, row * 512 + 400)
                newImage.paste(im, pos)
                if comment is not None:
                    id = ImageDraw.Draw(newImage)
                    id.text(text_pos, text = comment,  font = arial_body, fill = body_font_colour)

                column += 1
                if column >= 2:
                    row += 1
                    column = 0
                    if row >= 2:
                        newImage.save(outpath + str(page) + "_" + temp_file_Front, "JPEG")
                        image_array.append(outpath + str(page) + "_" + temp_file_Front)
                        newImage.paste(Image.new("RGB", (1024, 974), color = white ), (0, 50))
                        page += 1
                        row = 0
                        column = 0

            else:
                print("File not available.")
        else:
            loop = False

        num += 1

    newImage.save(outpath + str(page) + "_" + temp_file_Front, "JPEG")

    image_array.append(outpath + str(page) + "_" + temp_file_Front)

    images = [Image.open(f) for f in image_array]
    pdf = outpath + "Report.pdf"
    images[0].save(pdf,"PDF", resolution =100.0, save_all=True, append_images=images[1:])
    print("PDF Saved.. ")


read_excel_data(excel_filename)