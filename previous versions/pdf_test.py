import io
from PIL import Image
from selenium import webdriver
from selenium import webdriver

driver = webdriver.Firefox(executable_path = '/usr/local/bin/geckodriver')

def get_concat_h_cut(im1, im2):
    dst = Image.new('RGB', (im1.width + im2.width,
                            min(im1.height, im2.height)))
    dst.paste(im1, (0, 0))
    dst.paste(im2, (im1.width, 0))
    return dst


def get_concat_v_cut(im1, im2):
    dst = Image.new(
        'RGB', (min(im1.width, im2.width), im1.height + im2.height))
    dst.paste(im1, (0, 0))
    dst.paste(im2, (0, im1.height))
    return dst

driver.get("https://www.google.com")
a=io.BytesIO(driver.get_screenshot_as_png(
))

driver.get("https://www.facebook.com")
b = io.BytesIO(driver.get_screenshot_as_png(
))

a= Image.open(a)
b = Image.open(b)

get_concat_h_cut(a, b).save('./pillow_concat_h_cut.jpg')
get_concat_v_cut(a, b).save(
    './pillow_concat_v_cut.jpg')