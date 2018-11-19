from PIL import Image
from PIL import ImageEnhance
import pytesseract
import PIL

# var=Image.open('test.png')
# print(pytesseract.image_to_string(Image.open('test.png')))

img = Image.open("screenshot.png")
img = img.convert('L')
img2 = img.crop((334, 570, 396, 597))
#
#
img2.save("img2.png")
#
# # var=Image.open('img2.png')
#
# valor = pytesseract.image_to_string(Image.open('img2.png'))
# print(valor)

# img2 = ImageEnhance.Contrast(img2)

'''CONTRAST TREATMENT'''
img = Image.open("img2.png")
contrast = ImageEnhance.Contrast(img)
img = contrast.enhance(10)
img.save("img3.png")
print(pytesseract.image_to_string(Image.open('img3.png')))