from PIL import Image, ImageChops

img1 = Image.open("two diff images/Alien1.JPG")
img2 = Image.open("two diff images/Alien2.JPG")

diff = ImageChops.difference(img1,img2)
if diff.getbbox():
    diff.show()
