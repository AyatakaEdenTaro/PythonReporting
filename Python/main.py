import shutil
import cv2

BEFORE_IMAGE_PATH = "./before/image.jpg"
AFTER_IMAGE_PATH = "./after/image.jpg"

shutil.copyfile(BEFORE_IMAGE_PATH, AFTER_IMAGE_PATH)

trim_img = cv2.imread(AFTER_IMAGE_PATH)
check_img = cv2.imread(AFTER_IMAGE_PATH, cv2.IMREAD_GRAYSCALE)

print(check_img.shape)

get_trim_x = check_img.shape[1]
get_trim_y = check_img.shape[0]

for x in range(check_img.shape[1]):
    if check_img[0][x] == 0:
        get_trim_x = x - 2
        break

for y in range(check_img.shape[0]):
    if check_img[y][0] == 0:
        get_trim_y = y
        break

print(get_trim_x)
print(get_trim_y)

trim_img = trim_img[0 : get_trim_y,0 : get_trim_x]
cv2.imwrite(AFTER_IMAGE_PATH, trim_img)

