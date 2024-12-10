from PIL import Image
import numpy as np

# 创建一个50x50的RGB图像
size = 50
image = Image.new('RGB', (size, size))
pixels = image.load()

# 创建一个简单的渐变图案
for i in range(size):
    for j in range(size):
        r = int(255 * i / size)
        g = int(255 * j / size)
        b = int(255 * (i + j) / (2 * size))
        pixels[i, j] = (r, g, b)

# 保存图片
image.save('test_images/gradient.png')
