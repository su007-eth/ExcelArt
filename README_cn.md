# Excel像素图片转换器

一个将图片转换为Excel像素艺术的Python程序。每个Excel单元格代表图片中的一个像素。程序包含智能像素尺寸检测算法，可以处理各种类型的像素艺术图片。

## 功能特点

- 支持常见图片格式（PNG、JPG、GIF等）
- 智能像素尺寸检测算法
- 大图片自动预处理
- 保持单元格为正方形以确保像素显示正确
- 自动设置单元格颜色来匹配图片像素

## 安装

```bash
pip install -r requirements.txt
```

## 使用方法

1. 命令行方式：
```bash
python excel_pixel_image.py 输入图片路径 输出Excel路径
```

2. 在Python代码中使用：
```python
from excel_pixel_image import convert_image_to_excel

convert_image_to_excel("input.jpg", "output.xlsx")
```

## 像素检测算法

程序包含一个先进的像素检测算法，具有以下特点：
- 自动检测图片是否为像素艺术
- 寻找最佳的像素尺寸
- 可以处理原始和放大的像素艺术图片
- 通过预处理高效处理大尺寸图片
- 提供清晰的检测过程反馈

## 测试图片

`test_images`目录包含多种测试用例：
- `pixel_simple.png`：简单像素艺术
- `pixel_large.png`：大尺寸像素艺术
- `pixel_game.png`：游戏风格像素艺术
- `pixel_icon.gif`：图标风格像素艺术
- `pixel_16.png`：16x16像素艺术
- `pixel_16_enlarged.png`：放大的16x16像素艺术
- `pixel_heart_24_enlarged.png`：放大的心形像素艺术
- `non_pixel_photo.png`：普通照片（非像素艺术）
- `non_pixel_art.png`：数字艺术（非像素艺术）

## 注意事项

- 为了获得最佳显示效果，建议在Excel中将视图缩放比例设置为100%
- 大尺寸图片会自动进行预处理以提高性能
- 单元格被设置为正方形以确保像素显示正确
- 对于非像素艺术图片，程序会建议使用较大的像素尺寸
