# ExcelArt

ExcelArt is a Python program that converts images into Excel art, where each Excel cell represents a pixel in the image. It features an intelligent pixel size detection algorithm to handle various types of pixel art images.

## Features

- Supports common image formats (PNG, JPG, GIF, etc.)
- Intelligent pixel size detection for pixel art images
- Automatic preprocessing for large images
- Maintains square cells for accurate pixel display
- Automatically sets cell colors to match image pixels

## Installation

1. Clone this repository
2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

```bash
python excelart.py input_image [output_excel]
```

If output_excel is not specified, it will create an Excel file with the same name as the input image in the same directory.

Example:
```bash
python excelart.py test_images/pixel_art.png
# This will create test_images/pixel_art.xlsx
```

## Pixel Detection Algorithm

The program includes an advanced pixel detection algorithm that:
- Automatically detects if an image is pixel art
- Finds the optimal pixel size for conversion
- Handles both original and enlarged pixel art images
- Efficiently processes large images through preprocessing
- Provides clear feedback about the detection process

## Test Images

The `test_images` directory contains various test cases:
- `pixel_simple.png`: Simple pixel art
- `pixel_large.png`: Large pixel art
- `pixel_game.png`: Game-style pixel art
- `pixel_icon.gif`: Icon-style pixel art
- `pixel_16.png`: 16x16 pixel art
- `pixel_16_enlarged.png`: Enlarged version of 16x16 pixel art
- `pixel_heart_24_enlarged.png`: Enlarged heart pixel art
- `non_pixel_photo.png`: Regular photo (non-pixel art)
- `non_pixel_art.png`: Digital art (non-pixel art)

## Notes

- For best display, set Excel view zoom to 100%
- Large images are automatically preprocessed for better performance
- Cells are set to square shape for accurate pixel display
- The program will suggest using larger pixel sizes for non-pixel art images
