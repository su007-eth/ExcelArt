import os
from pathlib import Path
from .excelart import convert_image_to_excel
import time

def batch_test_images(input_dir="test_images"):
    """
    批量测试ExcelArt程序
    :param input_dir: 输入图片目录
    """
    # 获取test_images目录
    test_dir = Path(input_dir)
    
    # 获取所有图片文件
    image_files = [f for f in test_dir.glob("*") 
                  if f.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif']]
    
    print(f"找到 {len(image_files)} 个测试图片")
    print("-" * 50)
    
    # 测试每个图片
    for img_file in sorted(image_files):
        print("\n" + "=" * 50)
        print(f"测试图片: {img_file.name}")
        print("-" * 50)
        
        print(f"处理图片: {img_file}")
        
        # 记录开始时间
        start_time = time.time()
        
        try:
            convert_image_to_excel(str(img_file))  # 不指定输出路径，使用默认路径
            print("转换成功!")
        except Exception as e:
            print(f"转换失败: {e}")
        
        # 计算并打印执行时间
        elapsed_time = time.time() - start_time
        print(f"\n执行时间: {elapsed_time:.2f} 秒")
        print("=" * 50)

if __name__ == "__main__":
    batch_test_images()
