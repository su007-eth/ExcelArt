import os
from excel_art import convert_image_to_excel

def batch_test_images(input_dir="test_images", output_dir="test_results"):
    """
    批量测试ExcelArt程序
    :param input_dir: 输入图片目录
    :param output_dir: 输出Excel文件目录
    """
    # 获取test_images目录
    test_dir = Path(input_dir)
    
    # 获取所有图片文件
    image_files = [f for f in test_dir.glob("*") 
                  if f.suffix.lower() in ['.png', '.jpg', '.gif'] 
                  and not f.stem.endswith('_excel')]
    
    print(f"找到 {len(image_files)} 个测试图片")
    print("-" * 50)
    
    # 测试每个图片
    for img_file in sorted(image_files):
        output_file = Path(output_dir) / f"{img_file.stem}_batch_excel.xlsx"
        
        # 构建命令
        cmd = f"python3 excel_pixel_image.py {img_file} {output_file}"
        
        # 打印分隔线和文件名
        print("\n" + "=" * 50)
        print(f"测试图片: {img_file.name}")
        print("-" * 50)
        
        # 记录开始时间
        start_time = time.time()
        
        # 执行命令
        os.system(cmd)
        
        # 计算并打印执行时间
        elapsed_time = time.time() - start_time
        print(f"\n执行时间: {elapsed_time:.2f} 秒")
        print("=" * 50)

if __name__ == "__main__":
    batch_test_images()
