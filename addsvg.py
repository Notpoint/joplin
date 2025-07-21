# -*- coding: utf-8 -*-
from pathlib import Path

def add_svg_extension_to_files(folder_path_str: str):
    """
    遍历指定文件夹，为所有没有后缀名的文件添加 .svg 后缀。
    
    Args:
        folder_path_str (str): 要处理的文件夹的路径字符串。
    """
    try:
        # 使用 pathlib 将字符串路径转换为更健壮的 Path 对象
        target_dir = Path(folder_path_str)
        
        # --- 安全检查 1: 确认文件夹存在 ---
        if not target_dir.is_dir():
            print(f"错误：文件夹 '{folder_path_str}' 不存在或不是一个有效的目录。")
            return
            
        print(f"开始扫描文件夹: {target_dir}")
        
        # 遍历文件夹中的每一个项目（文件或子文件夹）
        for item_path in target_dir.iterdir():
            
            # --- 条件 1: 必须是文件 (跳过子文件夹) ---
            # --- 条件 2: 文件名必须没有后缀 (item_path.suffix 会是空字符串 '') ---
            if item_path.is_file() and not item_path.suffix:
                
                # 创建新的文件名（在原文件名后加上.svg）
                new_path = item_path.with_suffix('.svg')
                
                # --- 安全检查 2: 如果目标文件名（如 "file.svg"）已存在，则跳过以防覆盖 ---
                if new_path.exists():
                    print(f"跳过: 文件 '{item_path.name}' 未重命名，因为 '{new_path.name}' 已存在。")
                    continue
                
                try:
                    # 执行重命名
                    item_path.rename(new_path)
                    print(f"成功: 已将 '{item_path.name}' 重命名为 '{new_path.name}'")
                except Exception as e:
                    print(f"失败: 重命名 '{item_path.name}' 时发生错误: {e}")

    except Exception as e:
        print(f"脚本执行过程中发生意外错误: {e}")
        
    print("\n处理完成。")


# --- 用户配置 ---
if __name__ == '__main__':
    # 【请在这里修改为您需要处理的文件夹路径】
    # 例如，根据您之前的问题，这个路径可能是Joplin导出的资源文件夹。
    # 注意：路径前的 r 表示这是一个“原始字符串”，可以防止Windows路径中的反斜杠被误解。
    FOLDER_TO_PROCESS = r"C:\Users\13684\Desktop\jophin\joplindirectoutput\_resources"
    
    # 调用函数，开始执行
    add_svg_extension_to_files(FOLDER_TO_PROCESS)