# compare_pdf_images.py
"""
对比PDF内容和已提取的图片，创建映射关系
"""

from pathlib import Path
import json

def analyze_pdf_slides():
    """分析PDF幻灯片内容"""
    slides = []
    
    # 读取提取的PPT文本
    try:
        with open("GEOTERRA SDN BHD Company Profile_extracted.txt", "r", encoding="utf-8") as f:
            content = f.read()
        
        # 按幻灯片分割
        slide_sections = content.split("=" * 60)
        
        for i, section in enumerate(slide_sections[1:], 1):  # 跳过第一个空部分
            lines = section.strip().split("\n")
            if not lines:
                continue
            
            slide_info = {
                "number": i,
                "title": "",
                "content": []
            }
            
            for line in lines:
                if line.startswith("幻灯片"):
                    continue
                elif line.startswith("标题:"):
                    slide_info["title"] = line.replace("标题:", "").strip()
                elif line.strip():
                    slide_info["content"].append(line.strip())
            
            slides.append(slide_info)
    except FileNotFoundError:
        print("未找到PPT提取文件")
        return []
    
    return slides

def analyze_images():
    """分析图片文件"""
    images_dir = Path("images")
    
    if not images_dir.exists():
        return []
    
    image_files = sorted([f for f in images_dir.iterdir() 
                         if f.is_file() and f.suffix.lower() in ['.jpg', '.jpeg', '.png', '.gif', '.svg', '.webp']])
    
    images = []
    for img_file in image_files:
        try:
            size = img_file.stat().st_size
            images.append({
                "name": img_file.name,
                "size_kb": size / 1024,
                "size_mb": size / (1024 * 1024)
            })
        except:
            pass
    
    # 按大小排序
    images.sort(key=lambda x: x["size_kb"], reverse=True)
    return images

def create_mapping():
    """创建PDF内容与图片的映射关系"""
    slides = analyze_pdf_slides()
    images = analyze_images()
    
    # 当前网页使用的图片
    current_usage = {
        "hero-bg.jpg": "首页背景 (幻灯片1)",
        "intro-excavator.jpg": "公司简介 (幻灯片2)",
        "tech-head.jpg": "混合头 (幻灯片10)",
        "tech-silo.jpg": "智能系统/料仓 (幻灯片11-12)",
        "test-cbr.jpg": "CBR测试 (幻灯片18)",
        "chart-comparison.jpg": "时间对比 (幻灯片21-23)",
        "project-wuhan.jpg": "武汉项目 (幻灯片26)",
        "project-hangzhou.jpg": "杭州项目 (幻灯片27)",
        "project-yuhuan.jpg": "玉环项目 (幻灯片28)"
    }
    
    print("=" * 80)
    print("PDF内容与图片映射分析")
    print("=" * 80)
    print()
    
    print(f"PDF幻灯片总数: {len(slides)}")
    print(f"图片文件总数: {len(images)}")
    print()
    
    print("=" * 80)
    print("当前网页使用的图片")
    print("=" * 80)
    for img_name, description in current_usage.items():
        img_path = Path("images") / img_name
        if img_path.exists():
            size = img_path.stat().st_size / 1024
            print(f"[OK] {img_name:30s} -> {description:40s} ({size:.1f} KB)")
        else:
            print(f"[X]  {img_name:30s} -> {description:40s} (缺失)")
    
    print()
    print("=" * 80)
    print("PDF幻灯片内容摘要")
    print("=" * 80)
    for slide in slides[:15]:  # 显示前15张
        print(f"\n幻灯片 {slide['number']}: {slide['title']}")
        if slide['content']:
            preview = " ".join(slide['content'][:2])[:100]
            print(f"  内容: {preview}...")
    
    print()
    print("=" * 80)
    print("大文件图片（可能是主要图片）")
    print("=" * 80)
    large_images = [img for img in images if img["size_kb"] > 500]
    for i, img in enumerate(large_images[:10], 1):
        size_str = f"{img['size_mb']:.1f} MB" if img['size_mb'] > 1 else f"{img['size_kb']:.1f} KB"
        print(f"{i:2d}. {img['name']:35s} ({size_str})")
    
    print()
    print("=" * 80)
    print("建议的图片映射关系")
    print("=" * 80)
    
    suggestions = {
        "幻灯片1 (首页)": ["image1.jpeg (11.6 MB) - 可能是更高质量的首页背景"],
        "幻灯片2 (公司简介)": ["intro-excavator.jpg (当前使用)", "检查是否有更清晰的挖掘机图片"],
        "幻灯片8 (设备)": ["image42.png (1.4 MB) - 可能是设备图", "检查其他设备相关图片"],
        "幻灯片9 (组件)": ["可能需要组件示意图"],
        "幻灯片10 (混合头)": ["tech-head.jpg (当前使用)", "确认是否正确"],
        "幻灯片11-12 (料仓/系统)": ["tech-silo.jpg (当前使用)", "可能需要单独的控制系统图片"],
        "幻灯片13 (结果)": ["可能需要结果展示图"],
        "幻灯片14-19 (测试)": ["test-cbr.jpg (当前使用，对应幻灯片18)", "可能需要其他测试方法图片"],
        "幻灯片21-23 (对比)": ["chart-comparison.jpg (当前使用)", "确认是否正确"],
        "幻灯片26 (武汉)": ["project-wuhan.jpg (当前使用)", "确认是否正确"],
        "幻灯片27 (杭州)": ["project-hangzhou.jpg (当前使用)", "确认是否正确"],
        "幻灯片28 (玉环)": ["project-yuhuan.jpg (当前使用)", "确认是否正确"]
    }
    
    for slide_desc, img_suggestions in suggestions.items():
        print(f"\n{slide_desc}:")
        for suggestion in img_suggestions:
            print(f"  - {suggestion}")
    
    print()
    print("=" * 80)
    print("未使用的图片文件（需要识别）")
    print("=" * 80)
    used_images = set(current_usage.keys())
    unused = [img for img in images if img['name'] not in used_images and img['size_kb'] > 100]
    
    print(f"\n找到 {len(unused)} 个未使用的中等/大文件图片:")
    for img in unused[:20]:  # 显示前20个
        size_str = f"{img['size_mb']:.1f} MB" if img['size_mb'] > 1 else f"{img['size_kb']:.1f} KB"
        print(f"  - {img['name']:35s} ({size_str})")
    
    print()
    print("=" * 80)
    print("操作建议")
    print("=" * 80)
    print("""
1. 打开大文件图片（>1MB）查看内容，确认对应关系
2. 根据PDF幻灯片内容，识别哪些图片应该用于哪些部分
3. 检查当前使用的图片是否正确
4. 考虑替换为更高质量的图片（如image1.jpeg可能是更好的首页背景）
5. 为新添加的设备、组件、材料部分找到合适的图片
6. 更新网页中的图片路径
    """)

if __name__ == "__main__":
    create_mapping()

