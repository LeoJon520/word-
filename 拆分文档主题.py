import os
from docx import Document
from docx.oxml.ns import qn

def split_word_by_headings(input_path, output_dir):
    """
    将Word文档按标题拆分为多个子文档
    
    Args:
        input_path: 输入Word文件路径
        output_dir: 输出目录路径
    """
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 读取原始文档
    doc = Document(input_path)
    
    # 用于存储标题和内容的列表
    sections = []  # [(title_text, [paragraphs]), ...]
    current_title = None
    current_content = []
    
    for paragraph in doc.paragraphs:
        # 检查是否为标题样式
        is_heading = False
        if paragraph.style.name.startswith('Heading') or paragraph.style.name.startswith('heading'):
            is_heading = True
        
        # 也可以通过段落格式判断是否为标题
        if not is_heading and paragraph.style.name == 'Normal':
            # 检查是否使用了标题相关的字体大小或加粗等特征
            pass
        
        if is_heading:
            # 保存上一个标题的内容
            if current_title is not None:
                sections.append((current_title, current_content))
            
            # 开始新的标题
            current_title = paragraph.text.strip()
            current_content = [paragraph]  # 包含标题本身
        else:
            if current_title is not None:
                current_content.append(paragraph)
            else:
                # 文档开头的非标题内容，作为无标题部分
                if current_title is None:
                    current_title = "开头内容"
                    current_content = [paragraph]
                else:
                    current_content.append(paragraph)
    
    # 保存最后一个标题的内容
    if current_title is not None and current_content:
        sections.append((current_title, current_content))
    
    # 如果没有找到任何标题，将整个文档作为一个文件
    if not sections:
        print("警告：未找到任何标题样式，将整个文档保存为一个文件")
        output_path = os.path.join(output_dir, f"完整文档.docx")
        doc.save(output_path)
        return
    
    # 创建子文档并保存
    file_count = 0
    for i, (title, paragraphs) in enumerate(sections, 1):
        if not title:  # 跳过空标题
            continue
            
        # 清理文件名中的非法字符
        safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_'))
        safe_title = safe_title.strip()[:50]  # 限制文件名长度
        
        if not safe_title:
            safe_title = f"未命名_{i}"
        
        # 创建新文档
        new_doc = Document()
        
        # 复制段落内容
        for para in paragraphs:
            # 获取段落文本和格式
            text = para.text
            style = para.style
            
            if text.strip():  # 只添加非空段落
                p = new_doc.add_paragraph(text, style=style)
                
                # 复制段落格式
                p.alignment = para.alignment
                p.paragraph_format.space_before = para.paragraph_format.space_before
                p.paragraph_format.space_after = para.paragraph_format.space_after
        
        # 生成输出文件名
        filename = f"{i:03d}_{safe_title}.docx"
        output_path = os.path.join(output_dir, filename)
        
        # 保存文档
        try:
            new_doc.save(output_path)
            file_count += 1
            print(f"已保存: {filename}")
        except Exception as e:
            print(f"保存失败 {filename}: {str(e)}")
    
    print(f"\n完成！共拆分出 {file_count} 个文档")

# 使用示例
if __name__ == "__main__":
    input_file = r"D:\筛重复\待编码的\大文档没主题分类的\有主题分类\模型_系统_理论_认知 核心_个体_.docx"
    output_dir = r"D:\筛重复\待编码的\大文档没主题分类的\有主题分类\try"
    
    # 首先安装依赖：pip install python-docx
    try:
        from docx import Document
        split_word_by_headings(input_file, output_dir)
    except ImportError:
        print("请先安装python-docx库：pip install python-docx")