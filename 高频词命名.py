import os
import re
from collections import Counter
import jieba
from docx import Document
from pathlib import Path


def extract_text_from_docx(docx_path):
    """从docx文件中提取文本内容"""
    try:
        doc = Document(docx_path)
        full_text = []
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                full_text.append(paragraph.text)
        return '\n'.join(full_text)
    except Exception as e:
        print(f"读取文件 {docx_path} 时出错: {e}")
        return ""


def clean_text(text):
    """清洗文本，移除标点符号、数字和英文"""
    # 移除标点符号
    text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', ' ', text)
    # 移除数字
    text = re.sub(r'\d+', ' ', text)
    # 移除英文字母
    text = re.sub(r'[a-zA-Z]', ' ', text)
    # 移除多余空格
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def get_top_keywords(text, top_n=10):
    """获取文本中的高频关键词"""
    if not text:
        return []

    # 清洗文本
    cleaned_text = clean_text(text)

    # 使用jieba分词
    words = jieba.lcut(cleaned_text)

    # 过滤单个字符和停用词
    stop_words = {'的', '了', '在', '是', '我', '有', '和', '就',
                  '不', '人', '都', '一', '一个', '上', '也', '很',
                  '到', '说', '要', '去', '你', '会', '着', '没有',
                  '看', '好', '自己', '这'}

    filtered_words = [word for word in words
                      if len(word) > 1 and word not in stop_words]

    # 统计词频
    word_counts = Counter(filtered_words)

    # 获取前N个高频词
    top_words = [word for word, count in word_counts.most_common(top_n)]

    return top_words


def rename_docx_files(folder_path):
    """重命名文件夹中的docx文件"""
    folder = Path(folder_path)

    if not folder.exists():
        print(f"文件夹 {folder_path} 不存在！")
        return

    # 获取所有docx文件
    docx_files = list(folder.glob("*.docx"))

    if not docx_files:
        print("文件夹中没有找到docx文件！")
        return

    print(f"找到 {len(docx_files)} 个docx文件")

    renamed_count = 0
    skipped_count = 0

    for docx_file in docx_files:
        print(f"\n处理文件: {docx_file.name}")

        # 提取文本
        text = extract_text_from_docx(docx_file)

        if not text:
            print(f"  ⚠️ 文件内容为空，跳过")
            skipped_count += 1
            continue

        # 获取高频词
        top_keywords = get_top_keywords(text)

        if not top_keywords:
            print(f"  ⚠️ 无法提取关键词，跳过")
            skipped_count += 1
            continue

        # 生成新文件名（用下划线连接前10个高频词，限制总长度）
        new_name = "_".join(top_keywords)

        # 限制文件名长度（避免文件名过长）
        if len(new_name) > 100:
            new_name = new_name[:100]

        # 添加文件扩展名
        new_filename = f"{new_name}.docx"
        new_path = folder / new_filename

        # 如果文件名已存在，添加数字后缀
        counter = 1
        while new_path.exists():
            new_filename = f"{new_name}_{counter}.docx"
            new_path = folder / new_filename
            counter += 1

        # 重命名文件
        try:
            docx_file.rename(new_path)
            print(f"  ✅ 重命名为: {new_filename}")
            print(f"  高频词: {', '.join(top_keywords)}")
            renamed_count += 1
        except Exception as e:
            print(f"  ❌ 重命名失败: {e}")
            skipped_count += 1

    print(f"\n{'='*50}")
    print(f"处理完成！")
    print(f"成功重命名: {renamed_count} 个文件")
    print(f"跳过: {skipped_count} 个文件")


def main():
    # 设置文件夹路径
    folder_path = r"D:\筛重复\内容重复特别大"

    # 确认操作
    print("="*50)
    print("文档重命名工具")
    print("="*50)
    print(f"将处理文件夹: {folder_path}")
    print("每个文档将根据其内容的前10个高频词重命名")
    print("\n请确认:")
    print("1. 确保文件夹路径正确")
    print("2. 建议先备份原始文件")
    print("3. 重命名操作不可逆")

    confirm = input("\n是否继续? (输入 'yes' 继续): ").strip().lower()

    if confirm == 'yes':
        # 执行重命名
        rename_docx_files(folder_path)

        print("\n操作完成！")
        print("注意事项:")
        print("1. 高频词提取基于jieba分词")
        print("2. 已过滤常见停用词")
        print("3. 文件名过长时会被截断")
    else:
        print("操作已取消")


if __name__ == "__main__":
    # 安装必要的库（如果尚未安装）
    print("检查依赖库...")
    try:
        from docx import Document
    except ImportError:
        print("请先安装必要的库:")
        print("pip install python-docx jieba")
        exit(1)

    main()
