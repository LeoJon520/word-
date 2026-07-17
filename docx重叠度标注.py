"""
DOCX文档内容重叠检测工具
功能：检测指定文件夹内所有docx文档之间的内容相似度，识别交叉重叠
使用前先安装依赖：pip install python-docx scikit-learn
"""

import os
import sys
from pathlib import Path
import re
from collections import defaultdict
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from typing import List, Dict, Tuple, Set
import zipfile
import json
import warnings
warnings.filterwarnings('ignore')

# 检查依赖
try:
    from docx import Document
    HAVE_DOCX_LIB = True
except ImportError:
    HAVE_DOCX_LIB = False
    print("⚠️ 警告：未安装 python-docx 库，将使用备用方法提取文本")

def extract_text_from_docx_advanced(docx_path: str) -> str:
    """
    从docx文件中提取文本（两种方法）
    方法1：使用 python-docx 库（推荐，能更好地处理格式）
    方法2：直接解析zip（备用，无需安装额外库）
    """
    docx_path = str(docx_path)
    
    if HAVE_DOCX_LIB:
        try:
            # 方法1：使用 python-docx
            doc = Document(docx_path)
            full_text = []
            
            # 提取段落
            for para in doc.paragraphs:
                if para.text.strip():
                    full_text.append(para.text.strip())
            
            # 提取表格中的文本
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if para.text.strip():
                                full_text.append(para.text.strip())
            
            return ' '.join(full_text)
        except Exception as e:
            print(f"⚠️ 使用python-docx提取失败 {Path(docx_path).name}: {e}")
            # 回退到方法2
    
    # 方法2：直接解析docx的zip结构
    try:
        with zipfile.ZipFile(docx_path) as z:
            # 读取document.xml（正文内容）
            if 'word/document.xml' in z.namelist():
                xml_content = z.read('word/document.xml').decode('utf-8', errors='ignore')
            else:
                # 尝试其他可能的路径
                for name in z.namelist():
                    if 'document.xml' in name:
                        xml_content = z.read(name).decode('utf-8', errors='ignore')
                        break
                else:
                    return ""
            
            # 清理XML标签，提取文本
            # 移除XML标签
            text = re.sub(r'<[^>]+>', ' ', xml_content)
            # 移除XML实体
            text = re.sub(r'&[a-z]+;', ' ', text)
            # 移除多余空格
            text = re.sub(r'\s+', ' ', text)
            
            return text.strip()
    except Exception as e:
        print(f"❌ 无法读取文件 {Path(docx_path).name}: {e}")
        return ""

def preprocess_text(text: str) -> str:
    """文本预处理：清洗、标准化"""
    if not text:
        return ""
    
    # 转换为小写
    text = text.lower()
    
    # 移除标点符号（保留中文字符）
    text = re.sub(r'[^\w\u4e00-\u9fff\s]', ' ', text)
    
    # 移除多余空白字符
    text = re.sub(r'\s+', ' ', text)
    
    # 移除数字（可选，根据需求调整）
    # text = re.sub(r'\d+', ' ', text)
    
    return text.strip()

def split_into_chunks(text: str, chunk_size: int = 500, overlap: int = 100) -> List[str]:
    """
    将文本分割成重叠的块，以便检测局部重叠
    chunk_size: 每个块的字数
    overlap: 块之间的重叠字数
    """
    if not text:
        return []
    
    chunks = []
    words = text.split()
    
    for i in range(0, len(words), chunk_size - overlap):
        chunk = ' '.join(words[i:i + chunk_size])
        if chunk:
            chunks.append(chunk)
        if i + chunk_size >= len(words):
            break
    
    return chunks

def calculate_document_similarity(doc_texts: Dict[str, str]) -> List[Tuple[str, str, float, List[Tuple[int, int, float]]]]:
    """
    计算文档间的相似度，并返回详细的相似信息
    返回：(文档1, 文档2, 整体相似度, [(块1索引, 块2索引, 块相似度), ...])
    """
    results = []
    doc_names = list(doc_texts.keys())
    doc_contents = list(doc_texts.values())
    
    # 为每个文档创建重叠的文本块
    all_chunks = []
    chunk_to_doc = []  # 记录每个块属于哪个文档
    
    for idx, (doc_name, text) in enumerate(zip(doc_names, doc_contents)):
        chunks = split_into_chunks(text, chunk_size=300, overlap=50)
        all_chunks.extend(chunks)
        chunk_to_doc.extend([(doc_name, idx, j) for j, _ in enumerate(chunks)])
    
    if len(all_chunks) < 2:
        return results
    
    # 使用TF-IDF向量化文本
    vectorizer = TfidfVectorizer(
        max_features=5000,
        stop_words=None,  # 中文需要另外处理停用词
        ngram_range=(1, 2)  # 考虑1-2个词的组合
    )
    
    try:
        tfidf_matrix = vectorizer.fit_transform(all_chunks)
        
        # 计算所有块之间的相似度
        chunk_similarities = cosine_similarity(tfidf_matrix)
        
        # 找出不同文档间高度相似的块
        n_chunks = len(all_chunks)
        
        for i in range(n_chunks):
            for j in range(i + 1, n_chunks):
                doc1_name, doc1_idx, chunk1_idx = chunk_to_doc[i]
                doc2_name, doc2_idx, chunk2_idx = chunk_to_doc[j]
                
                # 只关心不同文档间的相似度
                if doc1_name == doc2_name:
                    continue
                
                similarity = chunk_similarities[i, j]
                
                if similarity > 0.3:  # 相似度阈值
                    # 找到对应的文档对
                    found = False
                    for k, (d1, d2, overall_sim, chunk_pairs) in enumerate(results):
                        if (d1 == doc1_name and d2 == doc2_name) or (d1 == doc2_name and d2 == doc1_name):
                            # 添加块相似对
                            chunk_pairs.append((chunk1_idx, chunk2_idx, similarity))
                            # 更新整体相似度（取最大值）
                            if similarity > overall_sim:
                                results[k] = (d1, d2, similarity, chunk_pairs)
                            found = True
                            break
                    
                    if not found:
                        results.append((doc1_name, doc2_name, similarity, [(chunk1_idx, chunk2_idx, similarity)]))
    
    except Exception as e:
        print(f"❌ 计算相似度时出错: {e}")
    
    return results

def find_overlapping_content_pairs(folder_path: str, similarity_threshold: float = 0.4) -> Dict:
    """
    主函数：查找文件夹内docx文档的内容重叠
    """
    folder = Path(folder_path)
    if not folder.exists():
        print(f"❌ 文件夹不存在: {folder_path}")
        return {}
    
    # 收集所有docx文件
    docx_files = list(folder.glob("*.docx"))
    if not docx_files:
        print(f"⚠️ 文件夹中没有找到docx文件: {folder_path}")
        return {}
    
    print(f"📁 找到 {len(docx_files)} 个docx文件")
    
    # 提取文档内容
    doc_contents = {}
    failed_files = []
    
    for file_path in docx_files:
        print(f"  正在处理: {file_path.name}...")
        text = extract_text_from_docx_advanced(file_path)
        if text and len(text.strip()) > 50:  # 只保留有内容的文档
            processed_text = preprocess_text(text)
            if processed_text:
                doc_contents[file_path.name] = processed_text
            else:
                failed_files.append(file_path.name)
        else:
            failed_files.append(file_path.name)
    
    if failed_files:
        print(f"⚠️ 以下文件处理失败或内容过少: {', '.join(failed_files[:5])}")
        if len(failed_files) > 5:
            print(f"  等共 {len(failed_files)} 个文件")
    
    if len(doc_contents) < 2:
        print("❌ 至少需要2个有效文档才能进行对比")
        return {}
    
    print(f"✅ 成功处理 {len(doc_contents)} 个文档")
    
    # 计算相似度
    print("🔍 正在分析文档间的相似度...")
    similarities = calculate_document_similarity(doc_contents)
    
    # 过滤和排序结果
    filtered_results = []
    for doc1, doc2, overall_sim, chunk_pairs in similarities:
        if overall_sim >= similarity_threshold:
            # 计算平均块相似度和最大块相似度
            chunk_sims = [sim for _, _, sim in chunk_pairs]
            avg_chunk_sim = np.mean(chunk_sims) if chunk_sims else 0
            max_chunk_sim = max(chunk_sims) if chunk_sims else 0
            
            # 按最大相似度排序
            filtered_results.append({
                'doc1': doc1,
                'doc2': doc2,
                'overall_similarity': round(overall_sim, 3),
                'avg_chunk_similarity': round(avg_chunk_sim, 3),
                'max_chunk_similarity': round(max_chunk_sim, 3),
                'overlapping_chunks': len(chunk_pairs),
                'chunk_pairs': chunk_pairs[:10]  # 只保留前10个最相似的块对
            })
    
    # 按相似度排序
    filtered_results.sort(key=lambda x: x['max_chunk_similarity'], reverse=True)
    
    return {
        'total_documents': len(doc_contents),
        'similarity_threshold': similarity_threshold,
        'document_pairs': filtered_results,
        'all_documents': list(doc_contents.keys())
    }

def display_results(results: Dict, show_samples: bool = True):
    """以友好格式显示结果"""
    if not results or not results.get('document_pairs'):
        print("\n✅ 未发现明显的内容重叠文档")
        return
    
    print(f"\n{'='*60}")
    print(f"📊 内容重叠分析结果")
    print(f"{'='*60}")
    print(f"📁 分析文件夹: D:\\筛重复")
    print(f"📄 文档总数: {results['total_documents']}")
    print(f"🎯 相似度阈值: {results['similarity_threshold']}")
    print(f"🔍 发现 {len(results['document_pairs'])} 对重叠文档")
    print(f"{'='*60}\n")
    
    for i, pair in enumerate(results['document_pairs'], 1):
        print(f"{i}. 📂 文档对: {pair['doc1']} ↔ {pair['doc2']}")
        print(f"   📈 整体相似度: {pair['overall_similarity']:.1%}")
        print(f"   📊 最大块相似度: {pair['max_chunk_similarity']:.1%}")
        print(f"   📊 平均块相似度: {pair['avg_chunk_similarity']:.1%}")
        print(f"   🔢 重叠块数量: {pair['overlapping_chunks']}")
        
        if show_samples and pair['chunk_pairs']:
            print(f"   📋 高相似度内容块:")
            for chunk_idx, (chunk1_idx, chunk2_idx, sim) in enumerate(pair['chunk_pairs'][:3], 1):
                print(f"     块{chunk_idx}: 相似度 {sim:.1%}")
        print()

def save_results_to_file(results: Dict, output_file: str = "docx_overlap_report.json"):
    """将结果保存到JSON文件"""
    with open(output_file, 'w', encoding='utf-8') as f:
        # 转换numpy类型为Python原生类型以便JSON序列化
        json_serializable = {}
        for key, value in results.items():
            if key == 'document_pairs':
                json_serializable[key] = []
                for pair in value:
                    pair_copy = pair.copy()
                    # 处理chunk_pairs中的元组
                    if 'chunk_pairs' in pair_copy:
                        pair_copy['chunk_pairs'] = [
                            (int(a), int(b), float(c)) for a, b, c in pair_copy['chunk_pairs']
                        ]
                    json_serializable[key].append(pair_copy)
            else:
                json_serializable[key] = value
        
        json.dump(json_serializable, f, ensure_ascii=False, indent=2)
    
    print(f"💾 详细结果已保存到: {output_file}")

def main():
    """主函数"""
    print("="*60)
    print("🔍 DOCX文档内容重叠检测工具")
    print("="*60)
    
    # 设置文件夹路径
    folder_path = r"D:\筛重复"
    
    # 检查依赖
    if not HAVE_DOCX_LIB:
        print("⚠️ 未检测到 python-docx 库，将使用备用文本提取方法")
        print("   如需更好效果，请安装: pip install python-docx scikit-learn")
        print("   继续使用基础模式...\n")
    
    # 设置相似度阈值（0-1之间，建议0.3-0.6）
    similarity_threshold = 0.4
    
    try:
        # 分析文档
        results = find_overlapping_content_pairs(folder_path, similarity_threshold)
        
        if results:
            # 显示结果
            display_results(results, show_samples=True)
            
            # 保存详细结果
            save_results_to_file(results)
            
            # 输出摘要
            print("\n📋 检测摘要:")
            print(f"   文档总数: {results['total_documents']}")
            print(f"   发现重叠文档对: {len(results['document_pairs'])}")
            
            if results['document_pairs']:
                print("\n🔴 建议重点检查的文档:")
                seen_docs = set()
                for pair in results['document_pairs'][:5]:  # 显示前5对
                    if pair['max_chunk_similarity'] > 0.7:  # 高度相似
                        if pair['doc1'] not in seen_docs:
                            print(f"   • {pair['doc1']}")
                            seen_docs.add(pair['doc1'])
                        if pair['doc2'] not in seen_docs:
                            print(f"   • {pair['doc2']}")
                            seen_docs.add(pair['doc2'])
        else:
            print("❌ 未获取到有效分析结果")
    
    except Exception as e:
        print(f"❌ 程序执行出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # 检查并尝试安装必要库
    try:
        import sklearn
    except ImportError:
        print("❌ 缺少必要库: scikit-learn")
        print("   请运行: pip install scikit-learn")
        sys.exit(1)
    
    main()
