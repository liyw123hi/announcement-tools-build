#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
建行公告附件验证工具 - 【PDF公告名称校验版】
只校验PDF附件名称是否与Excel标题一致，不校验落款日期
Excel标题截断规则：取"2025年四季度暨年度投资管理报告"之前（含）的部分
"""

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import sys
import shutil
from datetime import datetime
import logging
import re
import traceback
from typing import Dict, Any, List, Optional, Tuple
import glob
import tempfile
import platform

# ==================== 配置日志记录器 ====================
def setup_logger():
    """设置详细的状态日志记录器"""
    logger = logging.getLogger("公告验证")
    logger.setLevel(logging.INFO)
    
    if not logger.hasHandlers():
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_format = logging.Formatter(
            '[%(asctime)s] [%(levelname)s] %(message)s',
            datefmt='%H:%M:%S'
        )
        console_handler.setFormatter(console_format)
        logger.addHandler(console_handler)
    
    return logger

logger = setup_logger()

# ==================== 清理函数 ====================
def clean_download_folder(download_path: str):
    """静默清理文件夹"""
    if os.path.exists(download_path):
        try:
            for filename in os.listdir(download_path):
                file_path = os.path.join(download_path, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    pass
        except Exception as e:
            pass

# ==================== 辅助函数 ====================
def clean_text(text: str) -> str:
    """清理文本，移除多余空格和换行符"""
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

TITLE_CUTOFF_KEYWORD = "2025年四季度暨年度投资管理报告"

def truncate_excel_title(title: str) -> str:
    """
    截断 Excel 标题：取 TITLE_CUTOFF_KEYWORD（含）之前的部分。
    若标题中不含该关键词，则原样返回。
    """
    if not title:
        return title
    idx = title.find(TITLE_CUTOFF_KEYWORD)
    if idx == -1:
        return title.strip()
    return title[:idx + len(TITLE_CUTOFF_KEYWORD)].strip()

def is_title_match(expected_title: str, actual_title: str) -> bool:
    """严格检查标题是否匹配"""
    if not expected_title or not actual_title:
        return False
    
    expected_clean = clean_text(expected_title)
    actual_clean = clean_text(actual_title)
    
    # 完全一致
    if expected_clean == actual_clean:
        return True
    
    # 包含关系
    if expected_clean in actual_clean or actual_clean in expected_clean:
        return True
    
    # 移除标点符号和空格后比较
    expected_simple = re.sub(r'[^\w\u4e00-\u9fff]', '', expected_clean)
    actual_simple = re.sub(r'[^\w\u4e00-\u9fff]', '', actual_clean)
    
    if expected_simple and actual_simple:
        if expected_simple in actual_simple or actual_simple in expected_simple:
            return True
    
    return False

def contains_target_date(url: str, target_date: str) -> bool:
    """检查URL中是否包含目标日期"""
    if not url or not target_date:
        return False
    
    date_formats = [
        target_date.replace('-', ''),  # 20260316
        target_date,                   # 2026-03-16
        target_date.replace('-', '/'),  # 2026/03/16
    ]
    
    for fmt in date_formats:
        if fmt in url:
            return True
    
    year, month, day = target_date.split('-')
    patterns = [
        rf'{year}[\/\-]?{month}[\/\-]?{day}',
        rf'{year}{month}{day}',
    ]
    
    for pattern in patterns:
        if re.search(pattern, url):
            return True
    
    return False

def extract_date_from_text(text: str, target_date: str) -> Optional[str]:
    """从文本中提取并验证日期 - 修复返回类型"""
    if not text or not target_date:
        return None
    
    # 目标日期格式
    target_year, target_month, target_day = target_date.split('-')
    
    # 支持的日期格式模式
    date_patterns = [
        (r'(\d{4})年(\d{1,2})月(\d{1,2})日', 1, 2, 3),  # 2026年03月16日
        (r'(\d{4})年(\d{1,2})月(\d{1,2})号', 1, 2, 3),  # 2026年03月16号
        (r'(\d{4})-(\d{1,2})-(\d{1,2})', 1, 2, 3),      # 2026-03-16
        (r'(\d{4})/(\d{1,2})/(\d{1,2})', 1, 2, 3),      # 2026/03/16
    ]
    
    for pattern, year_idx, month_idx, day_idx in date_patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            if len(match) >= 3:
                year = match[year_idx-1]
                month = match[month_idx-1].zfill(2)
                day = match[day_idx-1].zfill(2)
                
                # 检查是否匹配目标日期
                if year == target_year and month == target_month and day == target_day:
                    return f"{year}-{month}-{day}"
    
    return None

def read_docx_file(file_path: str) -> str:
    """
    读取.docx文件内容
    注意：需要安装python-docx库：pip install python-docx
    """
    try:
        # 尝试导入python-docx
        try:
            from docx import Document
        except ImportError:
            logger.error("未安装python-docx库，无法读取.docx文件")
            logger.info("请运行: pip install python-docx")
            return ""
        
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        
        return '\n'.join(full_text)
    except Exception as e:
        logger.error(f"读取.docx文件失败: {str(e)}")
        return ""

def read_text_file(file_path: str) -> str:
    """
    读取文本文件内容（跨平台、多编码支持）
    """
    if not os.path.exists(file_path):
        logger.warning(f"文件不存在: {file_path}")
        return ""
    
    try:
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            logger.warning(f"文件为空: {file_path}")
            return ""
    except Exception as e:
        logger.warning(f"无法读取文件大小: {str(e)}")
        return ""
    
    # 尝试多种编码
    encodings = [
        'utf-8',
        'utf-8-sig',
        'gbk',
        'gb2312',
        'big5',
        'latin-1',
        'cp1252',
    ]
    
    file_content = ""
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                file_content = f.read()
            
            # 检查是否成功读取了内容
            if file_content and len(file_content.strip()) > 0:
                logger.debug(f"使用编码 {encoding} 成功读取文件，内容长度: {len(file_content)}")
                return file_content
            
        except Exception as e:
            logger.debug(f"尝试编码 {encoding} 失败: {str(e)}")
            continue
    
    # 如果所有文本编码都失败，尝试二进制读取然后解码
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read()
        
        # 尝试检测并解码
        for encoding in encodings:
            try:
                file_content = raw_data.decode(encoding, errors='ignore')
                if file_content and len(file_content.strip()) > 0:
                    logger.debug(f"使用二进制模式和编码 {encoding} 成功读取文件")
                    return file_content
            except:
                continue
    except Exception as e:
        logger.warning(f"二进制读取失败: {str(e)}")
    
    logger.warning(f"无法使用任何编码读取文件: {file_path}")
    return ""

def read_pdf_file(file_path: str) -> str:
    """
    读取PDF文件内容
    注意：需要安装PyPDF2库：pip install PyPDF2
    """
    try:
        # 尝试导入PyPDF2
        try:
            import PyPDF2
        except ImportError:
            logger.error("未安装PyPDF2库，无法读取.pdf文件")
            logger.info("请运行: pip install PyPDF2")
            return ""
        
        text = ""
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text += page.extract_text()
        
        return text
    except Exception as e:
        logger.error(f"读取PDF文件失败: {str(e)}")
        return ""

def get_most_recent_file(download_path: str, extension: str = None) -> Optional[str]:
    """获取下载目录中最新的文件"""
    if not os.path.exists(download_path):
        return None
    
    files = []
    for filename in os.listdir(download_path):
        file_path = os.path.join(download_path, filename)
        if os.path.isfile(file_path):
            # 过滤临时文件
            if not filename.endswith(('.tmp', '.crdownload', '.part')):
                if extension is None or filename.lower().endswith(extension.lower()):
                    files.append((file_path, os.path.getmtime(file_path)))
    
    if not files:
        return None
    
    # 按修改时间排序，返回最新的文件
    files.sort(key=lambda x: x[1], reverse=True)
    return files[0][0]

def get_downloaded_files_since(download_path: str, start_time: float) -> List[str]:
    """获取指定时间后修改的所有文件（跨平台兼容）"""
    downloaded_files = []
    
    try:
        for filename in os.listdir(download_path):
            # 跳过系统文件和临时文件
            if filename.startswith('.'):
                continue
            if filename.endswith(('.tmp', '.crdownload', '.part', '.downloading')):
                continue
            if filename.endswith('.DS_Store'):  # macOS 系统文件
                continue
            
            file_path = os.path.join(download_path, filename)
            
            # 确保是文件而不是文件夹
            if not os.path.isfile(file_path):
                continue
            
            try:
                # 获取修改时间
                mtime = os.path.getmtime(file_path)
                
                # 只获取在 start_time 之后修改的文件
                if mtime >= start_time:
                    downloaded_files.append(file_path)
            except Exception as e:
                logger.debug(f"无法读取文件 {file_path} 的修改时间: {str(e)}")
                continue
    
    except Exception as e:
        logger.error(f"获取已下载文件列表失败: {str(e)}")
    
    return downloaded_files


def wait_for_file_download(download_path: str, max_wait_time: int = 30, 
                          initial_wait: float = 2.0) -> Optional[str]:
    """
    等待文件下载完成（跨平台优化版）
    返回：下载的文件路径，如果超时返回 None
    """
    # 获取初始文件列表（用于识别新下载的文件）
    initial_files = set()
    try:
        for filename in os.listdir(download_path):
            if not filename.startswith('.') and filename != '.DS_Store':
                initial_files.add(os.path.join(download_path, filename))
    except:
        pass
    
    # 初始等待，让浏览器开始下载
    time.sleep(initial_wait)
    
    start_time = time.time()
    last_valid_file = None
    last_valid_size = 0
    stable_count = 0
    
    while time.time() - start_time < max_wait_time:
        try:
            # 获取下载目录中的所有文件
            current_files = []
            for filename in os.listdir(download_path):
                # 跳过系统文件和临时文件
                if filename.startswith('.'):
                    continue
                if filename.endswith(('.tmp', '.crdownload', '.part', '.downloading')):
                    continue
                
                file_path = os.path.join(download_path, filename)
                if os.path.isfile(file_path):
                    current_files.append(file_path)
            
            # 查找新文件（不在初始列表中的文件）
            new_files = [f for f in current_files if f not in initial_files]
            
            if new_files:
                # 获取最新修改的文件
                new_files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
                current_file = new_files[0]
                
                # 检查文件大小是否稳定
                try:
                    current_size = os.path.getsize(current_file)
                    
                    if current_size > 0:
                        # 如果大小相同且稳定了多次，认为下载完成
                        if current_size == last_valid_size:
                            stable_count += 1
                            if stable_count >= 2:  # 连续 2 次检查大小相同即认为完成
                                logger.info(f"文件下载完成: {os.path.basename(current_file)} ({current_size} 字节)")
                                return current_file
                        else:
                            last_valid_file = current_file
                            last_valid_size = current_size
                            stable_count = 0
                            logger.debug(f"检测到新文件: {os.path.basename(current_file)} ({current_size} 字节)")
                except OSError as e:
                    logger.debug(f"无法读取文件大小: {str(e)}")
                    stable_count = 0
            
            time.sleep(0.8)  # 减少检查间隔以提高响应性
            
        except Exception as e:
            logger.debug(f"等待下载时发生错误: {str(e)}")
            time.sleep(1)
    
    # 超时后，如果有找到文件，尝试最后一次验证
    if last_valid_file and os.path.exists(last_valid_file):
        try:
            final_size = os.path.getsize(last_valid_file)
            if final_size > 0:
                logger.info(f"文件下载（超时前）: {os.path.basename(last_valid_file)} ({final_size} 字节)")
                return last_valid_file
        except:
            pass
    
    return None


def download_and_read_file(driver, attachment_url: str, download_path: str) -> Tuple[str, str]:
    """
    下载文件并读取内容
    返回：(文件内容, 文件路径)
    """
    # 清理下载目录
    clean_download_folder(download_path)
    
    # 记录当前窗口
    original_window = driver.current_window_handle
    
    try:
        # 通过JavaScript在新标签页中打开链接（触发下载）
        logger.debug(f"正在下载: {attachment_url[:80]}...")
        driver.execute_script(f"window.open('{attachment_url}');")
        
        # 等待文件下载完成
        downloaded_file = wait_for_file_download(download_path, max_wait_time=30, initial_wait=2.0)
        
        if not downloaded_file or not os.path.exists(downloaded_file):
            logger.warning(f"未找到下载的文件或文件不存在")
            return "", ""
        
        # 最后验证文件大小
        try:
            file_size = os.path.getsize(downloaded_file)
            if file_size == 0:
                logger.warning(f"文件大小为 0: {os.path.basename(downloaded_file)}")
                return "", ""
        except Exception as e:
            logger.warning(f"无法读取文件大小: {str(e)}")
            return "", ""
        
        # 读取文件内容
        file_ext = os.path.splitext(downloaded_file)[1].lower()
        file_content = ""
        
        logger.debug(f"识别文件类型: {file_ext}")
        
        if file_ext == '.docx':
            file_content = read_docx_file(downloaded_file)
        elif file_ext == '.pdf':
            file_content = read_pdf_file(downloaded_file)
        elif file_ext in ['.txt', '.html', '.htm', '.xlsx', '.xls']:
            file_content = read_text_file(downloaded_file)
        elif file_ext in ['.doc']:
            # .doc 文件较难处理，标记为需要手动检查
            file_content = f"[.doc文件，需要手动检查内容] 文件路径: {downloaded_file}"
        else:
            # 尝试作为文本文件读取
            try:
                file_content = read_text_file(downloaded_file)
            except Exception as e:
                logger.debug(f"无法作为文本文件读取: {str(e)}")
                file_content = f"[二进制文件类型: {file_ext}，文件路径: {downloaded_file}]"
        
        # 验证是否成功读取内容
        if not file_content or (isinstance(file_content, str) and len(file_content.strip()) == 0):
            logger.warning(f"文件内容为空: {os.path.basename(downloaded_file)}")
            return "", downloaded_file  # 返回文件路径但内容为空
        
        logger.info(f"✓ 文件读取成功，内容长度: {len(file_content)} 字符")
        return file_content, downloaded_file
        
    except Exception as e:
        logger.error(f"下载文件时发生异常: {str(e)}")
        import traceback
        logger.debug(f"详细错误: {traceback.format_exc()}")
        return "", ""
    
    finally:
        # 确保切换回原始窗口
        try:
            if len(driver.window_handles) > 1:
                # 关闭新打开的标签页
                for handle in driver.window_handles:
                    if handle != original_window:
                        try:
                            driver.switch_to.window(handle)
                            driver.close()
                        except:
                            pass
                driver.switch_to.window(original_window)
        except:
            pass

# ==================== 核心验证函数 ====================
def validate_announcement_title(driver, expected_title: str) -> Tuple[bool, str]:
    """严格验证公告标题是否匹配"""
    if not expected_title:
        return True, "无预期标题，跳过验证"
    
    # 1. 检查页面标题
    page_title = driver.title.strip()
    if page_title and is_title_match(expected_title, page_title):
        return True, f"页面标题匹配: {page_title[:50]}..."
    
    # 2. 查找页面中的标题元素
    title_selectors = [
        "h1", "h2", "h3", 
        ".title", ".tit", ".news-title", ".article-title",
        "div.title", "div.tit", ".content-title", ".news-detail-title"
    ]
    
    for selector in title_selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            for element in elements:
                element_text = element.text.strip()
                if element_text and is_title_match(expected_title, element_text):
                    return True, f"找到匹配标题: {element_text[:50]}..."
        except:
            continue
    
    # 3. 查找正文中的公告标题
    try:
        announcement_elements = driver.find_elements(By.XPATH, 
            "//*[contains(text(), '公告') or contains(text(), '通知')]")
        for element in announcement_elements:
            element_text = element.text.strip()
            if len(element_text) > 20 and is_title_match(expected_title, element_text):
                return True, f"在正文中找到匹配: {element_text[:50]}..."
    except:
        pass
    
    return False, f"未找到与'{expected_title[:50]}...'匹配的标题"

def verify_attachment_by_name(attachment_text: str, expected_title: str) -> Tuple[str, str]:
    """
    验证附件名称 - 附件名（除扩展名外）必须与截断后的 Excel 标题完全一致
    返回：(状态, 结果描述)
    """
    
    try:
        logger.info(f"开始验证附件名称: {attachment_text[:50]}...")
        
        if not attachment_text:
            return "失败", "附件名称为空"
        
        if not expected_title:
            return "失败", "预期标题为空"
        
        # 直接使用完整预期标题（不再截断）
        expected = expected_title.strip()
        
        # 清理附件文本
        att_name = clean_text(attachment_text)
        
        # 移除文件扩展名（含大小写）
        for ext in ['.pdf', '.PDF', '.docx', '.DOCX', '.doc', '.DOC',
                    '.xlsx', '.XLSX', '.xls', '.XLS', '.rar', '.RAR', '.zip', '.ZIP']:
            if att_name.endswith(ext):
                att_name = att_name[:-len(ext)]
                break
        
        att_name = att_name.strip()
        
        # 完全一致
        if att_name == expected:
            return "成功", f"附件名称与标题完全一致: {expected}"
        
        # 大小写不敏感
        if att_name.upper() == expected.upper():
            return "成功", f"附件名称与标题一致（忽略大小写）: {expected}"
        
        # 完全不匹配
        return "失败", f"附件名称 '{att_name}' 与预期标题 '{expected}' 不一致"
        
    except Exception as e:
        logger.error(f"验证附件名称时发生异常: {str(e)}")
        return "异常", f"验证异常: {str(e)[:50]}"

def extract_pdf_title(file_content: str) -> str:
    """
    从PDF文本中提取第一个有意义的标题行。
    策略：取非空行中第一个长度在6~80字符之间、包含中文的行。
    """
    if not file_content:
        return ""
    lines = file_content.splitlines()
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # 跳过明显的页眉页脚噪声（纯数字、纯英文短串等）
        if len(line) < 6 or len(line) > 120:
            continue
        # 必须包含中文
        if not re.search(r'[\u4e00-\u9fff]', line):
            continue
        return line
    return ""


def verify_pdf_content_title(driver, attachment_url: str,
                              expected_title: str, download_path: str) -> Tuple[str, str]:
    """
    下载PDF附件，读取其文本内容，提取第一个有意义的标题行，
    与截断后的 Excel 标题比对。
    返回：(状态, 结果描述)
    """
    expected = expected_title.strip()

    try:
        file_content, downloaded_file = download_and_read_file(driver, attachment_url, download_path)

        if not downloaded_file or not os.path.exists(downloaded_file):
            return "失败", "PDF文件未下载成功"

        if not file_content or len(file_content.strip()) == 0:
            # 尝试再读一次（有时下载后需要稍等）
            time.sleep(1)
            file_content = read_pdf_file(downloaded_file)

        # 清理临时文件
        try:
            if os.path.exists(downloaded_file):
                os.remove(downloaded_file)
        except:
            pass

        if not file_content or len(file_content.strip()) == 0:
            return "失败", "PDF内容为空，无法读取标题"

        pdf_title = extract_pdf_title(file_content)

        if not pdf_title:
            return "失败", "未能从PDF中提取到标题行"

        logger.info(f"    │  PDF提取标题: '{pdf_title[:60]}'")
        logger.info(f"    │  预期标题:     '{expected[:60]}'")

        # 完全一致
        if pdf_title == expected:
            return "成功", f"PDF标题完全一致: {pdf_title[:60]}"

        # 忽略大小写
        if pdf_title.upper() == expected.upper():
            return "成功", f"PDF标题一致（忽略大小写）: {pdf_title[:60]}"

        # 互相包含（处理PDF标题可能带换行拼接的情况）
        pdf_simple = re.sub(r'\s+', '', pdf_title)
        exp_simple = re.sub(r'\s+', '', expected)
        if pdf_simple == exp_simple:
            return "成功", f"PDF标题一致（忽略空白）: {pdf_title[:60]}"
        if exp_simple and pdf_simple and (exp_simple in pdf_simple or pdf_simple in exp_simple):
            return "成功", f"PDF标题包含预期标题: {pdf_title[:60]}"

        return "失败", f"PDF标题 '{pdf_title[:50]}' 与预期 '{expected[:50]}' 不一致"

    except Exception as e:
        logger.error(f"验证PDF内容标题时发生异常: {str(e)}")
        return "异常", f"PDF标题验证异常: {str(e)[:50]}"


def verify_announcement_date(driver, announcement_url: str, expected_date: str) -> Tuple[str, str]:
    """
    验证公告发布日期是否与预期日期一致。
    优先从 URL 中提取日期（如 newsdetail/20260428_xxx.html → 2026-04-28），
    若 URL 中无日期，则从页面内容中提取。
    返回：(状态, 结果描述)
    """
    if not expected_date:
        return "跳过", "未设定预期日期"
    
    try:
        # 1. 从 URL 中提取日期
        url_date = None
        # 匹配 newsdetail/YYYYMMDD_xxx.html 或 newsdetail/YYYY-MM-DD_xxx.html
        url_patterns = [
            r'newsdetail/(\d{4})(\d{2})(\d{2})_',  # 20260428
            r'newsdetail/(\d{4})-(\d{2})-(\d{2})_',  # 2026-04-28
            r'/(\d{4})(\d{2})(\d{2})/',  # /20260428/
        ]
        for pattern in url_patterns:
            m = re.search(pattern, announcement_url)
            if m:
                url_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                break
        
        # 2. 若 URL 无日期，从页面内容中提取
        page_date = None
        if not url_date:
            page_text = driver.execute_script("return document.body.innerText || '';") or ''
            # 优先找靠近标题或发布时间的日期
            date_patterns = [
                r'(\d{4})年(\d{1,2})月(\d{1,2})日',
                r'(\d{4})-(\d{1,2})-(\d{1,2})',
                r'(\d{4})/(\d{1,2})/(\d{1,2})',
            ]
            for pattern in date_patterns:
                matches = re.findall(pattern, page_text)
                for match in matches:
                    if len(match) >= 3:
                        try:
                            y, m, d = match[0], match[1].zfill(2), match[2].zfill(2)
                            # 验证是否为合法日期
                            dt(int(y), int(m), int(d))
                            page_date = f"{y}-{m}-{d}"
                            break
                        except:
                            continue
                if page_date:
                    break
        
        actual_date = url_date or page_date
        
        if not actual_date:
            return "失败", "无法从URL或页面中提取到公告日期"
        
        # 标准化比较
        expected_clean = expected_date.replace('/', '-').strip()
        actual_clean = actual_date.replace('/', '-').strip()
        
        if actual_clean == expected_clean:
            source = "URL" if url_date else "页面内容"
            return "成功", f"公告日期匹配 ({source}): {actual_date}"
        else:
            return "失败", f"公告日期不匹配: 预期 {expected_clean}, 实际 {actual_clean}"
            
    except Exception as e:
        logger.error(f"验证公告日期时发生异常: {str(e)}")
        return "异常", f"日期验证异常: {str(e)[:50]}"


def find_search_button(driver):
    """查找搜索按钮的辅助函数"""
    # 尝试多种选择器
    selectors = [
        # CSS选择器
        "input.but[type='button'][value='搜索']",
        "input[type='button'][value='搜索']",
        "input[type='submit'][value='搜索']",
        "button[type='submit']",
        "input[value='搜索']",
        ".but[type='button']",
        ".but[type='submit']",
        
        # 更通用的选择器
        "input[type='button']",
        "input[type='submit']",
        "button",
        
        # 通过文本查找
        "//input[@value='搜索']",
        "//button[contains(text(), '搜索')]",
        "//input[@type='button' and contains(@value, '搜索')]",
        "//input[@type='submit' and contains(@value, '搜索')]",
    ]
    
    for selector in selectors:
        try:
            if selector.startswith("//"):
                # XPath选择器
                elements = driver.find_elements(By.XPATH, selector)
            else:
                # CSS选择器
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
            
            for element in elements:
                try:
                    # 检查元素是否可见和可点击
                    if element.is_displayed() and element.is_enabled():
                        # 检查元素文本或值是否包含"搜索"
                        element_text = element.text or element.get_attribute('value') or ''
                        if '搜索' in element_text:
                            return element
                except:
                    continue
        except:
            continue
    
    return None

def search_and_verify(driver, announcement_name: str, 
                      announcement_url: str, announcement_date: str, 
                      download_path: str, index: int, total: int, 
                      processed_count: int) -> Dict[str, Any]:
    """完整的验证逻辑（6个步骤）
    用公告名称在建行官网搜索，找到对应公告后校验。
    校验条件（四项全过才算成功）：
      1. 公告发布日期 == config中设定的日期
      2. 公告详情名称 == Excel标题截断后的值
      3. PDF附件名（去扩展名）== Excel标题截断后的值
      4. PDF内部第一个标题行 == Excel标题截断后的值
    """
    
    announcement_display = (f"'{announcement_name[:30]}...'"
                            if announcement_name and len(announcement_name) > 30
                            else f"'{announcement_name}'")
    logger.info(f"[{index+1}/{total}] 正在检验: {announcement_display}")
    
    # 直接使用完整公告名称作为预期标题（不再截断）
    expected_title = announcement_name if announcement_name else ''
    
    log_entry = {
        '公告名称':      str(announcement_name) if announcement_name else '',
        '公告URL':       str(announcement_url)  if announcement_url  else '',
        '验证时间':      datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        '验证状态':      '进行中',
        '失败原因':      '',
        '附件数量':      0,
        '附件状态':      '',
        '日期验证':      '',
        '公告标题验证':  '',
        '附件名称验证':  '',
        'PDF内容标题验证': '',
    }
    
    try:
        # 步骤1: 打开建行官网
        logger.info(f"  ├─ 步骤1/6: 正在打开建行官网...")
        driver.get("https://finance1.ccb.com/chn/finance/yfgg.shtml")
        time.sleep(2)
        
        # 步骤2: 输入公告名称搜索
        search_keyword = announcement_name if announcement_name else ''
        logger.info(f"  ├─ 步骤2/6: 正在输入公告名称搜索（关键词: '{search_keyword}'）...")
        try:
            search_box_selectors = [
                "input.kuang[type='text']",
                "input[type='text']",
                "input[name='keyword']",
                "input[name='search']",
                "input[placeholder*='搜索']",
                "input[placeholder*='请输入']",
                "input.search-input",
                "#keyword",
                "#search",
            ]
            search_box = None
            for selector in search_box_selectors:
                try:
                    search_box = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    if search_box:
                        break
                except:
                    continue
            if not search_box:
                for input_elem in driver.find_elements(By.TAG_NAME, "input"):
                    if input_elem.get_attribute('type') in ['text', 'search']:
                        search_box = input_elem
                        break
            if not search_box:
                log_entry['验证状态'] = '失败'
                log_entry['失败原因'] = '找不到搜索框'
                logger.warning(f"  └─ 结果: 失败 - {log_entry['失败原因']}")
                return log_entry
            search_box.clear()
            search_box.send_keys(search_keyword)
            time.sleep(1)
        except Exception as e:
            log_entry['验证状态'] = '失败'
            log_entry['失败原因'] = f'输入搜索关键词失败: {str(e)[:30]}'
            logger.warning(f"  └─ 结果: 失败 - {log_entry['失败原因']}")
            return log_entry
        
        # 步骤3: 执行搜索
        logger.info(f"  ├─ 步骤3/6: 正在执行搜索...")
        try:
            search_button = find_search_button(driver)
            if not search_button:
                try:
                    search_button = driver.execute_script("""
                        var elements = document.querySelectorAll('input, button');
                        for (var i = 0; i < elements.length; i++) {
                            var elem = elements[i];
                            var value = elem.value || elem.textContent || elem.innerText || '';
                            var type = elem.type || '';
                            if (value.includes('搜索') || elem.getAttribute('onclick') ||
                                elem.className.includes('but') || elem.id.includes('search')) {
                                return elem;
                            }
                            if (type === 'submit' || type === 'button') { return elem; }
                        }
                        return null;
                    """)
                except:
                    pass
            if not search_button:
                log_entry['验证状态'] = '失败'
                log_entry['失败原因'] = '找不到搜索按钮'
                logger.warning(f"  └─ 结果: 失败 - {log_entry['失败原因']}")
                return log_entry
            try:
                search_button.click()
            except:
                driver.execute_script("arguments[0].click();", search_button)
            time.sleep(3)
        except Exception as e:
            log_entry['验证状态'] = '失败'
            log_entry['失败原因'] = f'执行搜索失败: {str(e)[:30]}'
            logger.warning(f"  └─ 结果: 失败 - {log_entry['失败原因']}")
            return log_entry
        
        # 步骤4: 查找目标公告链接
        # 用公告名称关键词匹配链接文字，同时检查URL中日期
        logger.info(f"  ├─ 步骤4/6: 正在查找目标公告...")
        announcement_link = None
        all_links = driver.find_elements(By.TAG_NAME, "a")
        logger.info(f"  │  页面中找到 {len(all_links)} 个链接")
        
        # 提取公告名称的关键词（取前10个字符作为核心匹配词）
        name_core = announcement_name[:10] if announcement_name else ''
        
        # 优先：URL含日期 且 链接文字含公告名称关键词
        for link in all_links:
            try:
                href = link.get_attribute('href')
                link_text = link.text.strip()
                if href and 'newsdetail' in href and contains_target_date(href, announcement_date):
                    if name_core and name_core in link_text:
                        announcement_link = href
                        logger.info(f"  │  找到精确匹配链接: {link_text[:50]}...")
                        break
            except:
                continue
        # 兜底：只匹配日期
        if not announcement_link:
            for link in all_links:
                try:
                    href = link.get_attribute('href')
                    if href and 'newsdetail' in href and contains_target_date(href, announcement_date):
                        announcement_link = href
                        logger.info(f"  │  找到日期匹配链接: {link.text.strip()[:50]}...")
                        break
                except:
                    continue
        if not announcement_link:
            log_entry['验证状态'] = '失败'
            log_entry['失败原因'] = f'未找到{announcement_date}的公告链接'
            logger.warning(f"  └─ 结果: 失败 - {log_entry['失败原因']}")
            return log_entry
        
        # 步骤5: 打开公告详情页
        logger.info(f"  ├─ 步骤5/6: 正在打开公告详情页...")
        driver.get(announcement_link)
        time.sleep(3)
        
        # 步骤6: 四项校验
        # 6-1: 校验公告发布日期
        logger.info(f"  ├─ 步骤6/6: 四项校验...")
        logger.info(f"  │  6-1: 校验公告发布日期（预期: {announcement_date}）...")
        date_status, date_msg = verify_announcement_date(driver, announcement_link, announcement_date)
        log_entry['日期验证'] = f"{date_status}：{date_msg}"
        date_sym = "√" if date_status == "成功" else ("○" if date_status == "跳过" else "✗")
        logger.info(f"  │  {date_sym} {date_msg}")
        date_ok = (date_status == "成功" or date_status == "跳过")
        
        # 6-2: 校验公告详情名称与Excel标题是否一致
        logger.info(f"  │  6-2: 校验公告详情名称与Excel标题是否一致...")
        title_ok, title_msg = validate_announcement_title(driver, expected_title)
        log_entry['公告标题验证'] = f"{'成功' if title_ok else '失败'}：{title_msg}"
        title_sym = "√" if title_ok else "✗"
        logger.info(f"  │  {title_sym} {title_msg}")
        
        # 6-3: 扫描并校验附件名称
        logger.info(f"  │  6-3: 扫描并校验附件名称...")
        attachment_links = []
        for link in driver.find_elements(By.TAG_NAME, "a"):
            try:
                href = link.get_attribute('href')
                text = link.text.strip()
                if href and ('.pdf' in href.lower() or '.doc' in href.lower() or
                            '.docx' in href.lower() or '.xls' in href.lower() or
                            '.xlsx' in href.lower() or '.rar' in href.lower() or
                            '.zip' in href.lower()):
                    if text and ('下载' in text or '附件' in text or
                                 'PDF' in text.upper() or 'DOC' in text.upper() or
                                 'XLS' in text.upper()):
                        attachment_links.append({'text': text, 'href': href})
                        logger.info(f"  │    找到附件: {text[:60]}...")
            except:
                continue
        # 去重
        seen_hrefs = set()
        unique_attachments = []
        for att in attachment_links:
            if att['href'] not in seen_hrefs:
                unique_attachments.append(att)
                seen_hrefs.add(att['href'])
        attachment_links = unique_attachments
        log_entry['附件数量'] = len(attachment_links)
        if not attachment_links:
            log_entry['验证状态'] = '失败'
            log_entry['失败原因'] = '无附件'
            logger.warning(f"  └─ 结果: 失败 - {log_entry['失败原因']}")
            return log_entry
        
        name_statuses = []
        name_results  = []
        for i, att in enumerate(attachment_links):
            att_text = att.get('text', '')
            logger.info(f"  │    附件{i+1}: {att_text[:60]}")
            status, msg = verify_attachment_by_name(att_text, announcement_name)
            name_statuses.append(status)
            name_results.append(msg)
            log_sym = "√" if status == "成功" else "✗"
            logger.info(f"  │    {log_sym} 附件名: {msg}")
        log_entry['附件状态']    = '; '.join(name_statuses)
        log_entry['附件名称验证'] = '; '.join(name_results)
        
        # 6-4: 下载PDF并校验内部标题
        logger.info(f"  │  6-4: 下载PDF校验内部标题（预期: '{expected_title}'）...")
        pdf_statuses = []
        pdf_results  = []
        clean_download_folder(download_path)
        for i, att in enumerate(attachment_links):
            href = att.get('href', '')
            # 只处理PDF，跳过其他格式
            if not href.lower().endswith('.pdf'):
                logger.info(f"  │    附件{i+1}: 非PDF格式，跳过内容校验")
                pdf_statuses.append("跳过")
                pdf_results.append("非PDF格式，跳过")
                continue
            logger.info(f"  │    附件{i+1}: 下载并读取PDF内容...")
            p_status, p_msg = verify_pdf_content_title(driver, href, announcement_name, download_path)
            pdf_statuses.append(p_status)
            pdf_results.append(p_msg)
            log_sym = "√" if p_status == "成功" else "✗"
            logger.info(f"  │    {log_sym} PDF标题: {p_msg}")
        log_entry['PDF内容标题验证'] = '; '.join(pdf_results)
        
        # ---- 综合判断（四项全过才算成功）----
        # 日期验证
        date_ok = (date_status == "成功" or date_status == "跳过")
        # 公告详情名称
        title_ok_bool = title_ok
        # 附件名：至少有一个成功
        name_ok  = any(s == "成功" for s in name_statuses)
        name_all = all(s == "成功" for s in name_statuses)
        # PDF标题：有效PDF（非跳过）都通过
        effective_pdf = [(s, r) for s, r in zip(pdf_statuses, pdf_results) if s != "跳过"]
        if effective_pdf:
            pdf_ok  = any(s == "成功" for s, _ in effective_pdf)
            pdf_all = all(s == "成功" for s, _ in effective_pdf)
        else:
            # 全是非PDF附件，PDF内容校验视为跳过，不计入失败
            pdf_ok = pdf_all = True

        failures = []
        if not date_ok:
            failures.append(f"日期不匹配({date_msg})")
        if not title_ok_bool:
            failures.append("公告详情名称不匹配")
        if not name_all:
            failures.append(f"附件名不一致({sum(1 for s in name_statuses if s!='成功')}个)")
        if not pdf_all and effective_pdf:
            failures.append(f"PDF内容标题不一致({sum(1 for s,_ in effective_pdf if s!='成功')}个)")

        # 四项全过 = 成功
        if date_ok and title_ok_bool and name_all and pdf_all:
            log_entry['验证状态'] = '成功'
            log_entry['失败原因'] = '无'
            logger.info(f"  └─ 结果: 成功 - 日期+公告详情+附件名+PDF标题均通过")
        # 至少有一项通过 = 部分失败
        elif date_ok or title_ok_bool or name_ok or pdf_ok:
            log_entry['验证状态'] = '部分失败'
            log_entry['失败原因'] = '; '.join(failures)
            logger.warning(f"  └─ 结果: 部分失败 - {log_entry['失败原因']}")
        # 全部失败
        else:
            log_entry['验证状态'] = '失败'
            log_entry['失败原因'] = '; '.join(failures) or '四项验证均失败'
            logger.warning(f"  └─ 结果: 失败 - {log_entry['失败原因']}")

    except Exception as e:
        error_msg = str(e)
        log_entry['验证状态'] = '失败'
        log_entry['失败原因'] = f"异常:{error_msg[:30] if error_msg else '未知错误'}"
        logger.error(f"  └─ 结果: 失败 - {error_msg[:100] if error_msg else '未知错误'}")
        clean_download_folder(download_path)
    
    return log_entry

# ==================== 主函数 ====================
def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    download_path = os.path.join(script_dir, 'temp_downloads')
    if not os.path.exists(download_path):
        os.makedirs(download_path)
    
    # 读取配置
    config_file = os.path.join(script_dir, 'config.txt')
    if not os.path.exists(config_file):
        logger.error("配置文件 config.txt 不存在！")
        return
    
    with open(config_file, 'r', encoding='utf-8') as f:
        lines = [l.strip() for l in f if l.strip() and not l.strip().startswith('#')]
    
    if len(lines) < 2:
        logger.error("配置文件格式错误！第一行应为Excel路径，第二行应为公告日期")
        return
    
    excel_path, announcement_date = lines[0], lines[1]
    
    # 处理路径：将 Windows 路径分隔符转换为系统分隔符
    excel_path = excel_path.replace('\\', os.sep).replace('/', os.sep)
    
    # 验证公告日期格式
    if not re.match(r'\d{4}-\d{2}-\d{2}', announcement_date):
        logger.error(f"公告日期格式错误: {announcement_date}，应为YYYY-MM-DD格式")
        return
    
    logger.info(f"公告日期: {announcement_date}")
    
    # 读取Excel文件
    if not os.path.exists(excel_path):
        logger.error(f"Excel文件不存在: {excel_path}")
        return
    
    logger.info(f"加载Excel文件: {excel_path}")
    
    # 读取Excel文件，使用object类型保持原始格式
    try:
        df = pd.read_excel(excel_path, dtype=object)
        logger.info(f"成功加载数据，共 {len(df)} 条记录")
    except Exception as e:
        logger.error(f"读取Excel文件失败: {str(e)}")
        return
    
    # 识别列名
    n_col, u_col = None, None
    for c in df.columns:
        col_str = str(c)
        if '标题' in col_str or '公告名称' in col_str: 
            n_col = c
        elif '公告地址' in col_str or 'URL' in col_str.upper(): 
            u_col = c
    
    if n_col is None:
        logger.error("未找到'公告名称'或'标题'列！")
        logger.info(f"可用列: {list(df.columns)}")
        return
    
    logger.info(f"识别列名: 公告名称={n_col}, 公告地址={u_col}")
    
    # 确保结果列存在
    result_cols = ['验证时间', '验证状态', '失败原因', '附件数量', '附件状态', '日期验证', '公告标题验证', '附件名称验证', 'PDF内容标题验证']
    for col in result_cols:
        if col not in df.columns:
            if col == '附件数量':
                df[col] = pd.NA
            else:
                df[col] = ''
            logger.info(f"添加结果列: {col}")
    
    # 分析已处理记录，实现断点续跑
    logger.info("正在分析Excel文件，检测已处理记录...")
    
    # 查找已完成的记录
    completed_records = []
    pending_records = []
    
    for idx in df.index:
        status_value = df.at[idx, '验证状态']
        # 判断是否已处理：状态列不为空且不是"进行中"
        if pd.notna(status_value) and str(status_value).strip() not in ['', '进行中']:
            completed_records.append(idx)
        else:
            pending_records.append(idx)
    
    total = len(df)
    completed = len(completed_records)
    pending = len(pending_records)
    
    logger.info(f"=== 断点续跑分析 ===")
    logger.info(f"总记录数: {total}")
    logger.info(f"已处理记录: {completed}")
    logger.info(f"待处理记录: {pending}")
    
    if pending == 0:
        logger.info("所有记录已处理完成，无需继续验证。")
        return
    
    # 询问用户是否继续处理
    if completed > 0:
        logger.info(f"发现 {completed} 条已处理记录，将从第 {pending_records[0]+1} 条记录开始继续处理。")
        response = input("是否继续处理？(y/n): ").strip().lower()
        if response != 'y':
            logger.info("程序退出")
            return
    
    # 浏览器配置
    logger.info("正在初始化浏览器...")
    options = webdriver.ChromeOptions()
    
    # 无头模式
    options.add_argument('--headless=new')
    
    # 优化参数
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-images')
    options.add_argument('--disable-plugins')
    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-dev-tools')
    options.add_argument('--log-level=3')
    options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    
    # 下载配置 - 关键设置
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_settings.popups": 0,
        "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.javascript": 1,
        # 关键：禁止浏览器预览PDF，强制直接下载
        "plugins.always_open_pdf_externally": True,
        "download.extensions_to_open": ""
    }
    options.add_experimental_option("prefs", prefs)
    
    # 初始化驱动 - 跨平台处理
    logger.info(f"检测系统平台: {platform.system()}")
    
    try:
        # 使用 webdriver-manager 自动管理 ChromeDriver 版本
        logger.info("正在获取与当前 Chrome 版本匹配的 ChromeDriver...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        logger.info("浏览器驱动初始化成功")
    except Exception as e:
        logger.error(f"自动管理方式失败: {str(e)}")
        logger.info("\n尝试手动指定 ChromeDriver...")
        
        # 备选方案：手动查找（Windows 优先策略）
        possible_driver_known_ = []

        # 1. 当前脚本目录
        possible_driver_paths.append(os.path.join(script_dir, "chromedriver.exe"))

        # 2. Windows 默认安装路径
        possible_driver_paths.append(r"C:\Program Files\Google\Chrome\Application\chromedriver.exe")
        possible_driver_paths.append(r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe")

        # 3. 系统 PATH
        possible_driver_paths.append("chromedriver.exe")

        driver_path = None
        for path in possible_driver_paths:
            if os.path.exists(path):
                driver_path = path
                logger.info(f"找到 ChromeDriver: {driver_path}")
                break
        
        if not driver_path:
            logger.error(f"ChromeDriver 未找到！已尝试的路径:")
            for path in possible_driver_paths:
                logger.error(f"  - {path}")
            logger.info("\n请按以下步骤解决:")
            logger.info("方案 1: 安装/更新 webdriver-manager")
            logger.info("  pip install --upgrade webdriver-manager")
            logger.info("\n方案 2: 手动下载正确版本的 ChromeDriver")
            logger.info("  1. 访问: https://googlechromelabs.github.io/chrome-for-testing/")
            logger.info("  2. 下载与 Chrome 145 对应的 ChromeDriver")
            logger.info("  3. 放在: /usr/local/bin/ 或与本脚本相同目录")
            if platform.system() == "Darwin":
                logger.info("  4. 运行: chmod +x /usr/local/bin/chromedriver")
            return
        
        try:
            service = Service(driver_path)
            driver = webdriver.Chrome(service=service, options=options)
            logger.info("浏览器驱动初始化成功（手动路径）")
        except Exception as e2:
            logger.error(f"初始化浏览器驱动失败: {str(e2)}")
            logger.error("\n版本不匹配问题:")
            logger.error("  ChromeDriver 版本与 Chrome 浏览器版本不一致")
            logger.error("  请检查: chromedriver --version")
            logger.error("  请检查: /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --version")
            logger.error("\n解决方案:")
            logger.error("  1. 使用 webdriver-manager 自动管理版本")
            logger.error("  2. 或手动下载匹配版本的 ChromeDriver")
            return
    
    # 开始处理
    actually_processed_count = 0
    start_time = time.time()
    
    logger.info(f"=== 开始批量验证 ===")
    logger.info(f"公告日期: {announcement_date}")
    logger.info(f"结果保存到: {excel_path}")
    logger.info("-" * 50)
    
    for idx in pending_records:
        row = df.iloc[idx]
        name = str(row[n_col]) if n_col and pd.notna(row[n_col]) else ''
        url = str(row[u_col]) if u_col and pd.notna(row[u_col]) else ''
        
        if not name:
            logger.warning(f"记录 {idx+1} 公告名称为空，跳过")
            continue
        
        # 执行验证（用公告名称搜索）
        res = search_and_verify(driver, name, url, announcement_date, download_path, 
                              idx, total, completed + actually_processed_count)
        
        # 更新DataFrame
        df.at[df.index[idx], '验证时间'] = str(res['验证时间'])
        df.at[df.index[idx], '验证状态'] = str(res['验证状态'])
        df.at[df.index[idx], '失败原因'] = str(res['失败原因'])
        df.at[df.index[idx], '附件数量'] = int(res['附件数量'])
        df.at[df.index[idx], '附件状态'] = str(res['附件状态'])
        df.at[df.index[idx], '日期验证'] = str(res.get('日期验证', ''))
        df.at[df.index[idx], '公告标题验证'] = str(res['公告标题验证'])
        df.at[df.index[idx], '附件名称验证'] = str(res['附件名称验证'])
        df.at[df.index[idx], 'PDF内容标题验证'] = str(res['PDF内容标题验证'])
        
        # 保存到原Excel文件
        try:
            df.to_excel(excel_path, index=False)
            logger.info(f"  √ 已保存到原文件: {excel_path}")
        except Exception as e:
            logger.error(f"保存Excel文件失败: {str(e)}")
            # 尝试使用备份名称保存
            backup_name = f"{os.path.splitext(excel_path)[0]}_备份_{int(time.time())}.xlsx"
            try:
                df.to_excel(backup_name, index=False)
                logger.info(f"已保存备份到: {backup_name}")
            except Exception as e2:
                logger.error(f"保存备份文件也失败: {str(e2)}")
        
        actually_processed_count += 1
        
        # 显示进度
        current_total_processed = completed + actually_processed_count
        progress = (current_total_processed) / total * 100
        elapsed = time.time() - start_time
        
        if actually_processed_count > 0:
            avg_time_per_item = elapsed / actually_processed_count
            remaining_items = total - current_total_processed
            eta = avg_time_per_item * remaining_items
        else:
            eta = 0
        
        logger.info(f"进度: {current_total_processed}/{total} [{progress:.1f}%] | "
                   f"用时: {elapsed:.0f}s | "
                   f"剩余: {eta:.0f}s")
        logger.info("-" * 50)
    
    # 清理
    try:
        driver.quit()
    except:
        pass
    
    clean_download_folder(download_path)
    try: 
        os.rmdir(download_path)
    except: 
        pass
    
    # 最终统计
    logger.info("=== 验证完成 ===")
    total_time = time.time() - start_time
    success_count = 0
    failed_count = 0
    partial_count = 0
    
    for idx in df.index:
        status_value = df.at[idx, '验证状态']
        if pd.notna(status_value):
            status_str = str(status_value).strip()
            if status_str == '成功':
                success_count += 1
            elif status_str == '失败':
                failed_count += 1
            elif status_str in ['部分失败', '部分成功']:
                partial_count += 1
    
    logger.info(f"总计: {total} 条")
    logger.info(f"成功: {success_count} 条")
    logger.info(f"部分成功/失败: {partial_count} 条")
    logger.info(f"失败: {failed_count} 条")
    logger.info(f"总用时: {total_time:.1f}秒")
    if actually_processed_count > 0:
        logger.info(f"平均每条: {total_time/actually_processed_count:.1f}秒")
    logger.info(f"结果已保存到原文件: {excel_path}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.warning("\n程序被用户中断，Excel文件已保存，下次运行将自动从断点继续。")
    except Exception as e:
        logger.error(f"程序异常终止: {str(e)}")
        import traceback
        logger.error(f"详细错误信息:\n{traceback.format_exc()}")
        logger.info("Excel文件已保存，下次运行将自动从断点继续。")
