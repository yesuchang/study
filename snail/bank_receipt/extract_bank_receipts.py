import os
import re
import sys
from pathlib import Path
from datetime import datetime

# 设置控制台编码为UTF-8
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

try:
    from PIL import Image
    import pytesseract
    from pdf2image import convert_from_path
    import openpyxl
except ImportError:
    print("正在安装必要的依赖...")
    os.system("pip install pytesseract openpyxl pdf2image Pillow")
    from PIL import Image
    import pytesseract
    from pdf2image import convert_from_path
    import openpyxl

class BankReceiptExtractor:
    def __init__(self, base_path):
        self.base_path = base_path
        self.results = []
        self.total_files = 0
        self.processed_files = 0
        self.skipped_files = 0
    
    def is_valid_file(self, file_path):
        """检查文件是否有效（非空、可读）"""
        try:
            return os.path.getsize(file_path) > 0
        except:
            return False
    
    def extract_from_image(self, image_path):
        """从图片提取文本"""
        try:
            image = Image.open(image_path)
            # 转灰度图提高OCR准确率
            if image.mode != 'L':
                image = image.convert('L')
            text = pytesseract.image_to_string(image, lang='chi_sim+eng')
            return text.strip()
        except Exception as e:
            print(f"  [警告] 提取图片失败: {e}")
            return ""
    
    def extract_from_pdf(self, pdf_path):
        """从PDF提取文本"""
        try:
            images = convert_from_path(pdf_path, dpi=300)
            text = ""
            for i, img in enumerate(images):
                print(f"  [进度] 处理PDF第 {i+1}/{len(images)} 页...")
                # 转灰度图
                if img.mode != 'L':
                    img = img.convert('L')
                page_text = pytesseract.image_to_string(img, lang='chi_sim+eng')
                text += page_text + "\n"
            return text.strip()
        except Exception as e:
            print(f"  [错误] 提取PDF失败: {e}")
            return ""
    
    def parse_receipt_info(self, text, file_path):
        """解析回单信息"""
        info = {
            '文件名': os.path.basename(file_path),
            '文件夹': os.path.basename(os.path.dirname(file_path)),
            '日期': '',
            '流水号': '',
            '付款人全称': '',
            '付款人账号': '',
            '付款人开户行': '',
            '收款人全称': '',
            '收款人账号': '',
            '收款人开户行': '',
            '金额（大写）': '',
            '金额（小写）': '',
            '用途': '',
            '摘要': ''
        }
        
        if not text:
            print(f"  [警告] OCR未提取到文本内容")
            return info
        
        # 调试：打印OCR结果的前200字符
        print(f"  [调试] OCR文本预览: {text[:200]}...")
        
        # 提取日期 - 多种格式
        date_patterns = [
            # 2024年01月01日 / 2024年1月1日
            r'(\d{4})\s*[年/-]\s*(\d{1,2})\s*[月/-]\s*(\d{1,2})\s*[日]?',
            # 交易日期/记账日期
            r'(?:交易|记账|业务)[日期\s]*[：:]\s*(\d{4}\D?\d{2}\D?\d{2})',
            # 2024-01-01 或 2024/01/01
            r'(\d{4})[-/](\d{2})[-/](\d{2})',
            # 纯数字 20240101
            r'(\d{4})(\d{2})(\d{2})',
        ]
        for pattern in date_patterns:
            match = re.search(pattern, text)
            if match:
                groups = match.groups()
                if len(groups) == 3:
                    info['日期'] = f"{groups[0]}-{groups[1].zfill(2)}-{groups[2].zfill(2)}"
                    break
        
        # 提取流水号/交易号
        flow_patterns = [
            r'(?:流水|交易|业务)[\s]*[号编号：:]\s*(\S{10,})',
            r'账户明细编号[-\u2014](\S+)',
            r'凭证[号编号：:]\s*(\S+)',
            r'(\d{16,})',  # 16位以上的数字
        ]
        for pattern in flow_patterns:
            match = re.search(pattern, text)
            if match:
                info['流水号'] = match.group(1).strip()
                break
        
        # 提取付款人信息 - 使用更宽松的正则
        payer_patterns = [
            # 付款人全称
            (r'(?:付款人|付款方|付方|出票人)[\s]*(?:名称|全称|户名|单位)[\s]*[：:]\s*(.+?)(?:\s|$)', '付款人全称'),
            (r'(?:付款人|付款方|付方)[\s]*(?:名称|全称|户名)[\s]*[：:]\s*(.+?)(?:\s{2,}|$)', '付款人全称'),
            # 付款人账号
            (r'(?:付款人|付款方|付方)[\s]*(?:账号|卡号|账户)[\s]*[：:]\s*(\d[\d\s]{5,})', '付款人账号'),
            # 付款人开户行
            (r'(?:付款人|付款方|付方)[\s]*(?:开户行|银行|行名)[\s]*[：:]\s*(.+?)(?:\s|$)', '付款人开户行'),
        ]
        for pattern, key in payer_patterns:
            match = re.search(pattern, text, re.DOTALL)
            if match:
                value = match.group(1).strip()
                if len(value) > 1:  # 忽略空值
                    info[key] = value
        
        # 提取收款人信息
        payee_patterns = [
            (r'(?:收款人|收款方|收方|持票人)[\s]*(?:名称|全称|户名|单位)[\s]*[：:]\s*(.+?)(?:\s|$)', '收款人全称'),
            (r'(?:收款人|收款方|收方)[\s]*(?:名称|全称|户名)[\s]*[：:]\s*(.+?)(?:\s{2,}|$)', '收款人全称'),
            (r'(?:收款人|收款方|收方)[\s]*(?:账号|卡号|账户)[\s]*[：:]\s*(\d[\d\s]{5,})', '收款人账号'),
            (r'(?:收款人|收款方|收方)[\s]*(?:开户行|银行|行名)[\s]*[：:]\s*(.+?)(?:\s|$)', '收款人开户行'),
        ]
        for pattern, key in payee_patterns:
            match = re.search(pattern, text, re.DOTALL)
            if match:
                value = match.group(1).strip()
                if len(value) > 1:
                    info[key] = value
        
        # 提取金额 - 大写
        amount_big_patterns = [
            r'[（(]大写[)）]\s*([壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分整]+)',
            r'大写[：:]\s*([壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分整]+)',
            r'金额[（(]大写[)）][\s：:]*([壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分整]+)',
        ]
        for pattern in amount_big_patterns:
            match = re.search(pattern, text)
            if match:
                info['金额（大写）'] = match.group(1).strip()
                break
        
        # 提取金额 - 小写
        amount_patterns = [
            r'[（(]小写[)）][\s：:]*[¥￥]?\s*([\d,，]+\.\d{2})',
            r'小写金额[\s：:]*[¥￥]?\s*([\d,，]+\.\d{2})',
            r'金额[（(]小写[)）][\s：:]*[¥￥]?\s*([\d,，]+\.\d{2})',
            r'[¥￥]\s*([\d,，]+\.\d{2})',  # 最后的匹配方式
        ]
        for pattern in amount_patterns:
            match = re.search(pattern, text)
            if match:
                amount_str = match.group(1).replace(',', '').replace('，', '')
                try:
                    info['金额（小写）'] = float(amount_str)
                except:
                    info['金额（小写）'] = amount_str
                break
        
        # 提取用途/附言
        purpose_patterns = [
            r'(?:用途|附言|交易用途|资金用途)[\s]*[：:]\s*(.+?)(?:\s{2,}|$)',
            r'(?:用途|附言)[\s]*[：:]\s*(.{2,50})',
        ]
        for pattern in purpose_patterns:
            match = re.search(pattern, text)
            if match:
                purpose = match.group(1).strip()
                if len(purpose) > 1:
                    info['用途'] = purpose
                    break
        
        # 提取摘要
        summary_patterns = [
            r'摘要[\s]*[：:]\s*(.+?)(?:\s{2,}|$)',
            r'摘要[\s]*[：:]\s*(.{2,100})',
        ]
        for pattern in summary_patterns:
            match = re.search(pattern, text)
            if match:
                summary = match.group(1).strip()
                if len(summary) > 1:
                    info['摘要'] = summary
                    break
        
        return info
    
    def process_all_files(self):
        """处理所有文件"""
        supported_exts = ['.png', '.jpg', '.jpeg', '.pdf', '.tif', '.tiff', '.bmp']
        
        # 统计文件总数
        for root, dirs, files in os.walk(self.base_path):
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in supported_exts:
                    self.total_files += 1
        
        print(f"\n找到 {self.total_files} 个待处理文件")
        
        for root, dirs, files in os.walk(self.base_path):
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            
            for file in sorted(files):  # 排序以保持一致性
                ext = os.path.splitext(file)[1].lower()
                if ext not in supported_exts:
                    continue
                    
                file_path = os.path.join(root, file)
                
                # 检查文件有效性
                if not self.is_valid_file(file_path):
                    self.skipped_files += 1
                    print(f"\n[跳过] {file_path} (文件为空或无效)")
                    continue
                
                self.processed_files += 1
                print(f"\n[{self.processed_files}/{self.total_files}] 处理: {file_path}")
                
                # 提取文本
                try:
                    if ext == '.pdf':
                        text = self.extract_from_pdf(file_path)
                    else:
                        text = self.extract_from_image(file_path)
                    
                    # 解析信息
                    info = self.parse_receipt_info(text, file_path)
                    self.results.append(info)
                    
                    # 输出解析结果摘要
                    if info['日期'] or info['金额（小写）']:
                        print(f"  [结果] 日期: {info['日期']}, 金额: {info['金额（小写）']}, "
                              f"付款人: {info['付款人全称'][:10]}..., 收款人: {info['收款人全称'][:10]}...")
                    else:
                        print(f"  [结果] 未提取到关键信息")
                        
                except Exception as e:
                    self.skipped_files += 1
                    print(f"  [错误] 处理文件异常: {e}")
    
    def save_to_excel(self, output_path):
        """保存到Excel"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "银行水单信息"
        
        # 设置列宽
        col_widths = {
            'A': 15, 'B': 30, 'C': 12, 'D': 25,
            'E': 25, 'F': 20, 'G': 25,
            'H': 25, 'I': 20, 'J': 25,
            'K': 15, 'L': 30, 'M': 30
        }
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width
        
        # 写入表头
        headers = [
            '文件夹', '文件名', '日期', '流水号',
            '付款人全称', '付款人账号', '付款人开户行',
            '收款人全称', '收款人账号', '收款人开户行',
            '金额（小写）', '用途', '摘要'
        ]
        ws.append(headers)
        
        # 设置表头样式
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # 写入数据
        for info in self.results:
            row = [
                info['文件夹'],
                info['文件名'],
                info['日期'],
                info['流水号'],
                info['付款人全称'],
                info['付款人账号'],
                info['付款人开户行'],
                info['收款人全称'],
                info['收款人账号'],
                info['收款人开户行'],
                info['金额（小写）'],
                info['用途'],
                info['摘要']
            ]
            ws.append(row)
        
        wb.save(output_path)
        print(f"\n{'='*50}")
        print(f"✅ Excel已保存到: {output_path}")
        print(f"📊 统计:")
        print(f"   - 总计文件: {self.total_files}")
        print(f"   - 成功处理: {self.processed_files}")
        print(f"   - 跳过/失败: {self.skipped_files}")
        print(f"   - 提取到信息: {len(self.results)}")
        print(f"{'='*50}")

if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='银行水单信息提取工具')
    parser.add_argument('--input', '-i', 
                       default=r'E:\老王\应付款发生变动对应的支付凭证 -20260703',
                       help='输入文件夹路径')
    parser.add_argument('--output', '-o',
                       default=None,
                       help='输出Excel文件路径（默认输入文件夹下）')
    parser.add_argument('--debug', '-d', action='store_true',
                       help='调试模式，打印OCR原始文本')
    
    args = parser.parse_args()
    
    base_path = args.input
    if args.output:
        output_path = args.output
    else:
        output_path = os.path.join(base_path, '银行水单信息汇总.xlsx')
    
    print(f"📁 输入目录: {base_path}")
    print(f"📄 输出文件: {output_path}")
    print(f"{'='*50}")
    
    extractor = BankReceiptExtractor(base_path)
    extractor.process_all_files()
    extractor.save_to_excel(output_path)
