#!/usr/bin/env python3
"""
从发票文件名提取信息（不需要OCR）
支持多种发票文件名格式
"""

import os
import re
from pathlib import Path
import openpyxl
from datetime import datetime

class FilenameExtractor:
    def __init__(self, base_path):
        self.base_path = base_path
        self.results = []
        
    def parse_filename(self, file_path):
        """从文件名提取信息"""
        filename = os.path.basename(file_path)
        folder = os.path.basename(os.path.dirname(file_path))
        
        info = {
            '文件名': filename,
            '文件夹': folder,
            '日期': '',
            '金额': '',
            '发票号': '',
            '公司名称': '',
            '项目名称': '',
            '付款人': '',
            '收款人': '',
            '用途': '',
            '备注': ''
        }
        
        name_without_ext = os.path.splitext(filename)[0]
        
        # 特殊处理：简单数字文件（如 26 4.pdf, 51 1.pdf 等）
        # 注意：只有纯数字（可能带小数点）才处理，避免误处理其他格式
        simple_num_match = re.match(r'^(\d+(?:\.\d+)?)\s*$', name_without_ext.strip())
        if simple_num_match:
            try:
                num_str = simple_num_match.group(1)
                # 只有当整个文件名就是数字时才处理，避免处理 65a8b60b4a92eb55260ff7c8.pdf
                if re.match(r'^\d+(?:\.\d+)?$', name_without_ext.strip()):
                    num = float(num_str)
                    if 100 <= num <= 1000000:
                        info['金额'] = num
                        info['用途'] = '待确认项目'
            except:
                pass
        
        # 模式1: dzfp_24312000000383480812_上海正欣隆餐饮管理有限公司_20241206130302.pdf
        pattern1 = r'dzfp_(\d+)_(.+?)_(\d{14})'
        match1 = re.search(pattern1, name_without_ext)
        if match1:
            info['发票号'] = match1.group(1)
            info['收款人'] = match1.group(2)
            date_str = match1.group(3)
            if len(date_str) == 14:
                info['日期'] = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
            # 注意：这种格式的文件名通常不包含金额，且末尾是时间戳
            # 不提取金额，避免把时间戳当金额
        
        # 模式2: 佳翔 dzfp_25312000000001487217_上海佳翔汽车租赁有限公司_20250307135324.pdf
        pattern2 = r'^(.+?)\s+dzfp_(\d+)_(.+?)_(\d{14})'
        match2 = re.search(pattern2, name_without_ext)
        if match2:
            info['项目名称'] = match2.group(1)
            info['发票号'] = match2.group(2)
            info['收款人'] = match2.group(3)
            date_str = match2.group(4)
            if len(date_str) == 14:
                info['日期'] = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
        
        # 模式3: 243120000004_06480088_51.70_苏州众勤会计师事务所有限公司.pdf
        pattern3 = r'^(\d+)_(\d+)_(\d+(?:\.\d+)?)_(.+)$'
        match3 = re.search(pattern3, name_without_ext)
        if match3:
            info['发票号'] = match3.group(1)
            info['金额'] = float(match3.group(3))
            info['收款人'] = match3.group(4)
        
        # 模式3b: 文件名中包含发票号（通常是一串连续数字，8-20位）
        # 例如：20210129-￥407000.00.png 这种格式可能不包含发票号
        # 但某些格式如：发票号_其他信息.ext
        if not info['发票号']:
            # 查找下划线分隔的长数字序列
            invoice_match = re.search(r'(?<![a-zA-Z])(\d{8,20})(?:_|$)', name_without_ext)
            if invoice_match:
                potential_invoice = invoice_match.group(1)
                # 排除14位时间戳（如20241206130302）
                if len(potential_invoice) != 14:
                    info['发票号'] = potential_invoice
        
        # 模式4: LH-P-2024-396发票.pdf
        pattern4 = r'LH-P-(\d{4})-(\d+)'
        match4 = re.search(pattern4, name_without_ext)
        if match4:
            info['项目名称'] = f"LH-P-{match4.group(1)}项目"
            # 尝试提取末尾的数字作为金额
            amount_match = re.search(r'(\d+(?:\.\d+)?)\s*$', name_without_ext)
            if amount_match:
                try:
                    amount = float(amount_match.group(1))
                    if 10 < amount < 1000000:  # 合理的发票金额范围
                        info['金额'] = amount
                except:
                    pass
        
        # 模式5: LK-1-1a1 预提期后发票-追梦者 41460 AQEE.pdf
        pattern5 = r'LK-\d+-\w+\s+.+?\s+(\d+(?:\.\d+)?)'
        match5 = re.search(pattern5, name_without_ext)
        if match5:
            info['项目名称'] = 'LK项目'
            try:
                info['金额'] = float(match5.group(1))
            except:
                pass
        
        # 模式6: 包含日期格式 YYYYMMDD 或 YYYY-MM-DD
        date_patterns = [
            r'(\d{4})(\d{2})(\d{2})',  # YYYYMMDD
            r'(\d{4})-(\d{2})-(\d{2})',  # YYYY-MM-DD
        ]
        if not info['日期']:
            for pattern in date_patterns:
                date_match = re.search(pattern, name_without_ext)
                if date_match:
                    info['日期'] = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"
                    break
        
        # 通用金额提取（如果还没找到金额）
        # 注意：大部分发票文件名不包含金额，强行提取会导致错误
        # 因此这里只提取非常明确的金额，否则留空
        if not info['金额']:
            # 只查找带有小数点且范围合理的数字（这通常是金额）
            amount_match = re.search(r'(?<!\d)(\d{1,3}(?:,\d{3})*\.\d{1,2}|\d+\.\d{1,2})(?!\d)', name_without_ext)
            if amount_match:
                try:
                    amount_str = amount_match.group(1).replace(',', '')
                    amount = float(amount_str)
                    # 只有金额在合理范围内才采用（100元到100万元之间）
                    if 100 <= amount <= 1000000:
                        info['金额'] = amount
                except:
                    pass
        
        # 尝试从文件名提取更多数字作为可能的发票号
        if not info['发票号']:
            # 查找所有8-20位的数字序列
            all_numbers = re.findall(r'(?<![a-zA-Z])(\d{8,20})(?![a-zA-Z])', name_without_ext)
            for num in all_numbers:
                # 排除日期时间戳（14位）、金额（通常带小数点或逗号）
                if len(num) != 14 and '.' not in num and ',' not in num:
                    # 检查是否是合理的发票号长度（通常是20位左右）
                    if 8 <= len(num) <= 30:
                        info['发票号'] = num
                        break
        
        # 如果没有公司名称，尝试从文件夹名获取
        if not info['公司名称'] and folder:
            # 移除 "D:\code\autotest\invoice_pdf" 这样的路径部分
            folder_clean = folder.replace('invoice_pdf', '').strip()
            if folder_clean and not folder_clean.startswith('D:') and not folder_clean.startswith('/'):
                info['公司名称'] = folder_clean
        
        # 推断用途
        if not info['用途']:
            if info['项目名称']:
                info['用途'] = info['项目名称']
            elif info['收款人']:
                info['用途'] = info['收款人']
        
        return info
    
    def process_all_files(self):
        """处理所有文件"""
        supported_exts = ['.png', '.jpg', '.jpeg', '.pdf']
        
        for root, dirs, files in os.walk(self.base_path):
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            
            for file in sorted(files):
                ext = os.path.splitext(file)[1].lower()
                if ext in supported_exts:
                    file_path = os.path.join(root, file)
                    print(f"处理: {file}")
                    info = self.parse_filename(file_path)
                    self.results.append(info)
    
    def save_to_excel(self, output_path):
        """保存到Excel"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "发票信息汇总"
        
        # 设置列宽
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 25
        ws.column_dimensions['G'].width = 20
        ws.column_dimensions['H'].width = 30
        ws.column_dimensions['I'].width = 30
        
        # 写入表头
        headers = [
            '文件夹', '文件名', '日期', '金额（元）', '发票号',
            '公司名称', '收款人', '项目名称', '用途'
        ]
        ws.append(headers)
        
        # 设置表头样式
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        
        # 写入数据
        for info in self.results:
            row = [
                info['文件夹'],
                info['文件名'],
                info['日期'],
                info['金额'],
                info['发票号'],
                info['公司名称'],
                info['收款人'],
                info['项目名称'],
                info['用途']
            ]
            ws.append(row)
        
        wb.save(output_path)
        print(f"\n[成功] Excel已保存到: {output_path}")
        print(f"[成功] 共处理 {len(self.results)} 个文件")
        
        # 统计信息
        total_amount = sum([r['金额'] for r in self.results if isinstance(r['金额'], (int, float))])
        print(f"[成功] 总金额: {total_amount:,.2f} 元")
        
        # 按收款人统计
        company_stats = {}
        for r in self.results:
            company = r['收款人'] or r['公司名称']
            if company:
                if company not in company_stats:
                    company_stats[company] = {'count': 0, 'amount': 0}
                company_stats[company]['count'] += 1
                if isinstance(r['金额'], (int, float)):
                    company_stats[company]['amount'] += r['金额']
        
        print(f"\n按收款公司统计:")
        for company, stats in sorted(company_stats.items(), key=lambda x: x[1]['amount'], reverse=True):
            print(f"  {company}: {stats['count']}笔, {stats['amount']:,.2f}元")

if __name__ == '__main__':
    import sys
    from datetime import datetime
    
    base_path = sys.argv[1] if len(sys.argv) > 1 else r'D:\code\autotest\invoice_pdf'
    
    # 在输入目录中生成输出文件，文件名包含时间戳
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(base_path, f'发票信息汇总_{timestamp}.xlsx')
    
    print("=" * 70)
    print("发票信息提取工具")
    print("=" * 70)
    print(f"输入目录: {base_path}")
    print(f"输出文件: {output_path}")
    print("=" * 70)
    
    extractor = FilenameExtractor(base_path)
    extractor.process_all_files()
    extractor.save_to_excel(output_path)
    
    print("\n" + "=" * 70)
    print("处理完成！")
    print("=" * 70)
