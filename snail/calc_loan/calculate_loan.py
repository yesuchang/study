import pandas as pd
from datetime import datetime
import numpy as np

# 读取原始文档，跳过第一行（因为第一行是表头）
df = pd.read_excel('D:/code/code/司法鉴定-还本+利息测算-工作簿1.xlsx', sheet_name='原始文档', header=0)

# 重命名列
df.columns = ['分类', '序号', '日期', '借', '还', '利率', '间隔天数', '还款期限/还款说明', '还款对应关系']

# 转换日期列，跳过第一行的表头
df = df.iloc[1:].copy().reset_index(drop=True)
df['日期'] = pd.to_datetime(df['日期'], format='%Y-%m-%d %H:%M:%S', errors='coerce')

# 将NaN替换为0以便计算
df['借'] = pd.to_numeric(df['借'], errors='coerce').fillna(0)
df['还'] = pd.to_numeric(df['还'], errors='coerce').fillna(0)
df['利率'] = pd.to_numeric(df['利率'], errors='coerce').fillna(0)
df['间隔天数'] = pd.to_numeric(df['间隔天数'], errors='coerce').fillna(0)

# 添加借款编号和还款编号
df['借款编号'] = ''
df['还款编号'] = ''

borrow_count = 0
repay_count = 0

for idx, row in df.iterrows():
    if row['借'] > 0:
        borrow_count += 1
        df.at[idx, '借款编号'] = borrow_count
    if row['还'] > 0:
        repay_count += 1
        df.at[idx, '还款编号'] = repay_count

print("=" * 80)
print("原始文档数据分析")
print("=" * 80)

total_borrowed = df['借'].sum()
total_repaid = df['还'].sum()

print(f"\n总借款金额: {total_borrowed:,.2f} 元")
print(f"总还款金额: {total_repaid:,.2f} 元")
print(f"借款总笔数: {borrow_count} 笔")
print(f"还款总笔数: {repay_count} 笔")

# 分类统计
interest_borrow = df[df['利率'] > 0]
no_interest_borrow = df[df['利率'] == 0]

print(f"\n--- 有息借款（约定利率）---")
print(f"有息借款笔数: {len(interest_borrow)}")
print(f"有息借款本金: {interest_borrow['借'].sum():,.2f} 元")

print(f"\n--- 无息借款（未约定利率）---")
print(f"无息借款笔数: {len(no_interest_borrow)}")
print(f"无息借款本金: {no_interest_borrow['借'].sum():,.2f} 元")

# =============================================
# 先息后本 + 摊余成本法计算
# =============================================
print(f"\n--- 利息计算（先息后本 · 摊余成本法）---")

# 利率：所有有息借款利率均为24%
INTEREST_RATE = 0.24

# 计息截止日
CUTOFF_DATE = datetime(2020, 8, 19)

# 按日期排序所有行：同一天先借款后还款
df_sorted = df[(df['借'] > 0) | (df['还'] > 0)].copy()
df_sorted['排序日期'] = df_sorted['日期']
df_sorted['排序顺序'] = df_sorted.apply(lambda row: 0 if row['借'] > 0 else 1, axis=1)
df_sorted = df_sorted.sort_values(['排序日期', '排序顺序']).reset_index(drop=True)

# 获取所有借款记录（用于匹配）
borrow_records = df[df['借'] > 0].copy()

# 为每个借款编号标记是否是有息借款（利率>0）
borrow_has_interest = {}
for idx, row in borrow_records.iterrows():
    borrow_num = row['借款编号']
    borrow_has_interest[borrow_num] = (row['利率'] > 0)

# 摊余成本法计算
cumulative_interest_principal = 0.0  # 有息本金余额（参与计息）
cumulative_free_principal = 0.0     # 无息本金余额（不参与计息）
accrued_interest = 0.0              # 累计未付利息
prev_date = None
total_accrued_interest = 0.0        # 累计产生的总利息
interest_details = []

# 初始化df的追加列
df['有息本金'] = 0.0
df['无息本金'] = 0.0
df['期初余额'] = 0.0
df['期末余额'] = 0.0
df['利息'] = 0.0
df['未付利息'] = 0.0
df['还利息'] = 0.0
df['还本金'] = 0.0

print(f"\n先息后本 摊余成本法利息计算明细（利率{INTEREST_RATE*100:.0f}%，仅对有息本金计息）：")
print(f"{'日期':<12} {'说明':<24} {'有息本金':>12} {'无息本金':>12} {'利息':>10} {'还利息':>10} {'还本金':>10} {'有息余额':>12} {'未付息':>10}")
print("-" * 112)

for idx, row in df_sorted.iterrows():
    current_date = row['日期']
    principal = row['借']
    repayment = row['还']
    rate = row['利率']
    
    # 计算天数间隔
    days = 0
    if prev_date is not None and pd.notna(current_date) and pd.notna(prev_date):
        days = (current_date - prev_date).days
    
    # === 第一步：计提利息（仅对有息本金计息）===
    if cumulative_interest_principal > 0 and days > 0:
        period_interest = cumulative_interest_principal * INTEREST_RATE * days / 365
    else:
        period_interest = 0.0
    
    accrued_interest += period_interest
    total_accrued_interest += period_interest
    
    # === 第二步：处理当日的借款/还款 ===
    repay_interest = 0.0
    repay_principal = 0.0
    
    period_start_balance = cumulative_interest_principal + cumulative_free_principal
    
    if principal > 0:
        if rate > 0:
            cumulative_interest_principal += principal
        else:
            cumulative_free_principal += principal
    
    if repayment > 0:
        relation = row['还款对应关系']
        if pd.notna(relation):
            relation_str = str(relation).strip()
            if relation_str.startswith('利息'):
                if repayment <= accrued_interest:
                    repay_interest = repayment
                    accrued_interest -= repayment
                else:
                    repay_interest = accrued_interest
                    remaining = repayment - accrued_interest
                    accrued_interest = 0.0
                    repay_principal = remaining
                    if cumulative_interest_principal >= remaining:
                        cumulative_interest_principal -= remaining
                    else:
                        remaining -= cumulative_interest_principal
                        cumulative_interest_principal = 0.0
                        cumulative_free_principal -= remaining
                        if cumulative_free_principal < 0:
                            cumulative_free_principal = 0
            else:
                matching_borrows = borrow_records[borrow_records['还款对应关系'] == relation_str]
                has_interest_match = False
                has_free_match = False
                for _, br in matching_borrows.iterrows():
                    if br['利率'] > 0:
                        has_interest_match = True
                    else:
                        has_free_match = True
                
                if has_interest_match and not has_free_match:
                    repay_principal = repayment
                    repay_interest = 0.0
                    cumulative_interest_principal -= repayment
                    if cumulative_interest_principal < 0:
                        cumulative_interest_principal = 0
                elif not has_interest_match and has_free_match:
                    repay_principal = repayment
                    repay_interest = 0.0
                    cumulative_free_principal -= repayment
                    if cumulative_free_principal < 0:
                        cumulative_free_principal = 0
                else:
                    repay_principal = repayment
                    repay_interest = 0.0
                    if cumulative_interest_principal >= repayment:
                        cumulative_interest_principal -= repayment
                    else:
                        remaining = repayment - cumulative_interest_principal
                        cumulative_interest_principal = 0.0
                        cumulative_free_principal -= remaining
                        if cumulative_free_principal < 0:
                            cumulative_free_principal = 0
        else:
            if repayment <= accrued_interest:
                repay_interest = repayment
                accrued_interest -= repayment
            else:
                repay_interest = accrued_interest
                remaining = repayment - accrued_interest
                accrued_interest = 0.0
                repay_principal = remaining
                if cumulative_interest_principal >= remaining:
                    cumulative_interest_principal -= remaining
                else:
                    remaining -= cumulative_interest_principal
                    cumulative_interest_principal = 0.0
                    cumulative_free_principal -= remaining
                    if cumulative_free_principal < 0:
                        cumulative_free_principal = 0
    
    period_end_balance = cumulative_interest_principal + cumulative_free_principal
    
    # 记录到原始df对应的行
    df.at[row.name, '有息本金'] = principal if rate > 0 else 0
    df.at[row.name, '无息本金'] = principal if rate == 0 else 0
    df.at[row.name, '期初余额'] = period_start_balance
    df.at[row.name, '期末余额'] = period_end_balance
    df.at[row.name, '利息'] = period_interest
    df.at[row.name, '未付利息'] = accrued_interest
    df.at[row.name, '还利息'] = repay_interest
    df.at[row.name, '还本金'] = repay_principal
    
    # 打印明细
    change = principal - repayment
    if change != 0 or period_interest > 0 or repay_interest > 0:
        if principal > 0:
            action = f"借入 {principal:,.2f}"
        elif repayment > 0:
            if repay_interest > 0 and repay_principal > 0:
                action = f"偿还 {repayment:,.2f}(利息{repay_interest:,.2f}+本金{repay_principal:,.2f})"
            elif repay_interest > 0:
                action = f"偿还 {repayment:,.2f}(还利息)"
            else:
                action = f"偿还 {repayment:,.2f}(还本金)"
        else:
            action = ""
        
        print(f"{current_date.strftime('%Y-%m-%d'):<12} {action:<24} {cumulative_interest_principal:>12,.2f} {cumulative_free_principal:>12,.2f} {period_interest:>10,.2f} {repay_interest:>10,.2f} {repay_principal:>10,.2f} {cumulative_interest_principal:>12,.2f} {accrued_interest:>10,.2f}")
        
        if period_interest > 0 or repay_interest > 0 or repay_principal > 0:
            interest_details.append({
                '日期': current_date.strftime('%Y-%m-%d'),
                '说明': action,
                '期初余额': period_start_balance,
                '计提利息': period_interest,
                '还利息': repay_interest,
                '还本金': repay_principal,
                '借款': principal,
                '期末余额': period_end_balance,
                '未付利息': accrued_interest
            })
    
    prev_date = current_date

# =============================================
# 计息至截止日（2020-08-19）
# =============================================
print(f"\n--- 计息至截止日 {CUTOFF_DATE.strftime('%Y-%m-%d')} ---")
if prev_date is not None and cumulative_interest_principal > 0:
    days_to_cutoff = (CUTOFF_DATE - prev_date).days
    if days_to_cutoff > 0:
        period_interest = cumulative_interest_principal * INTEREST_RATE * days_to_cutoff / 365
        accrued_interest += period_interest
        total_accrued_interest += period_interest
        print(f"{CUTOFF_DATE.strftime('%Y-%m-%d'):<12} {'计息至截止日':<24} {cumulative_interest_principal:>12,.2f} {cumulative_free_principal:>12,.2f} {period_interest:>10,.2f} {'':>10} {'':>10} {cumulative_interest_principal:>12,.2f} {accrued_interest:>10,.2f}")
        
        interest_details.append({
            '日期': CUTOFF_DATE.strftime('%Y-%m-%d'),
            '说明': f'计息至截止日({days_to_cutoff}天)',
            '期初余额': cumulative_interest_principal + cumulative_free_principal,
            '计提利息': period_interest,
            '还利息': 0,
            '还本金': 0,
            '借款': 0,
            '期末余额': cumulative_interest_principal + cumulative_free_principal,
            '未付利息': accrued_interest
        })

print(f"\n累计计提利息总额: {total_accrued_interest:,.2f} 元")
print(f"期末未付利息: {accrued_interest:,.2f} 元")
print(f"期末本金余额: {cumulative_interest_principal + cumulative_free_principal:,.2f} 元")

# 打印最终统计
print(f"\n{'='*80}")
print("计算结果汇总")
print("=" * 80)

no_interest_principal = no_interest_borrow['借'].sum()
interest_principal = interest_borrow['借'].sum()

print(f"\n【无息本金】: {no_interest_principal:,.2f} 元")
print(f"  - 共 {len(no_interest_borrow)} 笔借款未约定利息")

print(f"\n【有息本金】: {interest_principal:,.2f} 元")
print(f"  - 共 {len(interest_borrow)} 笔借款约定了利息")

print(f"\n【累计计提利息】: {total_accrued_interest:,.2f} 元")
print(f"  - 按先息后本摊余成本法计算至 {CUTOFF_DATE.strftime('%Y-%m-%d')}")

print(f"\n【期末未付利息】: {accrued_interest:,.2f} 元")

print(f"\n【借款本金总计】: {total_borrowed:,.2f} 元")
print(f"  - 无息本金 + 有息本金 = {no_interest_principal + interest_principal:,.2f} 元")

print(f"\n【还款总额】: {total_repaid:,.2f} 元")
print(f"  - 其中还利息: {sum([d['还利息'] for d in interest_details]):,.2f} 元")
print(f"  - 其中还本金: {sum([d['还本金'] for d in interest_details]):,.2f} 元")

print(f"\n【期末本金余额】: {cumulative_interest_principal + cumulative_free_principal:,.2f} 元")
total_end_principal = cumulative_interest_principal + cumulative_free_principal

# =============================================
# 保存结果
# =============================================
result_df = pd.DataFrame({
    '项目': ['无息本金', '有息本金', f'累计计提利息(至{CUTOFF_DATE.strftime("%Y-%m-%d")})', '期末未付利息', '借款本金总计', '还款总额', '其中还利息', '其中还本金', '期末本金余额'],
    '金额(元)': [no_interest_principal, interest_principal, total_accrued_interest, accrued_interest,
                total_borrowed, total_repaid, 
                sum([d['还利息'] for d in interest_details]),
                sum([d['还本金'] for d in interest_details]),
                total_end_principal],
    '说明': [
        f'共{len(no_interest_borrow)}笔借款未约定利息',
        f'共{len(interest_borrow)}笔借款约定了利息',
        f'按先息后本摊余成本法计算至{CUTOFF_DATE.strftime("%Y-%m-%d")}',
        '已计提但尚未支付的利息',
        '无息本金 + 有息本金',
        '所有还款金额总计',
        '无对应关系的还款优先还利息',
        '还利息后的剩余部分还本金',
        '借款本金 - 还本金'
    ]
})

# 利息明细
interest_detail_df = pd.DataFrame(interest_details)

# 选择要输出的列
output_columns = ['借款编号', '还款编号', '分类', '序号', '日期', '借', '还', '利率', '间隔天数', 
                  '还款期限/还款说明', '还款对应关系', '有息本金', '无息本金', '期初余额', '利息', '未付利息', '还利息', '还本金', '期末余额']
output_df = df[output_columns]

# 保存结果到Excel文件
output_file = 'D:/code/code/司法鉴定-还本+利息测算-计算结果_v3.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    result_df.to_excel(writer, sheet_name='汇总结果', index=False)
    interest_detail_df.to_excel(writer, sheet_name='利息明细', index=False)
    output_df.to_excel(writer, sheet_name='原始数据+对应关系', index=False)

print(f"\n计算结果已保存到: {output_file}")
print(f"  - '汇总结果' sheet: 汇总数据（计息截止日 {CUTOFF_DATE.strftime('%Y-%m-%d')}）")
print(f"  - '利息明细' sheet: 先息后本摊余成本法计算明细")
print(f"  - '原始数据+对应关系' sheet: 原始数据及追加列")
print(f"\n追加的列说明：")
print(f"  - '借款编号': 每笔借款的唯一编号（1-{borrow_count}）")
print(f"  - '还款编号': 每笔还款的唯一编号（1-{repay_count}）")
print(f"  - '有息本金': 该行借款中有利息的本金")
print(f"  - '无息本金': 该行借款中无利息的本金")
print(f"  - '期初余额': 该日期之前的累计借款本金余额")
print(f"  - '利息': 基于有息本金余额 × 利率(24%) × 天数 / 365 计提的利息")
print(f"  - '未付利息': 累计已计提但尚未支付的利息")
print(f"  - '还利息': 本次还款中用于支付利息的部分（先息后本）")
print(f"  - '还本金': 本次还款中用于还本金的部分")
print(f"  - '期末余额': 该日期之后的累计借款本金余额")
print(f"\n还款规则：")
print(f"  1. 有明确对应关系的还款 → 根据对应借款类型还本金")
print(f"  2. 无对应关系的还款 → 先息后本：先还利息，剩余还本金")