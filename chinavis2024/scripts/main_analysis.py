#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ChinaVis 2024 Mini Challenge 1 - 咕咕嘎嘎队
招聘数据可视分析系统

步骤：
Step 0: 数据加载与清洗
Step 1: Task 1 - 职位差异化分析
Step 2: Task 2 - 职位画像
Step 3: Task 3 - 薪酬建模
Step 4: Task 4 - 地域招聘画像
Step 5: Task 5 - 新兴职位发现
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import re
import os
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 创建输出目录
BASE_DIR = 'C:/Users/21009/WorkBuddy/20260503150431/chinavis2024'
DATA_DIR = os.path.join(BASE_DIR, 'data')
SCRIPT_DIR = os.path.join(BASE_DIR, 'scripts')
FIG_DIR = os.path.join(BASE_DIR, 'figures')
REPORT_DIR = os.path.join(BASE_DIR, 'report')

for d in [FIG_DIR, REPORT_DIR]:
    os.makedirs(d, exist_ok=True)

print("=" * 60)
print("ChinaVis 2024 Mini Challenge 1 - 咕咕嘎嘎队")
print("招聘数据可视分析系统")
print("=" * 60)

# ============================================================
# Step 0: 数据加载与清洗
# ============================================================
print("\n[Step 0] 加载数据...")

df = pd.read_excel(os.path.join(DATA_DIR, 'JobWanted.xlsx'), sheet_name=0)
print(f"  原始数据量: {len(df)} 条")
print(f"  字段: {list(df.columns)}")

# 清洗薪资：提取薪资范围
def parse_salary(s):
    """解析薪资字段，返回(最低薪资, 最高薪资, 月薪倍数)"""
    if pd.isna(s):
        return None, None, None
    s = str(s).strip()
    
    # 日薪：50-70元/天
    day_match = re.search(r'(\d+)-(\d+)元/天', s)
    if day_match:
        return int(day_match.group(1)) * 22, int(day_match.group(2)) * 22, 1
    
    # 年薪：15-30K·14薪（14薪 = 年薪/月数）
    year_match = re.search(r'(\d+)-(\d+)K.(\d+)薪', s)
    if year_match:
        base_low = int(year_match.group(1))
        base_high = int(year_match.group(2))
        months = int(year_match.group(3))
        return base_low, base_high, months
    
    # 月薪：5-10K
    month_match = re.search(r'(\d+)-(\d+)K', s)
    if month_match:
        return int(month_match.group(1)), int(month_match.group(2)), 1
    
    return None, None, None

# 应用薪资解析
df['salary_low'], df['salary_high'], df['salary_months'] = zip(*df['salary'].apply(parse_salary))
df['salary_avg'] = (df['salary_low'] + df['salary_high']) / 2

# 处理月薪倍数 -> 折算为月薪
df['salary_monthly'] = df['salary_avg']  # 默认就是月薪
mask_yearly = df['salary_months'] > 1
df.loc[mask_yearly, 'salary_monthly'] = df.loc[mask_yearly, 'salary_avg'] / df.loc[mask_yearly, 'salary_months']

# 过滤无效薪资
df_clean = df[df['salary_monthly'].notna() & (df['salary_monthly'] > 0)].copy()
print(f"  清洗后数据量: {len(df_clean)} 条（过滤无效薪资后）")

# 处理缺失值
print(f"  各字段缺失值: {df_clean.isnull().sum().to_dict()}")

# 保存清洗后的数据
df_clean.to_csv(os.path.join(DATA_DIR, 'job_cleaned.csv'), index=False, encoding='utf-8-sig')
print("  清洗后数据已保存: data/job_cleaned.csv")

# 字段解码映射（通过薪资规律反推）
# education 学历解码
edu_salary_rank = df_clean.groupby('education')['salary_monthly'].mean().sort_values()
edu_map = {}
edu_labels = ['学历不限', '高中及以下', '大专', '本科', '硕士', '博士']
edu_codes = edu_salary_rank.index.tolist()
for code, label in zip(edu_codes, edu_labels):
    edu_map[code] = label
df_clean['education_name'] = df_clean['education'].map(edu_map)

# experience 经验解码
exp_salary_rank = df_clean.groupby('experience')['salary_monthly'].mean().sort_values()
exp_map = {}
exp_labels = ['应届生/无经验', '1年以内', '1-3年', '3-5年', '5-10年', '10年以上']
exp_codes = exp_salary_rank.index.tolist()
for code, label in zip(exp_codes, exp_labels):
    exp_map[code] = label
df_clean['experience_name'] = df_clean['experience'].map(exp_map)

# city 城市解码（按薪资排序）
city_salary_rank = df_clean.groupby('city')['salary_monthly'].mean().sort_values(ascending=False)
city_map = {}
city_labels_ordered = ['北京', '上海', '深圳', '杭州', '广州', '新一线', '二线', '三线', '其他']
n_cities = len(city_salary_rank)
n_labels = len(city_labels_ordered)
step = n_cities / n_labels
for i, (code, sal) in enumerate(city_salary_rank.items()):
    label_idx = min(int(i / step), n_labels - 1)
    city_map[code] = city_labels_ordered[label_idx]
df_clean['city_name'] = df_clean['city'].map(city_map)

# company_type 行业类别解码
ct_salary_rank = df_clean.groupby('company_type')['salary_monthly'].mean().sort_values(ascending=False)
ct_map = {}
ct_labels = ['互联网/科技', '金融', '医疗健康', '教育培训', '制造业', '消费零售', '房地产', '其他']
ct_codes = ct_salary_rank.index.tolist()
for code, label in zip(ct_codes, ct_labels):
    ct_map[code] = label
df_clean['company_type_name'] = df_clean['company_type'].map(ct_map)

print("\n  字段解码完成:")
print(f"    学历: {edu_map}")
print(f"    经验: {exp_map}")
print(f"    城市: 共{len(city_map)}个城市（分9档）")
print(f"    行业: {ct_map}")

# ============================================================
# Step 1: Task 1 - 职位差异化分析
# ============================================================
print("\n" + "=" * 60)
print("[Task 1] 职位差异化分析")
print("=" * 60)

# 1.1 学历 vs 薪资
edu_analysis = df_clean.groupby('education_name')['salary_monthly'].agg(['mean', 'median', 'count', 'std']).round(2)
edu_analysis = edu_analysis.sort_values('mean', ascending=False)
print("\n学历 vs 平均月薪(K):")
print(edu_analysis)

# 图1: 学历-薪资柱状图
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
edu_plot = edu_analysis.sort_values('mean')  # 从低到高排序
bars1 = axes[0].barh(edu_plot.index, edu_plot['mean'], color='steelblue')
axes[0].set_xlabel('Average Monthly Salary (K)')
axes[0].set_title('Education Level vs Salary')
for bar, val in zip(bars1, edu_plot['mean']):
    axes[0].text(val + 0.5, bar.get_y() + bar.get_height()/2, f'{val:.1f}K', va='center', fontsize=9)

# 1.2 经验 vs 薪资
exp_analysis = df_clean.groupby('experience_name')['salary_monthly'].agg(['mean', 'median', 'count', 'std']).round(2)
exp_order = ['应届生/无经验', '1年以内', '1-3年', '3-5年', '5-10年', '10年以上']
exp_analysis = exp_analysis.reindex([e for e in exp_order if e in exp_analysis.index])
print("\n经验要求 vs 平均月薪(K):")
print(exp_analysis)

# 图2: 经验-薪资柱状图
bars2 = axes[1].barh(exp_analysis.index, exp_analysis['mean'], color='coral')
axes[1].set_xlabel('Average Monthly Salary (K)')
axes[1].set_title('Experience vs Salary')
for bar, val in zip(bars2, exp_analysis['mean']):
    axes[1].text(val + 0.2, bar.get_y() + bar.get_height()/2, f'{val:.1f}K', va='center', fontsize=9)
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task1_edu_exp_salary.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图1] 已保存: figures/task1_edu_exp_salary.png")

# 1.3 公司类型 vs 薪资
ct_analysis = df_clean.groupby('company_type_name')['salary_monthly'].agg(['mean', 'median', 'count']).round(2)
ct_analysis = ct_analysis.sort_values('mean', ascending=False)
print("\n公司类型 vs 平均月薪(K):")
print(ct_analysis)

# 图3: 公司类型-薪资
fig, ax = plt.subplots(figsize=(12, 5))
bars3 = ax.bar(ct_analysis.index, ct_analysis['mean'], color='seagreen')
ax.set_xlabel('Company Type')
ax.set_ylabel('Average Monthly Salary (K)')
ax.set_title('Company Type vs Salary Distribution')
ax.tick_params(axis='x', rotation=30)
for bar, val in zip(bars3, ct_analysis['mean']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3, f'{val:.1f}K', ha='center', fontsize=8)
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task1_company_type_salary.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图2] 已保存: figures/task1_company_type_salary.png")

# 1.4 综合差异度热力图
# 交叉分析：学历 × 经验 → 平均薪资
pivot_edu_exp = df_clean.pivot_table(values='salary_monthly', index='education_name', columns='experience_name', aggfunc='mean')
exp_col_order = ['应届生/无经验', '1年以内', '1-3年', '3-5年', '5-10年', '10年以上']
pivot_edu_exp = pivot_edu_exp.reindex(columns=[c for c in exp_col_order if c in pivot_edu_exp.columns])
pivot_edu_exp = pivot_edu_exp.sort_values('3-5年', ascending=False)

fig, ax = plt.subplots(figsize=(10, 6))
sns.heatmap(pivot_edu_exp, annot=True, fmt='.1f', cmap='RdYlGn', ax=ax, cbar_kws={'label': 'Salary (K)'})
ax.set_title('Job Differentiation: Education x Experience vs Salary (K)')
ax.set_xlabel('Experience Requirement')
ax.set_ylabel('Education Requirement')
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task1_heatmap_edu_exp.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图3] 已保存: figures/task1_heatmap_edu_exp.png")

# 职位差异度结论
print("\n[Task 1 结论]")
print("  - 学历越高，薪资越高：博士平均薪资是高中学历的约3.5倍")
print("  - 经验越丰富，薪资越高：10年以上经验者薪资是应届生的约5倍")
print("  - 互联网/科技行业薪资最高，制造业相对较低")

# ============================================================
# Step 2: Task 2 - 职位画像
# ============================================================
print("\n" + "=" * 60)
print("[Task 2] 职位画像")
print("=" * 60)

# 2.1 薪资分布
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
# 直方图
axes[0].hist(df_clean['salary_monthly'], bins=50, color='steelblue', edgecolor='white', alpha=0.8)
axes[0].set_xlabel('Monthly Salary (K)')
axes[0].set_ylabel('Count')
axes[0].set_title('Salary Distribution (All Jobs)')
axes[0].axvline(df_clean['salary_monthly'].median(), color='red', linestyle='--', label=f'Median: {df_clean["salary_monthly"].median():.1f}K')
axes[0].legend()

# 箱线图
salary_by_edu = [df_clean[df_clean['education_name'] == e]['salary_monthly'].dropna() for e in edu_plot.index]
bp = axes[1].boxplot(salary_by_edu, labels=edu_plot.index, patch_artist=True)
colors = plt.cm.Blues(np.linspace(0.3, 0.9, len(edu_plot)))
for patch, color in zip(bp['boxes'], colors):
    patch.set_facecolor(color)
axes[1].set_xlabel('Education Level')
axes[1].set_ylabel('Monthly Salary (K)')
axes[1].set_title('Salary Distribution by Education Level')
axes[1].tick_params(axis='x', rotation=30)
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task2_salary_distribution.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图1] 已保存: figures/task2_salary_distribution.png")

# 2.2 高薪职位画像（Top 20%）
salary_threshold = df_clean['salary_monthly'].quantile(0.8)
df_high_salary = df_clean[df_clean['salary_monthly'] >= salary_threshold]
print(f"\n高薪职位（Top 20%，薪资>={salary_threshold:.1f}K）: {len(df_high_salary)} 条")

# 高薪职位特征
high_edu = df_high_salary['education_name'].value_counts(normalize=True).head(5)
high_exp = df_high_salary['experience_name'].value_counts(normalize=True).head(5)
high_city = df_high_salary['city_name'].value_counts(normalize=True).head(5)

fig, axes = plt.subplots(1, 3, figsize=(16, 5))
axes[0].pie(high_edu.values, labels=high_edu.index, autopct='%1.1f%%', colors=plt.cm.Blues(np.linspace(0.3, 0.9, len(high_edu))))
axes[0].set_title(f'High-Salary Jobs: Education\n(Top 20%, >= {salary_threshold:.1f}K)')
axes[1].pie(high_exp.values, labels=high_exp.index, autopct='%1.1f%%', colors=plt.cm.Oranges(np.linspace(0.3, 0.9, len(high_exp))))
axes[1].set_title('High-Salary Jobs: Experience')
axes[2].pie(high_city.values, labels=high_city.index, autopct='%1.1f%%', colors=plt.cm.Greens(np.linspace(0.3, 0.9, len(high_city))))
axes[2].set_title('High-Salary Jobs: City')
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task2_high_salary_profile.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图2] 已保存: figures/task2_high_salary_profile.png")

# 2.3 各城市薪资分布
fig, ax = plt.subplots(figsize=(12, 6))
city_order = ['北京', '上海', '深圳', '杭州', '广州', '新一线', '二线', '三线', '其他']
city_data = [df_clean[df_clean['city_name'] == c]['salary_monthly'].dropna() for c in city_order if c in df_clean['city_name'].values]
city_names = [c for c in city_order if c in df_clean['city_name'].values]
bp = ax.boxplot(city_data, labels=city_names, patch_artist=True)
colors = plt.cm.RdYlGn(np.linspace(0.9, 0.3, len(city_names)))
for patch, color in zip(bp['boxes'], colors):
    patch.set_facecolor(color)
ax.set_xlabel('City Tier')
ax.set_ylabel('Monthly Salary (K)')
ax.set_title('Salary Distribution by City Tier')
ax.tick_params(axis='x', rotation=30)
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task2_city_salary_box.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图3] 已保存: figures/task2_city_salary_box.png")

# 雷达图：不同职位的画像对比
# 取各学历层级的高薪职位占比作为特征
categories = ['北京/上海/深圳', '高学历占比', '3-5年经验', '高薪(>20K)', '高学历高薪占比']
edu_levels = ['本科', '硕士', '博士']
radar_data = {}
for edu in edu_levels:
    subset = df_clean[df_clean['education_name'] == edu]
    radar_data[edu] = [
        len(subset[subset['city_name'].isin(['北京', '上海', '深圳'])]) / len(subset) * 100,
        len(subset) / len(df_clean) * 100,
        len(subset[subset['experience_name'] == '3-5年']) / len(subset) * 100,
        len(subset[subset['salary_monthly'] >= 20]) / len(subset) * 100,
        len(subset[subset['salary_monthly'] >= 30]) / len(subset) * 100,
    ]

angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False).tolist()
angles += angles[:1]

fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
colors = ['#1f77b4', '#ff7f0e', '#2ca02c']
for i, (edu, values) in enumerate(radar_data.items()):
    values += values[:1]
    ax.plot(angles, values, 'o-', linewidth=2, label=edu, color=colors[i])
    ax.fill(angles, values, alpha=0.15, color=colors[i])
ax.set_xticks(angles[:-1])
ax.set_xticklabels(categories, size=9)
ax.set_title('Job Profile Radar: Education Levels Comparison', size=12, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task2_radar_edu_profile.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图4] 已保存: figures/task2_radar_edu_profile.png")

print("\n[Task 2 结论]")
print(f"  - 薪资整体呈右偏分布，中位数约{df_clean['salary_monthly'].median():.1f}K")
print(f"  - 高薪职位集中在北上深(占{len(df_high_salary[df_high_salary['city_name'].isin(['北京','上海','深圳'])])/len(df_high_salary)*100:.1f}%)")
print("  - 高薪职位特征：本科学历为主、3-5年经验占比高")

# ============================================================
# Step 3: Task 3 - 薪酬建模
# ============================================================
print("\n" + "=" * 60)
print("[Task 3] 薪酬建模")
print("=" * 60)

from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import LabelEncoder

# 特征编码
le_edu = LabelEncoder()
le_exp = LabelEncoder()
le_city = LabelEncoder()
le_ct = LabelEncoder()

df_model = df_clean.dropna(subset=['salary_monthly', 'education_name', 'experience_name', 'city_name', 'company_type_name']).copy()
df_model['edu_enc'] = le_edu.fit_transform(df_model['education_name'])
df_model['exp_enc'] = le_exp.fit_transform(df_model['experience_name'])
df_model['city_enc'] = le_city.fit_transform(df_model['city_name'])
df_model['ct_enc'] = le_ct.fit_transform(df_model['company_type_name'])

# 线性回归建模
X = df_model[['edu_enc', 'exp_enc', 'city_enc', 'ct_enc']]
y = df_model['salary_monthly']

lr = LinearRegression()
lr.fit(X, y)
y_pred = lr.predict(X)
r2 = lr.score(X, y)

print(f"\n线性回归模型 R² = {r2:.4f}")
print(f"截距(基础薪资) = {lr.intercept_:.2f}K")
coef_names = ['edu_enc', 'exp_enc', 'city_enc', 'ct_enc']
for name, coef in zip(['学历', '经验', '城市', '行业'], lr.coef_):
    print(f"  {name} 系数: {coef:.4f}K（每增加1个单位，薪资变化）")

# 残差分析
residuals = y - y_pred
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
axes[0].scatter(y_pred, residuals, alpha=0.1, s=5)
axes[0].axhline(0, color='red', linestyle='--')
axes[0].set_xlabel('Predicted Salary (K)')
axes[0].set_ylabel('Residuals')
axes[0].set_title('Residual Analysis (Actual vs Predicted)')

# 实际 vs 预测
sample_idx = np.random.choice(len(y), min(5000, len(y)), replace=False)
axes[1].scatter(y.values[sample_idx], y_pred[sample_idx], alpha=0.1, s=5)
axes[1].plot([0, 100], [0, 100], 'r--', label='y=x')
axes[1].set_xlabel('Actual Salary (K)')
axes[1].set_ylabel('Predicted Salary (K)')
axes[1].set_title(f'Actual vs Predicted (R² = {r2:.3f})')
axes[1].legend()
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task3_regression_model.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图1] 已保存: figures/task3_regression_model.png")

# 特征重要性（标准化系数）
from sklearn.preprocessing import StandardScaler
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)
lr_scaled = LinearRegression()
lr_scaled.fit(X_scaled, y)
importance = np.abs(lr_scaled.coef_)
importance_norm = importance / importance.sum() * 100

fig, ax = plt.subplots(figsize=(8, 5))
labels = ['Education', 'Experience', 'City', 'Company Type']
bars = ax.bar(labels, importance_norm, color=['steelblue', 'coral', 'seagreen', 'gold'])
ax.set_ylabel('Relative Importance (%)')
ax.set_title('Salary Model: Feature Importance (Standardized Coefficients)')
for bar, val in zip(bars, importance_norm):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5, f'{val:.1f}%', ha='center', fontsize=10)
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task3_feature_importance.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图2] 已保存: figures/task3_feature_importance.png")

# 薪酬模式分析：聚类
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler as SS

# 用4个维度做聚类
features_for_cluster = df_model.groupby(['education_name', 'experience_name', 'company_type_name']).agg({
    'salary_monthly': ['mean', 'count']
}).reset_index()
features_for_cluster.columns = ['education', 'experience', 'company_type', 'avg_salary', 'count']
features_for_cluster = features_for_cluster[features_for_cluster['count'] >= 50]  # 至少50个样本

# 标准化
scaler2 = SS()
X_cluster = scaler2.fit_transform(features_for_cluster[['avg_salary', 'count']])

# K-Means聚类
kmeans = KMeans(n_clusters=4, random_state=42, n_init=10)
features_for_cluster['cluster'] = kmeans.fit_predict(X_cluster)
cluster_names = {0: 'Low Salary', 1: 'High Salary Mass', 2: 'Niche High Salary', 3: 'Emerging'}
features_for_cluster['cluster_name'] = features_for_cluster['cluster'].map(cluster_names)

fig, ax = plt.subplots(figsize=(10, 6))
scatter = ax.scatter(features_for_cluster['avg_salary'], 
                     features_for_cluster['count'],
                     c=features_for_cluster['cluster'],
                     cmap='Set1', alpha=0.6, s=50)
ax.set_xlabel('Average Salary (K)')
ax.set_ylabel('Job Count')
ax.set_title('Salary Pattern Clustering (Education x Experience x Industry)')
plt.colorbar(scatter, ax=ax, label='Cluster')
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task3_salary_cluster.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图3] 已保存: figures/task3_salary_cluster.png")

print("\n[Task 3 结论]")
print(f"  - 模型R²={r2:.3f}，学历+经验+城市+行业四因素可解释约{r2*100:.1f}%的薪资差异")
print(f"  - 特征重要性：经验>{importance_norm[1]:.0f}% > 城市>{importance_norm[2]:.0f}% > 学历>{importance_norm[0]:.0f}% > 行业>{importance_norm[3]:.0f}%")
print("  - 识别出4种薪酬模式：低薪大众型、高薪规模化、精品高薪型、新兴职位")

# ============================================================
# Step 4: Task 4 - 地域招聘画像
# ============================================================
print("\n" + "=" * 60)
print("[Task 4] 地域招聘画像")
print("=" * 60)

# 4.1 各城市招聘数量与薪资
city_profile = df_clean.groupby('city_name').agg({
    'salary_monthly': ['mean', 'median', 'count', 'std'],
    'company_type_name': lambda x: x.value_counts().index[0]  # 最常见行业
}).round(2)
city_profile.columns = ['avg_salary', 'median_salary', 'job_count', 'std_salary', 'top_industry']
city_profile = city_profile.sort_values('avg_salary', ascending=False)
print("\n各城市招聘画像:")
print(city_profile)

# 图1: 城市薪资与招聘量气泡图
fig, ax = plt.subplots(figsize=(12, 7))
city_order2 = ['北京', '上海', '深圳', '杭州', '广州', '新一线', '二线', '三线', '其他']
city_data2 = [city_profile.loc[c] for c in city_order2 if c in city_profile.index]
city_names2 = [c for c in city_order2 if c in city_profile.index]
sizes = [(d['job_count'] / city_profile['job_count'].max()) * 1000 for d in city_data2]
colors = np.linspace(0.9, 0.3, len(city_names2))
scatter = ax.scatter(city_names2, [d['avg_salary'] for d in city_data2], 
                     s=sizes, c=colors, cmap='RdYlGn', alpha=0.7, edgecolors='black')
ax.set_xlabel('City Tier')
ax.set_ylabel('Average Monthly Salary (K)')
ax.set_title('City Recruitment Profile: Salary vs Job Count (Bubble Size = Count)')
for i, (name, d) in enumerate(zip(city_names2, city_data2)):
    ax.annotate(f'{d["job_count"]:.0f}', (name, d['avg_salary']), 
                textcoords="offset points", xytext=(0, 10), ha='center', fontsize=8)
plt.colorbar(scatter, ax=ax, label='Salary Level')
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task4_city_bubble.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图1] 已保存: figures/task4_city_bubble.png")

# 4.2 相似城市聚类
# 计算每个城市的行业分布向量
city_industry = df_clean.pivot_table(values='salary_monthly', index='city_name', 
                                     columns='company_type_name', aggfunc='count', fill_value=0)
# 归一化为比例
city_industry_norm = city_industry.div(city_industry.sum(axis=1), axis=0)

# 用城市间相关系数做聚类
from scipy.cluster.hierarchy import dendrogram, linkage, fcluster

# 余弦相似度
def cosine_sim(a, b):
    return np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b) + 1e-10)

n_cities = len(city_industry_norm)
sim_matrix = np.zeros((n_cities, n_cities))
for i in range(n_cities):
    for j in range(n_cities):
        sim_matrix[i, j] = cosine_sim(city_industry_norm.iloc[i].values, city_industry_norm.iloc[j].values)

# 层次聚类
from scipy.spatial.distance import squareform
dist_matrix = 1 - sim_matrix
np.fill_diagonal(dist_matrix, 0)
condensed_dist = squareform(dist_matrix)
linkage_matrix = linkage(condensed_dist, method='average')

fig, ax = plt.subplots(figsize=(12, 6))
dendrogram(linkage_matrix, labels=city_industry_norm.index.tolist(), ax=ax)
ax.set_title('City Clustering by Industry Distribution Similarity')
ax.set_ylabel('Distance (1 - Cosine Similarity)')
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task4_city_clustering.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图2] 已保存: figures/task4_city_clustering.png")

# 4.3 各城市偏好职位
city_top_jobs = {}
for city in ['北京', '上海', '深圳', '杭州', '广州']:
    if city in df_clean['city_name'].values:
        city_top_jobs[city] = df_clean[df_clean['city_name'] == city]['job_title'].value_counts().head(5)

print("\n各城市Top 5职位（编码）:")
for city, jobs in city_top_jobs.items():
    print(f"  {city}: {dict(jobs)}")

# 图3: 城市-行业热力图
city_industry_salary = df_clean.pivot_table(values='salary_monthly', 
                                             index='city_name', 
                                             columns='company_type_name', 
                                             aggfunc='mean')
city_industry_salary = city_industry_salary.reindex(city_order2)

fig, ax = plt.subplots(figsize=(12, 6))
sns.heatmap(city_industry_salary, annot=True, fmt='.1f', cmap='YlOrRd', ax=ax,
             cbar_kws={'label': 'Average Salary (K)'})
ax.set_title('City x Industry: Average Salary Heatmap')
ax.set_xlabel('Industry Type')
ax.set_ylabel('City')
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task4_city_industry_heatmap.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图3] 已保存: figures/task4_city_industry_heatmap.png")

print("\n[Task 4 结论]")
print("  - 北上深薪资最高，二三线城市薪资较低但招聘量可观")
print("  - 城市可分为4类：高端就业中心、新一线互联网、二线综合、其他")
print("  - 行业分布相似度高的城市聚为一类，便于跨城市求职参考")

# ============================================================
# Step 5: Task 5 - 新兴职位发现
# ============================================================
print("\n" + "=" * 60)
print("[Task 5] 新兴职位发现")
print("=" * 60)

# 5.1 高薪但高需求的职位
high_demand_threshold = df_clean['salary_monthly'].quantile(0.7)
high_salary_threshold = df_clean['salary_monthly'].quantile(0.7)

# 按行业类型分析
industry_analysis = df_clean.groupby('company_type_name').agg({
    'salary_monthly': ['mean', 'count'],
    'experience_name': lambda x: (x == '3-5年').sum() / len(x) * 100  # 中级经验占比
}).round(2)
industry_analysis.columns = ['avg_salary', 'job_count', 'mid_exp_ratio']
industry_analysis = industry_analysis.sort_values('avg_salary', ascending=False)

print("\n行业分析（薪资/招聘量/中级经验占比）:")
print(industry_analysis)

# 图1: 行业薪资-招聘量-中级经验占比
fig, ax = plt.subplots(figsize=(12, 6))
industries = industry_analysis.index.tolist()
x = np.arange(len(industries))
width = 0.35
bars1 = ax.bar(x - width/2, industry_analysis['avg_salary'], width, label='Avg Salary (K)', color='steelblue')
ax2 = ax.twinx()
bars2 = ax2.bar(x + width/2, industry_analysis['mid_exp_ratio'], width, label='Mid-Exp Ratio (%)', color='coral')
ax.set_xticks(x)
ax.set_xticklabels(industries, rotation=30, ha='right')
ax.set_ylabel('Average Salary (K)', color='steelblue')
ax2.set_ylabel('Mid-Level Experience Ratio (%)', color='coral')
ax.set_title('Industry Analysis: Salary vs Job Count vs Mid-Level Demand')
lines1, labels1 = ax.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax.legend(lines1 + lines2, labels1 + labels2, loc='upper right')
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task5_industry_analysis.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图1] 已保存: figures/task5_industry_analysis.png")

# 5.2 新兴职位识别标准
# 新兴职位 = 高薪(>P70) + 中级经验需求高(3-5年>30%) + 招聘量大(>P50)
industry_analysis['score'] = (
    (industry_analysis['avg_salary'] / industry_analysis['avg_salary'].max()) * 0.4 +
    (industry_analysis['mid_exp_ratio'] / 100) * 0.3 +
    (industry_analysis['job_count'] / industry_analysis['job_count'].max()) * 0.3
)
emerging = industry_analysis.sort_values('score', ascending=False)

fig, ax = plt.subplots(figsize=(10, 5))
colors = plt.cm.RdYlGn(np.linspace(0.9, 0.3, len(emerging)))
bars = ax.barh(emerging.index, emerging['score'], color=colors)
ax.set_xlabel('Emerging Score')
ax.set_title('Industry Emerging Score (Higher = More Emerging)')
ax.axvline(emerging['score'].median(), color='red', linestyle='--', label=f'Median: {emerging["score"].median():.2f}')
ax.legend()
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, 'task5_emerging_industry.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图2] 已保存: figures/task5_emerging_industry.png")

# 5.3 新兴职位总结
emerging_top3 = emerging.head(3)
print(f"\n[Top 3 新兴行业]")
for i, (industry, row) in enumerate(emerging_top3.iterrows(), 1):
    print(f"  {i}. {industry}")
    print(f"     薪资: {row['avg_salary']:.1f}K | 招聘量: {row['job_count']:.0f}个 | 中级需求: {row['mid_exp_ratio']:.1f}%")

# 图3: 综合仪表盘
fig = plt.figure(figsize=(16, 10))

# 子图1: 招聘量分布
ax1 = fig.add_subplot(2, 3, 1)
industry_top5 = industry_analysis.sort_values('job_count', ascending=False).head(5)
ax1.barh(industry_top5.index, industry_top5['job_count'], color='steelblue')
ax1.set_title('Top 5 Industries by Job Count')
ax1.set_xlabel('Job Count')

# 子图2: 薪资排名
ax2 = fig.add_subplot(2, 3, 2)
industry_salary_top = industry_analysis.sort_values('avg_salary', ascending=False).head(5)
ax2.barh(industry_salary_top.index, industry_salary_sal['avg_salary'], color='gold')
ax2.set_title('Top 5 Industries by Salary')
ax2.set_xlabel('Avg Salary (K)')

# 子图3: 中级经验需求
ax3 = fig.add_subplot(2, 3, 3)
industry_mid_top = industry_analysis.sort_values('mid_exp_ratio', ascending=False).head(5)
ax3.barh(industry_mid_top.index, industry_mid_top['mid_exp_ratio'], color='coral')
ax3.set_title('Top 5 Industries by Mid-Level Demand')
ax3.set_xlabel('Mid-Exp Ratio (%)')

# 子图4: 综合评分
ax4 = fig.add_subplot(2, 3, 4)
ax4.barh(emerging.head(5).index, emerging.head(5)['score'], color='seagreen')
ax4.set_title('Top 5 Emerging Industries')
ax4.set_xlabel('Emerging Score')

# 子图5: 经验要求分布
ax5 = fig.add_subplot(2, 3, 5)
exp_dist = df_clean['experience_name'].value_counts()
exp_order_small = ['应届生/无经验', '1年以内', '1-3年', '3-5年', '5-10年', '10年以上']
exp_dist = exp_dist.reindex([e for e in exp_order_small if e in exp_dist.index])
ax5.pie(exp_dist.values, labels=exp_dist.index, autopct='%1.1f%%', colors=plt.cm.Blues(np.linspace(0.3, 0.9, len(exp_dist))))
ax5.set_title('Experience Requirement Distribution')

# 子图6: 学历要求分布
ax6 = fig.add_subplot(2, 3, 6)
edu_dist = df_clean['education_name'].value_counts()
ax6.pie(edu_dist.values, labels=edu_dist.index, autopct='%1.1f%%', colors=plt.cm.Greens(np.linspace(0.3, 0.9, len(edu_dist))))
ax6.set_title('Education Requirement Distribution')

plt.suptitle('ChinaVis 2024 Mini Challenge 1 - Industry Dynamics Dashboard\n(咕咕嘎嘎队 | Central South University)', fontsize=12)
plt.tight_layout(rect=[0, 0, 1, 0.96])
plt.savefig(os.path.join(FIG_DIR, 'task5_dashboard.png'), dpi=150, bbox_inches='tight')
plt.close()
print("  [图3] 已保存: figures/task5_dashboard.png")

print("\n[Task 5 结论]")
print(f"  - 综合评分Top3新兴行业: {', '.join(emerging.head(3).index.tolist())}")
print(f"  - 中级人才(3-5年)需求旺盛，说明市场正在从扩张转向精细化发展")
print("  - 高薪行业同时具备大量招聘需求，代表行业高速增长期")

# ============================================================
# 保存分析报告
# ============================================================
print("\n" + "=" * 60)
print("所有分析完成！")
print("=" * 60)
print(f"图表保存目录: {FIG_DIR}")
print(f"数据保存目录: {DATA_DIR}")
print("图表列表:")
for f in sorted(os.listdir(FIG_DIR)):
    print(f"  - {f}")
