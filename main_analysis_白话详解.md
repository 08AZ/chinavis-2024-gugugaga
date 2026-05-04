# main_analysis.py 白话详解文档
## 🎯 这个文件在干嘛？

**一句话概括**：读取43万条招聘数据，清洗→分析→画图→建模，完成ChinaVis竞赛的5个任务。

---

## 📚 第一部分：pandas 常用方法速查表

在讲代码前，先给你一个"pandas字典"，后面看到不懂的方法就来查！

### 🔧 数据读取与保存

| 代码 | 白话解释 |
|------|----------|
| `pd.read_excel(路径)` | 读取Excel文件，返回一个DataFrame（像Excel表格一样的数据结构） |
| `pd.read_csv(路径)` | 读取CSV文件 |
| `df.to_csv(路径, index=False)` | 把DataFrame保存成CSV文件，`index=False`表示不保存行号 |

### 🔍 数据查看

| 代码 | 白话解释 |
|------|----------|
| `df.head(5)` | 看前5行数据（默认5行，括号里可以改数字） |
| `df.tail(3)` | 看后3行 |
| `df.shape` | 返回 `(行数, 列数)`，比如 `(430664, 7)` |
| `df.columns` | 查看所有列名（字段名） |
| `df.info()` | 显示数据基本信息：列名、数据类型、非空值数量 |
| `df.describe()` | 统计描述：均值、标准差、最小值、四分位数、最大值 |

### 🎯 数据筛选

| 代码 | 白话解释 |
|------|----------|
| `df['薪资'] > 10` | 返回一个True/False的列表，哪些是薪资>10K的 |
| `df[df['薪资'] > 10]` | 筛选出薪资>10K的所有行 |
| `df[['姓名', '薪资']]` | 只显示"姓名"和"薪资"两列 |
| `df['薪资'].notna()` | 找出薪资不是空值的行 |

### 📊 数据分组与统计

| 代码 | 白话解释 |
|------|----------|
| `df.groupby('学历')` | 按"学历"分组，相同学历的放一起 |
| `df.groupby('学历')['薪资'].mean()` | 按学历分组后，计算每个学历的平均薪资 |
| `df.groupby('学历')['薪资'].agg(['mean', 'median', 'count'])` | 一次性计算多个指标：平均值、中位数、数量 |
| `df.value_counts()` | 统计每个值出现多少次，比如每种学历有多少人 |
| `df.pivot_table()` | 制作透视表（类似Excel的数据透视表） |

### 🔧 数据清洗与转换

| 代码 | 白话解释 |
|------|----------|
| `df.isnull().sum()` | 统计每列有多少个空值 |
| `df.dropna()` | 删除有空值的行 |
| `df.fillna(0)` | 把空值填充为0 |
| `df['新列'] = ...` | 创建新的一列 |
| `df['字段'].map(字典)` | 映射替换，比如把'A'变成'本科' |
| `df.apply(函数名)` | 对每一行/每一列应用某个函数 |

### 📈 排序与索引

| 代码 | 白话解释 |
|------|----------|
| `df.sort_values('薪资')` | 按薪资从小到大排序 |
| `df.sort_values('薪资', ascending=False)` | 按薪资从大到小排序 |
| `df.reset_index(drop=True)` | 重新排排行号（0,1,2,3...） |

---

## 📖 第二部分：代码逐段详解

---

### 🔧 Step 0：数据加载与清洗（第44-144行）

#### 第49行：读取Excel
```python
df = pd.read_excel(os.path.join(DATA_DIR, 'JobWanted.xlsx'), sheet_name=0)
```
**白话**：把Excel文件读进来，存到变量 `df` 里，`sheet_name=0` 表示读第一个工作表。

#### 第50-51行：看看数据长啥样
```python
print(f"  原始数据量: {len(df)} 条")
print(f"  字段: {list(df.columns)}")
```
**白话**：`len(df)` 就是多少行，`df.columns` 就是所有列名。

#### 第54-78行：解析薪资字段
```python
def parse_salary(s):
    # 这个函数把 "5-10K" 这样的字符串拆成 最低5K、最高10K
```
**为什么要这样写？**
- 原始数据里薪资是字符串，比如 "5-10K"、"15-30K·14薪"
- 电脑不认识字符串，我们要拆成数字才能计算

**正则表达式简单解释**：
- `r'(\d+)-(\d+)K'` → 匹配 "5-10K"，`\d+` 表示1个或多个数字
- `re.search(模式, 字符串)` → 在字符串里找符合模式的部分

#### 第81-82行：应用函数到整列
```python
df['salary_low'], df['salary_high'], df['salary_months'] = zip(*df['salary'].apply(parse_salary))
df['salary_avg'] = (df['salary_low'] + df['salary_high']) / 2
```
**白话**：
- `df['salary'].apply(parse_salary)` → 对"薪资"这一列的每一行，都执行 `parse_salary` 函数
- `zip(*...)` → 把函数返回的3个值（最低、最高、月数）拆成3列
- 新增了3列：`salary_low`、`salary_high`、`salary_months`

#### 第90行：过滤无效数据
```python
df_clean = df[df['salary_monthly'].notna() & (df['salary_monthly'] > 0)].copy()
```
**白话**：
- `df['salary_monthly'].notna()` → 薪资不是空值的
- `&` → 并且
- `(df['salary_monthly'] > 0)` → 薪资大于0的
- 合起来：只要薪资有效且大于0的行
- `.copy()` → 复制一份新的，避免修改原始数据

#### 第97行：保存清洗后的数据
```python
df_clean.to_csv(os.path.join(DATA_DIR, 'job_cleaned.csv'), index=False, encoding='utf-8-sig')
```
**白话**：把清洗后的数据保存成CSV，`utf-8-sig` 是避免中文乱码。

#### 第102-117行：字段解码（重点！）
```python
edu_salary_rank = df_clean.groupby('education')['salary_monthly'].mean().sort_values()
```
**这行在干嘛？**
1. `df_clean.groupby('education')` → 按学历分组
2. `['salary_monthly'].mean()` → 计算每个学历的平均薪资
3. `.sort_values()` → 按平均薪资从小到大排序

**结果举例**：
```
Gh   5K   (学历代码Gh的平均薪资是5K)
Gy   5K
Go   6K
...
GJ   151K  (学历代码GJ的平均薪资是151K)
```

**为什么要这样做？**
- 官方数据里 `education` 是虚拟代码（Gh、GJ等），没有告诉我们对应什么学历
- 但薪资高低和学历高低是相关的（博士薪资 > 本科薪资）
- 所以我们**通过薪资推断出**：薪资最低的代码对应"高中"，最高的对应"博士"

**第106-107行：建立映射字典**
```python
edu_labels = ['学历不限', '高中及以下', '大专', '本科', '硕士', '博士']
for code, label in zip(edu_codes, edu_labels):
    edu_map[code] = label
```
**白话**：把排序后的学历代码，按顺序对应到学历标签。

**第108行：应用映射到新列**
```python
df_clean['education_name'] = df_clean['education'].map(edu_map)
```
**白话**：在 `df_clean` 里新增一列 `education_name`，把 `education` 列的虚拟代码替换成"本科"、"硕士"等。

---

### 📊 Step 1：Task 1 - 职位差异化分析（第146-228行）

#### 第154行：按学历统计薪资
```python
edu_analysis = df_clean.groupby('education_name')['salary_monthly'].agg(['mean', 'median', 'count', 'std']).round(2)
```
**白话**：
- 按 `education_name`（学历）分组
- 对 `salary_monthly`（月薪）计算4个指标：
  - `mean`：平均值
  - `median`：中位数
  - `count`：数量
  - `std`：标准差（衡量薪资分散程度）
- `.round(2)` → 保留2位小数

**结果示例**：
```
             mean  median  count   std
学历不限      8.5    8.0    5000   3.2
大专         10.2   9.5    8000   4.1
本科         15.6   14.0   20000  6.8
硕士         22.3   20.0   12000  8.5
博士         35.1   32.0    3000  12.3
```

#### 第160-183行：画柱状图
```python
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
```
**白话**：创建1行2列的画布，用来画两个图。

```python
bars1 = axes[0].barh(edu_plot.index, edu_plot['mean'], color='steelblue')
```
**白话**：
- `axes[0]` → 第1个子图
- `.barh()` → 画横向柱状图
- `edu_plot.index` → Y轴（学历名称）
- `edu_plot['mean']` → X轴（平均薪资）

#### 第208行：透视表（重点！）
```python
pivot_edu_exp = df_clean.pivot_table(values='salary_monthly', index='education_name', columns='experience_name', aggfunc='mean')
```
**白话**：制作一个"学历×经验"的薪资透视表

**结果示例**：
```
experience_name  应届生  1年以内  1-3年  3-5年  5-10年  10年以上
education_name
学历不限          5K     6K     8K    10K    12K     15K
大专             6K     7K     9K    12K    15K     18K
本科             8K     9K    12K    16K    20K     25K
硕士            12K    14K    18K    24K    30K     40K
博士            20K    22K    28K    35K    45K     60K
```

这就是**热力图的数据来源**！

---

### 🎨 Step 2：Task 2 - 职位画像（第230-338行）

#### 第237行：画两个子图
```python
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
```
**白话**：1行2列的画面，左边画直方图，右边画箱线图。

#### 第239行：直方图
```python
axes[0].hist(df_clean['salary_monthly'], bins=50, color='steelblue', edgecolor='white', alpha=0.8)
```
**白话**：
- `.hist()` → 画直方图（薪资分布）
- `bins=50` → 分成50个区间
- `alpha=0.8` → 透明度80%

**直方图是什么？**
- 把薪资从0K到100K分成50段
- 每段里有多少人，柱子就有多高
- 用来观察薪资的**整体分布形状**

#### 第246-248行：箱线图
```python
salary_by_edu = [df_clean[df_clean['education_name'] == e]['salary_monthly'].dropna() for e in edu_plot.index]
bp = axes[1].boxplot(salary_by_edu, labels=edu_plot.index, patch_artist=True)
```
**白话**：
- 箱线图用来显示薪资的**分散程度**
- 每个学历画一个箱子，可以看到：最小值、25分位、中位数、75分位、最大值、异常值

**箱线图怎么看？**
```
  |   |
  |   |
======  ← 中位数
  |   |
  |   |
  +   +  ← 须（最小值和最大值）
  o      ← 异常值（特别高或特别低）
```

#### 第262行：分位数（重点！）
```python
salary_threshold = df_clean['salary_monthly'].quantile(0.8)
```
**白话**：
- `quantile(0.8)` → 取80%分位数
- 意思是：80%的人薪资低于这个值，20%的人高于这个值
- 这里用来定义"高薪职位"（Top 20%）

#### 第267-269行：占比统计
```python
high_edu = df_high_salary['education_name'].value_counts(normalize=True).head(5)
```
**白话**：
- `value_counts()` → 统计每个学历出现了多少次
- `normalize=True` → 转换成百分比（0.1表示10%）
- `.head(5)` → 取前5个

**结果示例**：
```
本科    0.45  (45%的高薪职位要求本科学历)
硕士    0.30  (30%)
博士    0.15  (15%)
...
```

---

### 🤖 Step 3：Task 3 - 薪酬建模（第340-458行）

#### 第346-347行：导入机器学习库
```python
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import LabelEncoder
```
**白话**：
- `LinearRegression` → 线性回归模型（用来预测薪资）
- `LabelEncoder` → 把文字转成数字（因为电脑只认识数字）

#### 第350-359行：编码分类变量（重点！）
```python
le_edu = LabelEncoder()
df_model['edu_enc'] = le_edu.fit_transform(df_model['education_name'])
```
**白话**：
- 线性回归模型不认识"本科"、"硕士"这种文字
- `LabelEncoder` 把它们转成数字：本科→0，硕士→1，博士→2
- 这样模型才能计算

#### 第362-368行：训练模型
```python
X = df_model[['edu_enc', 'exp_enc', 'city_enc', 'ct_enc']]
y = df_model['salary_monthly']

lr = LinearRegression()
lr.fit(X, y)
y_pred = lr.predict(X)
r2 = lr.score(X, y)
```
**白话**：
- `X` → 特征（学历、经验、城市、行业）
- `y` → 目标（薪资）
- `lr.fit(X, y)` → 训练模型（让模型学习特征和薪资的关系）
- `lr.predict(X)` → 用模型预测薪资
- `lr.score(X, y)` → 计算R²（模型准确度，1.0是完美）

**R²是什么意思？**
- R² = 0.7 → 模型能解释70%的薪资差异
- 剩下的30%可能是运气、公司规模等其他因素

#### 第373-374行：查看系数
```python
for name, coef in zip(['学历', '经验', '城市', '行业'], lr.coef_):
    print(f"  {name} 系数: {coef:.4f}K")
```
**白话**：
- 系数表示：这个因素每变化1个单位，薪资变化多少
- 比如经验的系数=5 → 经验每高1级，薪资增加5K

#### 第419-452行：K-Means聚类
```python
from sklearn.cluster import KMeans
kmeans = KMeans(n_clusters=4, random_state=42, n_init=10)
features_for_cluster['cluster'] = kmeans.fit_predict(X_cluster)
```
**白话**：
- K-Means是一种**无监督学习**算法
- 自动把数据分成4类（这里分成4种"薪酬模式"）
- 不需要人工标注，算法自己找规律

**聚类的结果**：
- 类别0：低薪大众型（薪资低，招聘量大）
- 类别1：高薪规模化（薪资高，招聘量也大）
- 类别2：精品高薪型（薪资很高，但招聘量少）
- 类别3：新兴职位（可能是新行业，模式不明显）

---

### 🗺️ Step 4：Task 4 - 地域招聘画像（第460-565行）

#### 第497-531行：城市聚类（重点！）
```python
from scipy.cluster.hierarchy import dendrogram, linkage, fcluster
```
**白话**：这是**层次聚类**，用来找出"哪些城市的行业分布相似"

**为什么要这样做？**
- 北京和上海的行业分布可能很像（都有很多互联网、金融行业）
- 通过聚类，可以把相似的城市分成一组
- 这样求职者可以参考："哦，原来杭州的互联网行业和上海很像"

#### 第508-515行：余弦相似度
```python
def cosine_sim(a, b):
    return np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b) + 1e-10)
```
**白话**：
- 余弦相似度用来衡量两个向量的相似程度
- 范围是0~1，1表示完全相同，0表示完全不相关
- 这里用来比较城市的行业分布向量

**举例**：
- 北京的行业分布：[互联网40%, 金融30%, 其他30%]
- 上海的行业分布：[互联网35%, 金融40%, 其他25%]
- 余弦相似度=0.95 → 说明两个城市的行业分布非常相似

---

### 🚀 Step 5：Task 5 - 新兴职位发现（第567-691行）

#### 第574-575行：分位数阈值
```python
high_demand_threshold = df_clean['salary_monthly'].quantile(0.7)
high_salary_threshold = df_clean['salary_monthly'].quantile(0.7)
```
**白话**：
- 70%分位数，用来定义"高需求"和"高薪资"
- 意思是：薪资高于70%的行业，才算"高薪行业"

#### 第611-616行：综合评分（重点！）
```python
industry_analysis['score'] = (
    (industry_analysis['avg_salary'] / industry_analysis['avg_salary'].max()) * 0.4 +
    (industry_analysis['mid_exp_ratio'] / 100) * 0.3 +
    (industry_analysis['job_count'] / industry_analysis['job_count'].max()) * 0.3
)
```
**白话**：
- 给每个行业打个分，综合考虑3个因素：
  - 平均薪资（权重40%）
  - 中级经验需求占比（权重30%）
  - 招聘量（权重30%）
- 分数越高，说明这个行业越"新兴"（高薪+需求大+需要中级人才）

**为什么要这样设计评分？**
- 高薪 → 行业有钱，发展迅速
- 中级经验需求高 → 行业正在扩张，需要大量"3-5年"的成熟人才
- 招聘量大 → 行业在快速增长

---

## 🎯 第三部分：常见问题解答

### Q1：为什么要用 `df.copy()`？
**A**：因为pandas有时会自动修改原始数据，用 `.copy()` 可以避免"改了这个，那个也变了"的bug。

### Q2：`axis=0` 和 `axis=1` 有什么区别？
**A**：
- `axis=0` → 按列操作（上下方向）
- `axis=1` → 按行操作（左右方向）

举例：
```python
df.dropna(axis=0)  # 删除有空值的行
df.dropna(axis=1)  # 删除有空值的列
```

### Q3：`apply()` 和 `map()` 有什么区别？
**A**：
- `map()` → 只能用于Series（一列），用来映射替换
- `apply()` → 可以用于Series和DataFrame，可以应用任意函数

举例：
```python
# map：把A变成本科，B变成硕士
df['学历'].map({'A': '本科', 'B': '硕士'})

# apply：对每一行执行自定义函数
df['薪资'].apply(lambda x: x * 2)  # 薪资全部翻倍
```

### Q4：为什么要做字段解码？
**A**：
- 官方数据用了虚拟代码（Gh、GJ、51b08e...），可读性差
- 解码后变成"本科"、"硕士"、"北京"，方便分析和展示
- **但要注意**：解码是基于"薪资推断"的，不是官方标准，需要在文档里说明！

### Q5：线性回归模型的R²低怎么办？
**A**：
- R²低说明"学历+经验+城市+行业"这4个因素解释不了薪资的所有差异
- 可以尝试加入更多特征：公司规模、职位名称、技能要求等
- 或者换用更复杂的模型：决策树、随机森林、XGBoost等

---

## 📋 第四部分：如何运行这个代码？

### 步骤1：安装依赖
```bash
pip install pandas numpy matplotlib seaborn scikit-learn openpyxl
```

### 步骤2：确认文件路径
- Excel文件要在 `chinavis2024/data/JobWanted.xlsx`
- 如果路径不对，修改第30行的 `BASE_DIR`

### 步骤3：运行代码
```bash
cd c:\Users\21009\WorkBuddy\20260503150431\chinavis2024\scripts
python main_analysis.py
```

### 步骤4：查看输出
- 图表保存在 `chinavis2024/figures/`
- 清洗后的数据保存在 `chinavis2024/data/job_cleaned.csv`

---

## 🎓 第五部分：学习建议

1. **先跑通代码** → 看到结果，有成就感
2. **修改参数试试** → 比如把 `quantile(0.8)` 改成 `quantile(0.9)`，看看高薪职位的定义变化后，结果有什么不同
3. **画自己的图** → 试着用 `df_clean` 画一个你感兴趣的图
4. **多看官方文档** → [pandas官网](https://pandas.pydata.org/docs/)有详细教程

---

**🎉 恭喜你读完了！**
如果还有不懂的地方，随时问我！
