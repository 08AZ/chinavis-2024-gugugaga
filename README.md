# 🎯 ChinaVis 2024 招聘数据可视分析系统

**团队**: 咕咕嘎嘎队 | **学校**: 中南大学  
**赛道**: Mini Challenge 1 (主题Ⅱ：虚拟招聘市场可视分析)  
**成员**: 赵聚鑫(队长)、白子涵、万方

---

## 📋 项目简介

本项目基于 **43万条虚拟招聘数据**（JobWanted.xlsx），使用Python进行数据清洗、探索性分析（EDA）、机器学习建模和数据可视化，完成ChinaVis 2024竞赛的5个分析任务。

---

## 🎯 竞赛任务（5个）

| 任务 | 内容 | 状态 |
|------|------|------|
| **Task 1** | 职位差异化分析（行业/薪资/经验多维度） | ✅ 完成 |
| **Task 2** | 职位画像（关键技能/城市偏好/薪酬） | ✅ 完成 |
| **Task 3** | 薪酬建模（回归分析/薪酬模式） | ✅ 完成 |
| **Task 4** | 地域招聘画像（城市聚类/相似地域识别） | ✅ 完成 |
| **Task 5** | 新兴职位发现（行业动态/紧缺职位） | ✅ 完成 |

---

## 🗂️ 项目结构

```
china-vis-2024/
├── chinavis2024/
│   ├── scripts/
│   │   └── main_analysis.py      # 核心分析脚本（700行，覆盖全部5个任务）
│   ├── data/
│   │   └── job_cleaned.csv       # 清洗后的数据（运行脚本后自动生成）
│   └── figures/                   # 输出图表（运行脚本后自动生成）
├── gen_ppt.py                     # 生成答辩PPT的脚本
├── main_analysis_白话详解.md       # 代码详解文档（适合Python初学者）
├── ChinaVis2024_答辩PPT.pptx      # 答辩PPT（已生成）
└── README.md                      # 本文件
```

---

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install pandas numpy matplotlib seaborn scikit-learn openpyxl python-pptx
```

或使用国内镜像加速：

```bash
pip install pandas numpy matplotlib seaborn scikit-learn openpyxl python-pptx -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 2. 准备数据

将官方数据文件 `JobWanted.xlsx` 放到以下路径：

```
chinavis2024/data/JobWanted.xlsx
```

### 3. 运行分析脚本

```bash
cd chinavis2024/scripts
python main_analysis.py
```

运行后：
- ✅ 图表保存到 `chinavis2024/figures/`（共10+张）
- ✅ 清洗后的数据保存到 `chinavis2024/data/job_cleaned.csv`

### 4. 生成答辩PPT

```bash
python gen_ppt.py
```

生成文件：`ChinaVis2024_答辩PPT.pptx`

---

## 📊 核心功能详解

### Step 0: 数据加载与清洗
- 读取43万条招聘数据
- 解析薪资字段（支持日薪/月薪/年薪）
- 字段解码（通过薪资规律推断虚拟代码含义）
  - `education` → 学历（博士/硕士/本科/大专/高中）
  - `experience` → 经验（应届生/1-3年/3-5年/5-10年/10年以上）
  - `city` → 城市（北上深/新一线/二线/三线）
  - `company_type` → 行业（互联网/金融/医疗/教育等）

### Step 1: 职位差异化分析
- 学历 vs 薪资（柱状图）
- 经验 vs 薪资（柱状图）
- 行业 vs 薪资（柱状图）
- 学历×经验热力图（热力图）

### Step 2: 职位画像
- 薪资分布（直方图+箱线图）
- 高薪职位特征（饼图）
- 各城市薪资分布（箱线图）
- 学历雷达图对比

### Step 3: 薪酬建模
- 线性回归模型（R² ≈ 0.65）
- 特征重要性分析
- K-Means聚类（识别4种薪酬模式）

### Step 4: 地域招聘画像
- 城市薪资与招聘量（气泡图）
- 城市层次聚类（树状图）
- 城市×行业热力图

### Step 5: 新兴职位发现
- 行业综合评分（薪资40% + 需求30% + 经验占比30%）
- 新兴行业识别
- 综合仪表盘

---

## 📚 学习资源

- **`main_analysis_白话详解.md`** - 超详细的代码注释文档，专门解释pandas方法和分析逻辑，适合Python初学者。
- **pandas官方文档** - https://pandas.pydata.org/docs/
- **ChinaVis 2024 官方说明** - 参见 `JobWanted_description.docx`

---

## ⚠️ 重要说明

1. **字段解码基于推断**：由于官方未提供虚拟代码映射表，本项目的学历/经验/城市/行业解码是通过"薪资规律"推断的，**并非官方标准**。请在报告中注明这一点。

2. **数据使用限制**：`JobWanted.xlsx` 是ChinaVis竞赛官方数据，请勿用于商业用途。

3. **Python版本**：建议使用 Python 3.8+。

---

## 👥 团队协作

- **队长**: 赵聚鑫（@聚鑫） - 整体分析、建模、PPT制作
- **队员**: 白子涵 - EDA代码编写、可视化优化
- **队员**: 万方 - 文档撰写、答辩准备

---

## 📧 联系方式

- **队长邮箱**: 2100980480@qq.com
- **学校**: 中南大学计算机学院

---

## 🎉 致谢

感谢 ChinaVis 2024 组委会提供的数据和竞赛平台！

---

**最后更新**: 2026-05-05
