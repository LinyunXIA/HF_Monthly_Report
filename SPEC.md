# 月度报告处理系统 (HF Monthly Report)

## 1. 项目概述

这是一个基于Django的Web应用，用于批量处理月度Excel报告文件。用户可上传多个xlsx文件，系统自动提取事件分类和事件描述数据，生成汇总报告。

## 2. 技术栈

- **后端框架**: Django 6.0.4
- **Python版本**: 3.x
- **数据库**: SQLite3
- **前端**: 纯HTML + JavaScript (无框架)
- **Excel处理**: openpyxl
- **部署环境**: Windows (Asia/Shanghai时区)

## 3. 主要用户

- 企业行政/办公人员
- 数据处理专员
- 需要批量汇总月度事件报告的用户

## 4. 功能特性

### 4.1 文件上传
- 支持多文件同时上传
- 仅接受.xlsx格式
- 支持文件夹选择上传

### 4.2 数据处理
- 自动提取xlsx文件中的"事件分类"和"事件描述"列
- 统计各事件分类的数量
- 单独提取"其他"类别的详细描述

### 4.3 报告生成
- 生成包含3个工作表的Excel文件:
  - `Details`: 所有事件明细
  - `Summary`: 事件分类统计
  - `Others`: "其他"类别详情

### 4.4 历史记录
- 记录每次上传的时间、文件数量、处理状态
- 支持下载已处理完成的报告

### 4.5 状态轮询
- 实时显示处理状态 (上传中 → 处理中 → 处理完毕)
- 自动刷新状态直到完成

## 5. API端点

| 端点 | 方法 | 功能 |
|------|------|------|
| `/` | GET | 首页 |
| `/upload/` | POST | 上传文件 |
| `/history/` | GET | 获取历史记录 |
| `/status/<id>/` | GET | 获取处理状态 |
| `/download/<id>/` | GET | 下载报告 |

## 6. 数据存储

- 上传的xlsx文件: `data/upload_{id}/`
- 处理结果: `data/result_{id}.xlsx`
- 历史记录: `data/records.json`

## 7. 目录结构

```
HF_Monthly_Report/
├── hf_report/          # Django项目配置
├── hf_app/             # 主应用
│   ├── views.py        # 视图函数
│   ├── urls.py        # 路由配置
│   └── models.py      # 数据模型(未使用)
├── templates/
│   └── index.html     # 前端页面
├── data/               # 数据存储目录
├── manage.py
└── db.sqlite3
```