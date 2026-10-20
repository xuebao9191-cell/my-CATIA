# my-CATIA

CATIA V5 二次开发脚本（VBA / EKL）与学习笔记仓库。

## 目录结构

```
my-CATIA/
├── scripts/
│   ├── VBA/          # VBA 宏脚本（自动化、批量处理）
│   └── EKL/          # EKL 知识工程脚本（参数化、规则）
└── notes/            # 学习笔记（Markdown）
```

## 脚本说明

- **scripts/VBA**：Excel / CATIA VBA 宏，用于批量导入导出、批量处理、测量统计等，可直接在 CATIA Tools → Macros 中运行。
- **scripts/EKL**：知识工程模块脚本，用于参数化建模、规则与检查约束。

## 笔记说明

笔记按 CATIA 常用工作台与功能模块组织，适合从入门到进阶，重点覆盖飞行器构造建模常用操作。

## 使用

1. VBA 脚本：CATIA 中 `Tools → Macros → Run`，选择对应 `.bas` 或宏模块运行。
2. EKL 脚本：在 `Knowledge` 或 `Knowledge Advisor` 工作台中新建知识文件（.EKL）加载。
3. 笔记：用任意 Markdown 阅读器打开。

> 注意：脚本基于 CATIA V5 编写，运行前请先打开对应类型的文档（Part / Product 等）。
