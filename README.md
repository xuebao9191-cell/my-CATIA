# my-CATIA

CATIA V5 二次开发脚本（VBA / EKL）、学习笔记与模型归档示例。

## 从这里开始

- [导入与启动指南](IMPORT_GUIDE.md)：VBA、EKL、模型文件分别如何加载。
- [VBA 脚本目录](scripts/VBA/)：CATIA 自动化宏。
- [EKL 示例说明](scripts/EKL/README.md)：知识工程示例状态与验证方式。
- [学习笔记](notes/)：CATIA 常用工作台和建模操作。
- [模型存放说明](models/xuebao/README.md)：模型目录约定。

## 重要说明

仓库中的 `.bas` 是 VBA 模块源文件，通常需要导入 CATIA 可访问的 VBA 项目后运行；`.EKL` 文件按知识工程示例片段管理，需在对应编辑器中验证；CATPart/CATProduct/STEP/IGES 模型需通过 CATIA 的打开命令单独加载。**打开仓库本身不会自动将所有文件导入 CATIA，也不建议自动执行所有宏。**

批量脚本运行前请检查路径、输出覆盖策略，并先在副本或测试模型上验证。代码静态修订不等于已在所有 CATIA V5 版本、服务包与许可证配置下实机验证。
