# Excel批量操作工具箱 🔧

一个专为非技术用户设计的简单易用Excel文件批量处理Web应用。无需安装任何软件，打开浏览器即可使用。

## ✨ 功能特性

- **📁 合并多个Excel文件** - 将结构相同的文件纵向合并
- **📊 合并单个文件的多个Sheet** - 将一个文件的所有Sheet合并
- **✂️ 按列拆分Sheet** - 根据列值拆分成多个文件
- **📄 按行数拆分Sheet** - 按指定行数拆分大文件
- **🔍 批量查找与替换** - 在多个文件中进行文本替换
- **🗑️ 批量删除指定列** - 删除不需要的数据列
- **🎯 批量数据筛选** - 根据条件筛选并合并数据
- **🔄 格式转换** - XLSX与CSV格式互转

## 🚀 快速开始

### 前置要求

- Python 3.8 或更高版本
- uv (推荐的Python包管理工具)

### 步骤 1: 安装 uv

如果您还没有安装 uv，请选择以下任一方式安装：

**方式一：使用 pip 安装**
```bash
pip install uv
```

**方式二：使用官方安装脚本**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

**方式三：使用包管理器**
```bash
# macOS
brew install uv

# Windows
winget install AstralSoftware.uv

# Linux (Debian/Ubuntu)
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### 步骤 2: 克隆项目并进入目录

```bash
git clone https://github.com/StanleyChanH/excel-toolkit.git
cd excel-toolkit
```

### 步骤 3: 创建虚拟环境

```bash
uv venv
```

### 步骤 4: 激活虚拟环境

**Windows:**
```bash
.venv\Scripts\activate
```

**macOS/Linux:**
```bash
source .venv/bin/activate
```

### 步骤 5: 安装依赖

```bash
uv sync
```

### 步骤 6: 启动应用

有两种方式启动应用：

**方式一：使用启动脚本（推荐）**
```bash
uv run python run.py
```
这种方式会自动询问是否打开浏览器，并提供更好的启动体验。

**方式二：直接使用Flask**
```bash
uv run flask --app app run
```

应用将在 `http://localhost:5000` 启动。

### 步骤 7: 创建示例文件（可选）

如果您想测试应用功能，可以运行以下命令创建示例数据文件：

```bash
uv run python create_samples.py
```

这将在 `examples/` 文件夹中创建以下示例文件：
- `sample_data1.xlsx` - 单Sheet示例文件
- `sample_data2.xlsx` - 单Sheet示例文件
- `multi_sheet_data.xlsx` - 多Sheet示例文件
- `sample_data1.csv` - CSV格式示例文件

## 📖 使用指南

### 基本操作流程

1. **选择功能** - 在顶部导航栏选择需要的功能
2. **上传文件** - 点击文件选择区域或拖拽文件
3. **设置参数** - 根据功能需要填写相关参数
4. **开始处理** - 点击处理按钮等待完成
5. **下载结果** - 处理完成后自动显示下载链接

### 各功能详细说明

#### 1. 合并多个Excel文件
- **用途**：将多个结构相同的Excel文件合并成一个文件
- **场景**：月度报告汇总、部门数据整合
- **选项**：
  - 保留标题行（默认）
  - 添加来源文件列

#### 2. 合并单个文件的多个Sheet
- **用途**：将一个Excel文件中的所有Sheet合并成一个Sheet
- **场景**：按月份分Sheet的数据汇总
- **选项**：
  - 添加来源Sheet列

#### 3. 按列拆分Sheet
- **用途**：根据指定列的值将数据拆分成多个文件
- **场景**：按部门、地区、类别等分发数据
- **参数**：列名（如"部门"、"地区"）
- **输出**：ZIP压缩包，包含所有拆分文件

#### 4. 按行数拆分Sheet
- **用途**：将大文件按指定行数拆分成小文件
- **场景**：控制文件大小、便于传输
- **参数**：每文件的行数
- **输出**：ZIP压缩包，包含所有拆分文件

#### 5. 批量查找与替换
- **用途**：在多个文件中进行全局文本替换
- **场景**：批量纠错、统一术语修改
- **参数**：
  - 查找内容（必填）
  - 替换内容（可为空，表示删除）
- **输出**：ZIP压缩包，包含所有处理后的文件

#### 6. 批量删除指定列
- **用途**：删除多个文件中的指定列
- **场景**：清理敏感信息、简化数据结构
- **参数**：要删除的列名（多个用逗号分隔）
- **输出**：ZIP压缩包，包含所有处理后的文件

#### 7. 批量数据筛选
- **用途**：从多个文件中筛选符合条件的数据
- **场景**：提取特定条件的数据进行汇总分析
- **参数**：
  - 列名
  - 条件（等于、包含、大于等）
  - 筛选值
- **输出**：单个Excel文件，包含所有符合条件的行

#### 8. 格式转换
- **用途**：Excel文件与CSV文件格式互转
- **场景**：数据格式标准化、系统导入导出
- **类型**：
  - XLSX转CSV：每个Sheet生成一个CSV文件
  - CSV转XLSX：每个CSV生成一个Excel文件
- **输出**：ZIP压缩包，包含所有转换后的文件

## 🔧 技术架构

- **后端**：Flask + Pandas + Openpyxl
- **前端**：原生 HTML/CSS/JavaScript
- **包管理**：uv
- **文件处理**：内存处理，无服务器残留

## 📋 系统要求

- **操作系统**：Windows、macOS、Linux
- **Python**：3.8+
- **浏览器**：Chrome、Firefox、Safari、Edge（现代浏览器）
- **内存**：建议至少2GB可用内存
- **存储**：临时处理文件可能需要额外空间

## 🛡️ 安全说明

- 所有文件处理都在内存中进行，处理完成后立即清理
- 不会永久存储用户上传的文件
- 支持100MB以内的文件上传
- 建议不要处理包含敏感信息的文件

## ❓ 常见问题

**Q: 支持哪些文件格式？**
A: 支持 .xlsx、.xls 和 .csv 格式。

**Q: 文件大小有限制吗？**
A: 单个文件最大支持100MB。

**Q: 可以同时处理多少个文件？**
A: 建议一次处理不超过20个文件，以确保性能稳定。

**Q: 处理大量数据时会卡住吗？**
A: 应用采用了内存处理机制，但对于特别大的文件（超过10万行），处理时间会相应延长。

**Q: 如果处理失败怎么办？**
A: 请检查文件格式是否正确、参数是否完整，然后重试。如问题持续，请查看控制台错误信息。

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request！

1. Fork 本项目
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

## 📝 更新日志

### v1.0.0
- 初始版本发布
- 实现8个核心Excel操作功能
- 完整的Web界面和用户体验
- 支持批量文件处理

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🙏 致谢

- [Flask](https://flask.palletsprojects.com/) - Web框架
- [Pandas](https://pandas.pydata.org/) - 数据处理
- [Openpyxl](https://openpyxl.readthedocs.io/) - Excel文件处理
- [uv](https://github.com/astral-sh/uv) - Python包管理

## 📞 联系我们

如有问题或建议，请通过以下方式联系：

- 提交 [Issue](https://github.com/StanleyChanH/excel-toolkit/issues)
- 发送邮件至：stanleychan@example.com

---

⭐ 如果这个项目对您有帮助，请给我们一个星标！