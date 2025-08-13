# Excelingo

> **💡 特别说明**：本项目代码100%由AI编写，开发者只负责提需求和指导方向。本项目为开发者自用，并不能保证实际使用效果。

基于OpenAI兼容接口的现代化Excel翻译工具，支持拖拽操作和异步批量翻译。

## 快速开始

### 1. 获取API密钥
- 支持任何OpenAI兼容的API服务
- 推荐使用 [DeepSeek](https://platform.deepseek.com/) 或其他兼容服务
- 注册账号并获取API密钥

### 2. 环境配置

**创建虚拟环境：**
```bash
# 推荐使用uv（更快）
uv venv --python 3.12 venv
source venv/bin/activate  # Windows用户: venv\Scripts\activate

# 或使用标准Python
python -m venv venv
source venv/bin/activate  # Windows用户: venv\Scripts\activate
```

**安装依赖：**
```bash
# 使用uv（推荐，速度更快）
uv pip install -e .

# 或使用pip
pip install -e .
```

**配置API密钥：**
在项目根目录创建 `.env` 文件：
```env
DEEPSEEK_API_KEY=你的API密钥
DEEPSEEK_BASE_URL=https://api.deepseek.com/v1
# 或使用其他OpenAI兼容服务
# DEEPSEEK_BASE_URL=https://api.openai.com/v1
```

### 3. 运行程序
```bash
python main.py
```

## 使用方法

1. **选择文件**：将Excel文件（.xlsx）拖拽到界面中，或点击选择文件
2. **选择语言**：从下拉菜单选择目标翻译语言
3. **设置并发数**：调整同时进行的API请求数量（数值越大翻译越快）
4. **开始翻译**：点击"Start Translation"按钮，实时查看翻译进度
5. **获取结果**：翻译完成后，文件会自动保存为 `原文件名_translated_语言.xlsx`

## 功能特色

- **现代化界面**：简洁专业的拖拽式用户界面
- **高速翻译**：异步批量处理，支持自定义并发数
- **多语言支持**：支持英语、中文、日语、韩语、法语、德语、西班牙语、俄语
- **实时反馈**：实时进度条和详细日志显示
- **智能处理**：只翻译文本单元格，保留原有格式
- **自动保存**：自动保存翻译结果，文件命名清晰

## 项目结构

```
Excelingo/
├── src/
│   ├── translator/
│   │   ├── llm_client.py      # OpenAI兼容API客户端
│   │   └── xlsx_processor.py  # Excel文件处理器
│   └── gui/
│       └── main_window.py     # 主应用窗口
├── main.py                    # 程序入口
├── pyproject.toml            # 项目配置
├── .env                      # API配置文件（需要创建）
└── README.md                 # 说明文档
```

## 系统要求

- Python 3.12+
- OpenAI兼容API密钥（如DeepSeek、OpenAI等）
- Excel文件格式：.xlsx

**API连接问题：**
- 检查 `.env` 文件中的API密钥是否正确
- 确保网络连接正常
- 验证API服务账户是否有足够余额
- 确认API服务地址是否正确

**翻译速度太慢：**
- 增加并发数设置（建议尝试80-100）
- 检查网络速度
- 大文件翻译需要更长时间