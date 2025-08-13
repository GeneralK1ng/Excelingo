# XLSX翻译工具

基于DeepSeek API的XLSX文件翻译工具，支持异步批量翻译。

## 安装

1. 激活虚拟环境：
```bash
source venv/bin/activate
```

2. 安装依赖：
```bash
uv pip install -e .
```

3. 配置API密钥：
编辑 `.env` 文件，填入你的DeepSeek API密钥：
```
DEEPSEEK_API_KEY=your_api_key_here
```

## 使用

运行程序：
```bash
python main.py
```

## 项目结构

```
xlsxTrans/
├── src/
│   ├── translator/          # 翻译核心模块
│   │   ├── llm_client.py   # DeepSeek API客户端
│   │   └── xlsx_processor.py # XLSX处理器
│   └── gui/                # GUI界面
│       └── main_window.py  # 主窗口
├── main.py                 # 程序入口
├── .env                    # 环境变量配置
└── pyproject.toml         # 项目配置
```

## 特性

- 异步批量翻译，提高效率
- 简洁的GUI界面
- 支持多种目标语言
- 自动保存翻译结果