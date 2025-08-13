# ExcelLingo - 智能表格翻译工具

基于DeepSeek API的智能表格翻译工具，支持异步批量翻译和拖拽操作。

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

- 现代化的拖拽界面，操作简单
- 异步批量翻译，速度快
- 可调节并发数，灵活控制翻译速度
- 支持多种目标语言
- 实时进度显示和日志记录
- 自动保存翻译结果