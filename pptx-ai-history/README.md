# pptx-ai-history（人工智能发展史 PPT）

使用 `python-pptx` 库生成的「人工智能发展史」演示文稿。

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `generate_ai_history_ppt.py` | Python 脚本：每页含标题与要点，共 5 页 |
| `ai_history.pptx` | 生成的 PowerPoint 成品（16:9） |

## 幻灯片大纲

1. 人工智能发展史（导言）
2. 思想的萌芽与奠基（1943–1955）
3. 黄金时代与两次寒冬（1956–1993）
4. 机器学习的复兴与深度学习崛起（1990s–2010s）
5. 当代 AI 浪潮与未来展望

## 运行方法

```bash
# 安装依赖
pip install python-pptx

# 重新生成
python generate_ai_history_ppt.py
```

脚本运行后会在当前目录重新生成 `ai_history.pptx`。
