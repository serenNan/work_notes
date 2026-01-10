# Markdown 转 Word 演示文档

这是一个演示文档，展示各种 Markdown 格式的转换效果。

## 文本格式

这段文字包含 **粗体文本**、*斜体文本* 和 ***粗斜体*** 文本。你还可以使用 `行内代码` 来表示技术术语。

## 代码块

### Python 示例

```python
def hello_world():
    """一个简单的问候函数"""
    message = "你好，世界！"
    print(message)
    return message

if __name__ == "__main__":
    hello_world()
```

### JavaScript 示例

```javascript
const greet = (name) => {
    console.log(`你好，${name}！`);
};

greet("世界");
```

## 列表

### 无序列表

- 第一项
- 第二项
  - 嵌套项 A
  - 嵌套项 B
- 第三项

### 有序列表

1. 步骤一
2. 步骤二
3. 步骤三

## 表格

| 功能 | 状态 | 备注 |
|------|------|------|
| 目录生成 | 已支持 | 可配置深度 |
| 代码高亮 | 已支持 | 多种主题 |
| 表格边框 | 已支持 | 自动添加 |

## 引用块

> 这是一段引用文本。引用可以跨越多行，适合用于突出显示重要信息。

## 数学公式

行内公式: $E = mc^2$

块级公式:

$$
\sum_{i=1}^{n} x_i = x_1 + x_2 + \cdots + x_n
$$

## 链接

访问 [GitHub](https://github.com) 获取更多信息。

---

*演示文档结束*
