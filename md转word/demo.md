# Markdown 转 Word 演示文档

这是一个演示文档，展示 Markdown 到 Word 转换的各种效果。

---

## 一、文本格式演示

### 1.1 基础文本格式

这是 **粗体文本** 的演示。

这是 *斜体文本* 的演示。

这是 ***粗体且斜体*** 的演示。

这是 `行内代码` 的演示。

这是一条 ~~删除线~~ 的演示。

### 1.2 列表演示

#### 无序列表
- 第一项
- 第二项
  - 嵌套项 2.1
  - 嵌套项 2.2
- 第三项

#### 有序列表
1. 第一步
2. 第二步
   1. 子步骤 2.1
   2. 子步骤 2.2
3. 第三步

---

## 二、代码块演示

### 2.1 Python 代码

```python
def hello_world():
    """这是一个示例函数"""
    print("Hello, World!")

    # 创建一个列表
    numbers = [1, 2, 3, 4, 5]
    for num in numbers:
        print(f"Number: {num}")

if __name__ == "__main__":
    hello_world()
```

### 2.2 JavaScript 代码

```javascript
// 异步函数示例
async function fetchData(url) {
    try {
        const response = await fetch(url);
        const data = await response.json();
        return data;
    } catch (error) {
        console.error('Error:', error);
    }
}

// 使用
fetchData('https://api.example.com/data')
    .then(data => console.log(data));
```

### 2.3 Shell 命令

```bash
#!/bin/bash
# 这是一个 bash 脚本

echo "Installing dependencies..."
pip install -r requirements.txt

echo "Running tests..."
pytest tests/

echo "Done!"
```

---

## 三、表格演示

| 功能 | 说明 | 状态 |
|------|------|------|
| Markdown 预处理 | 去掉水平分割线 | ✓ 完成 |
| 参考模板创建 | 自定义样式 | ✓ 完成 |
| 字体自定义 | 宋体和 Times New Roman | ✓ 完成 |
| 目录生成 | 自动生成 TOC | ✓ 完成 |
| 图片支持 | 相对路径解析 | ✓ 完成 |

---

## 四、引用和提示

### 4.1 块引用

> 这是一个块引用的示例。
>
> 它可以包含多行内容。
>
> 还可以包含其他格式，比如 **粗体** 和 *斜体*。

### 4.2 深层嵌套引用

> 第一层引用
>> 第二层引用
>>> 第三层引用

---

## 五、数学公式演示（如果支持）

### 5.1 行内公式

根据爱因斯坦的相对论，$E = mc^2$ 是一个非常著名的公式。

### 5.2 块级公式

$$
\int_{-\infty}^{\infty} e^{-x^2} dx = \sqrt{\pi}
$$

---

## 六、链接和参考

### 6.1 超链接

[访问 Pandoc 官网](https://pandoc.org/)

### 6.2 图片（如果存在）

如果 Markdown 文件所在目录有图片，使用以下格式：
```
![图片描述](image.png)
```

---

## 七、特殊内容

### 7.1 水平线

下面是一条水平线：

---

### 7.2 混合格式

在这个部分，我们展示 **粗体链接** [example](https://example.com) 和 `代码链接`。

### 7.3 定义列表演示

转换工具的主要特性：
- **快速**: 利用 Pandoc 的高效转换
- **可靠**: 支持多种 Markdown 语法
- **灵活**: 支持自定义模板和样式

---

## 八、总结

这个演示文档展示了 Markdown 到 Word 转换工具支持的各种格式：

1. ✓ 文本格式（粗体、斜体、删除线）
2. ✓ 列表（有序和无序）
3. ✓ 代码块（带语法高亮）
4. ✓ 表格
5. ✓ 引用
6. ✓ 链接
7. ✓ 数学公式
8. ✓ 多级标题

希望转换效果令人满意！

---

**文档生成时间**: 2026-01-09

**工具版本**: Markdown to Word Converter v1.0
