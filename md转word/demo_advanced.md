# 高级 Markdown 转 Word 演示文档

这是一个高级演示文档，展示 Markdown 到 Word 转换的复杂格式和高级特性。

---

## 一、多层级标题和复杂结构

### 1.1 二级标题示例

#### 1.1.1 三级标题示例

##### 1.1.1.1 四级标题示例

###### 1.1.1.1.1 五级标题示例

这是五级标题下的段落文本。注意不同的标题级别在 Word 中会自动生成目录。

###### 1.1.1.1.2 另一个五级标题

段落内容。

##### 1.1.1.2 回到四级标题

#### 1.1.2 另一个三级标题

### 1.2 继续二级标题

---

## 二、复杂列表结构

### 2.1 深层嵌套的无序列表

- 第一个顶级项
  - 子项 1.1
    - 子子项 1.1.1
      - 子子子项 1.1.1.1
        - 五层嵌套
        - 又一个五层项
      - 另一个子子子项 1.1.1.2
    - 子子项 1.1.2
      - 很深的嵌套演示
  - 子项 1.2
    - 子子项 1.2.1
- 第二个顶级项
  - 子项 2.1
  - 子项 2.2
  - 子项 2.3
    - 子子项 2.3.1
    - 子子项 2.3.2
      - 子子子项 2.3.2.1
      - 子子子项 2.3.2.2
- 第三个顶级项

### 2.2 深层嵌套的有序列表

1. 第一个主要步骤
   1. 子步骤 1.1
      1. 详细步骤 1.1.1
         1. 更多细节 1.1.1.1
            1. 极其详细的步骤 1.1.1.1.1
               1. 最细的步骤
               2. 另一个最细的步骤
            2. 另一个详细步骤 1.1.1.1.2
         2. 细节 1.1.1.2
      2. 详细步骤 1.1.2
   2. 子步骤 1.2
2. 第二个主要步骤
   1. 子步骤 2.1
   2. 子步骤 2.2
3. 第三个主要步骤

### 2.3 混合列表（有序和无序）

1. 有序项 1
   - 无序子项 1.1
   - 无序子项 1.2
     1. 有序子子项 1.2.1
     2. 有序子子项 1.2.2
   - 无序子项 1.3
2. 有序项 2
   - 无序子项 2.1
     1. 有序子子项 2.1.1
        - 无序子子子项
        - 另一个无序子子子项
     2. 有序子子项 2.1.2
   - 无序子项 2.2
3. 有序项 3

### 2.4 任务列表（如果支持）

- [x] 已完成的任务 1
- [x] 已完成的任务 2
- [ ] 未完成的任务 1
- [ ] 未完成的任务 2
  - [x] 子任务 2.1 已完成
  - [ ] 子任务 2.2 未完成
- [ ] 未完成的任务 3

---

## 三、代码块和代码格式

### 3.1 多种编程语言代码块

#### 3.1.1 Python 高级示例

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
这是一个复杂的 Python 示例，展示各种高级特性
"""

from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from abc import ABC, abstractmethod
import asyncio
from functools import lru_cache

@dataclass
class Person:
    """人物数据类"""
    name: str
    age: int
    skills: List[str]

    def __repr__(self) -> str:
        return f"Person(name={self.name}, age={self.age}, skills={self.skills})"

class Worker(ABC):
    """抽象工作者类"""

    def __init__(self, name: str):
        self.name = name

    @abstractmethod
    def work(self) -> str:
        """执行工作"""
        pass

    @lru_cache(maxsize=128)
    def calculate_efficiency(self, hours: int) -> float:
        """计算效率"""
        return 100.0 / (hours + 1)

class Developer(Worker):
    """开发者类"""

    def __init__(self, name: str, languages: List[str]):
        super().__init__(name)
        self.languages = languages

    def work(self) -> str:
        return f"{self.name} is coding in {', '.join(self.languages)}"

    async def deploy(self) -> str:
        """异步部署"""
        await asyncio.sleep(1)
        return f"{self.name} deployed successfully"

# 异步函数示例
async def main():
    dev = Developer("Alice", ["Python", "JavaScript", "Go"])
    print(dev.work())
    result = await dev.deploy()
    print(result)

if __name__ == "__main__":
    asyncio.run(main())
```

#### 3.1.2 JavaScript/TypeScript 示例

```typescript
// TypeScript 高级示例
interface IUser {
    id: number;
    name: string;
    email: string;
    roles?: string[];
}

type UserRole = 'admin' | 'user' | 'guest';

abstract class BaseService<T> {
    protected data: Map<number, T> = new Map();

    abstract validate(item: T): boolean;

    async save(item: T): Promise<void> {
        if (this.validate(item)) {
            // 保存逻辑
            console.log('Item saved');
        }
    }
}

class UserService extends BaseService<IUser> {
    validate(user: IUser): boolean {
        return user.id > 0 && user.email.includes('@');
    }

    async getUser(id: number): Promise<IUser | undefined> {
        return new Promise((resolve) => {
            setTimeout(() => {
                resolve(this.data.get(id));
            }, 100);
        });
    }
}

// 泛型函数
function processArray<T>(arr: T[], callback: (item: T) => T): T[] {
    return arr.map(callback);
}

// 使用示例
const userService = new UserService();
const users: IUser[] = [
    { id: 1, name: 'Alice', email: 'alice@example.com', roles: ['admin'] },
    { id: 2, name: 'Bob', email: 'bob@example.com' }
];

users.forEach(user => userService.save(user));
```

#### 3.1.3 Java 示例

```java
package com.example.advanced;

import java.util.*;
import java.util.stream.*;
import java.util.function.*;
import javax.annotation.*;

/**
 * 高级 Java 示例类
 */
public class AdvancedJavaExample {

    private static final Logger LOGGER = LoggerFactory.getLogger(AdvancedJavaExample.class);

    /**
     * 通用数据处理类
     */
    public static class DataProcessor<T> {
        private final List<T> data;
        private final Predicate<T> filter;

        public DataProcessor(List<T> data, Predicate<T> filter) {
            this.data = data;
            this.filter = filter;
        }

        public <R> List<R> process(Function<T, R> mapper) {
            return data.stream()
                .filter(filter)
                .map(mapper)
                .collect(Collectors.toList());
        }
    }

    /**
     * Lambda 和 Stream API 示例
     */
    public static void demonstrateStreams() {
        List<Integer> numbers = Arrays.asList(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);

        numbers.stream()
            .filter(n -> n % 2 == 0)
            .map(n -> n * n)
            .sorted()
            .forEach(System.out::println);
    }

    /**
     * 函数式接口示例
     */
    @FunctionalInterface
    public interface Operation<T> {
        T execute(T a, T b);
    }

    public static <T extends Number> void main(String[] args) {
        Operation<Integer> add = (a, b) -> a + b;
        Operation<Integer> multiply = (a, b) -> a * b;

        System.out.println(add.execute(5, 3));
        System.out.println(multiply.execute(5, 3));
    }
}
```

#### 3.1.4 C++ 示例

```cpp
#include <iostream>
#include <vector>
#include <memory>
#include <algorithm>
#include <functional>
#include <thread>
#include <mutex>

// 模板类示例
template<typename T>
class Container {
private:
    std::vector<T> elements;
    mutable std::mutex mutex_;

public:
    void push(const T& value) {
        std::lock_guard<std::mutex> lock(mutex_);
        elements.push_back(value);
    }

    template<typename Func>
    void forEach(Func callback) {
        std::lock_guard<std::mutex> lock(mutex_);
        for (const auto& elem : elements) {
            callback(elem);
        }
    }

    size_t size() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return elements.size();
    }
};

// 智能指针示例
class Resource {
public:
    Resource() { std::cout << "Resource allocated\n"; }
    ~Resource() { std::cout << "Resource freed\n"; }
};

int main() {
    // 使用 unique_ptr
    std::unique_ptr<Resource> resource(new Resource());

    // 使用容器
    Container<int> container;
    container.push(1);
    container.push(2);
    container.push(3);

    container.forEach([](int value) {
        std::cout << "Value: " << value << std::endl;
    });

    // Lambda 和 std::function
    std::vector<int> numbers{1, 2, 3, 4, 5};
    std::transform(numbers.begin(), numbers.end(),
                   numbers.begin(),
                   [](int n) { return n * n; });

    return 0;
}
```

#### 3.1.5 SQL 示例

```sql
-- 复杂 SQL 查询示例
WITH sales_summary AS (
    SELECT
        customer_id,
        DATE_TRUNC('month', order_date) as month,
        SUM(amount) as total_sales,
        COUNT(*) as order_count,
        AVG(amount) as avg_order
    FROM orders
    WHERE order_date >= NOW() - INTERVAL '12 months'
    GROUP BY customer_id, DATE_TRUNC('month', order_date)
),
customer_ranks AS (
    SELECT
        customer_id,
        month,
        total_sales,
        order_count,
        avg_order,
        ROW_NUMBER() OVER (
            PARTITION BY month
            ORDER BY total_sales DESC
        ) as rank_in_month,
        RANK() OVER (
            ORDER BY total_sales DESC
        ) as global_rank
    FROM sales_summary
)
SELECT
    c.customer_id,
    c.customer_name,
    cr.month,
    cr.total_sales,
    cr.order_count,
    cr.avg_order,
    cr.rank_in_month,
    CASE
        WHEN cr.global_rank <= 10 THEN 'Top 10'
        WHEN cr.global_rank <= 50 THEN 'Top 50'
        ELSE 'Regular'
    END as customer_tier
FROM customer_ranks cr
JOIN customers c ON cr.customer_id = c.customer_id
ORDER BY cr.month DESC, cr.rank_in_month ASC;
```

### 3.2 行内代码示例

在段落中使用 `inline code` 来显示代码片段。可以混合使用 **`粗体代码`** 和 *`斜体代码`*。

还可以写这样的代码：`const x = function() { return 42; }`

---

## 四、复杂表格示例

### 4.1 基础表格

| 编号 | 语言 | 特点 | 用途 | 学习难度 |
|------|------|------|------|---------|
| 1 | Python | 简洁易读 | 数据科学、AI、Web | ⭐ |
| 2 | JavaScript | 灵活动态 | Web 开发、前端 | ⭐⭐ |
| 3 | Java | 强类型 | 企业应用、后端 | ⭐⭐⭐ |
| 4 | C++ | 高性能 | 系统编程、游戏 | ⭐⭐⭐⭐ |
| 5 | Rust | 内存安全 | 系统编程 | ⭐⭐⭐⭐⭐ |

### 4.2 对齐文本的表格

| 左对齐 | 居中 | 右对齐 |
|:------|:----:|-------:|
| 左 | 中 | 右 |
| 一个较长的左对齐文本 | 居中文本 | 一个较长的右对齐文本 |
| 短 | 文本 | 短 |

### 4.3 包含特殊字符的表格

| 功能 | 支持 | 说明 |
|------|------|------|
| Unicode | ✓ | 支持 emoji 😀🎉 |
| 特殊字符 | ✓ | `<>&"'` 等 |
| 数学符号 | ✓ | α β γ ∑ ∫ |
| 箭头 | ✓ | → ← ↑ ↓ ⇒ ⇐ |
| 其他符号 | ✓ | ™ © ® … • ◆ ★ |

### 4.4 包含链接和格式的表格

| 工具 | 链接 | 说明 |
|------|------|------|
| Pandoc | https://pandoc.org | **文档转换** 工具 |
| VS Code | https://code.visualstudio.com | 代码编辑器 |
| Git | https://git-scm.com | 版本控制 |

### 4.5 大型复杂表格

| ID | 模块 | 功能描述 | 状态 | 负责人 | 优先级 | 备注 |
|:---:|------|---------|:-----:|--------|:-------:|------|
| 1 | 用户管理 | 用户注册、登录、注销 | 完成 | Alice | ⭐⭐⭐ | 已上线 |
| 2 | 权限管理 | 角色权限、访问控制 | 进行中 | Bob | ⭐⭐⭐ | 80% 完成 |
| 3 | 数据导出 | CSV、Excel 导出 | 规划中 | Charlie | ⭐⭐ | 下周开始 |
| 4 | 报表系统 | 图表、统计、分析 | 规划中 | Diana | ⭐ | 待评估 |
| 5 | API 接口 | REST API、WebSocket | 完成 | Eve | ⭐⭐⭐⭐ | 生产就绪 |

---

## 五、引用和脚注

### 5.1 单层引用

> 这是一个简单的块引用，展示基本的引用格式。

### 5.2 多层嵌套引用

> 第一层引用：这是最外层的引用。
>
>> 第二层引用：嵌套在第一层引用内。
>>
>>> 第三层引用：更深的嵌套。
>>>
>>>> 第四层引用：继续嵌套。
>>>
>> 回到第二层：展示引用的灵活性。
>
> 回到第一层：完成演示。

### 5.3 包含其他元素的引用

> **重要通知：** 这是一个带有强调的引用。
>
> - 可以包含列表项
> - 多个列表项
>   1. 甚至有序列表
>   2. 也可以混合使用
>
> 还可以包含 `代码` 和 **格式化文本**。

### 5.4 脚注示例（如果支持）

这是一个包含脚注的文本[^1]。

这是另一个脚注[^2]。

[^1]: 这是第一个脚注的内容。
[^2]: 这是第二个脚注的内容，可以包含多行。

---

## 六、数学公式

### 6.1 行内数学

勾股定理：$a^2 + b^2 = c^2$

欧拉公式：$e^{i\pi} + 1 = 0$

积分：$\int_0^\infty e^{-x^2} dx = \frac{\sqrt{\pi}}{2}$

### 6.2 块级数学公式

二次方程的求根公式：

$$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$

傅里叶级数：

$$f(x) = \frac{a_0}{2} + \sum_{n=1}^{\infty} \left(a_n \cos\frac{n\pi x}{L} + b_n \sin\frac{n\pi x}{L}\right)$$

矩阵示例：

$$A = \begin{pmatrix}
a_{11} & a_{12} & \cdots & a_{1n} \\
a_{21} & a_{22} & \cdots & a_{2n} \\
\vdots & \vdots & \ddots & \vdots \\
a_{m1} & a_{m2} & \cdots & a_{mn}
\end{pmatrix}$$

---

## 七、特殊字符和符号

### 7.1 数学符号

- 希腊字母：α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ τ υ φ χ ψ ω
- 大写希腊字母：Α Β Γ Δ Ε Ζ Η Θ Ι Κ Λ Μ Ν Ξ Ο Π Ρ Σ Τ Υ Φ Χ Ψ Ω
- 数学符号：∀ ∃ ∄ ∅ ∈ ∉ ∋ ∌ ∩ ∪ ∑ ∏ ∫ ∂ ∇ ≈ ≠ ≤ ≥ ∞

### 7.2 箭头和几何符号

- 箭头：→ ← ↑ ↓ ↔ ↕ ⇒ ⇐ ⇑ ⇓ ⇔
- 几何符号：△ ▽ ○ ◎ ◊ ★ ☆ ♠ ♣ ♥ ♦
- 其他符号：© ® ™ ° ′ ″ · … ‰

### 7.3 表情符号（Emoji）

😀 😃 😄 😁 😆 😅 😂 🤣 😊 😇 🙂 🙃 😉 😌 😍 🥰 😘 😗 😚 😙 😜 😛 😜 🤪 😝 😕 😟 🙁 ☹️ 😲 😞 😖 😢 😭 😤 😠 😡 🤬 😈 👿 💀 🤡 👹 👺 💩 🤖 😺 😸 😹 😻 😼 😽 🙀 😿 😾

---

## 八、HTML 元素和特殊格式

### 8.1 水平线演示

上面的内容

---

中间的分割线

___

另一种分割线

***

第三种分割线

### 8.2 强调和特殊格式

这是 **加粗文本**。
这是 __也是加粗文本__。
这是 *斜体文本*。
这是 _也是斜体文本_。
这是 ***粗斜体文本***。
这是 ~~删除线~~。

<u>这是下划线（HTML）</u>

<mark>这是高亮文本（HTML）</mark>

<small>这是小号文本（HTML）</small>

### 8.3 上标和下标

H<sub>2</sub>O 是水。

E = mc<sup>2</sup> 是著名的公式。

### 8.4 组合格式

***这是粗斜体加链接*** [访问示例](https://example.com)

**这个 `代码` 和 [链接](https://example.com) 在粗体中**

---

## 九、链接和引用

### 9.1 各种链接格式

[内联链接](https://example.com)

[带标题的链接](https://example.com "示例网站")

[参考式链接][reference]

[另一个参考链接][ref2]

自动链接：<https://example.com>

邮件链接：<user@example.com>

### 9.2 定义链接引用

[reference]: https://example.com
[ref2]: https://example.com/other "另一个示例"

---

## 十、代码块中的特殊情况

### 10.1 包含引号的代码

```
"double quotes" 和 'single quotes'
```

### 10.2 包含反斜杠的代码

```
Windows 路径：C:\Users\Name\Documents
正则表达式：^\w+@\w+\.\w+$
```

### 10.3 包含代码块标记的代码

```
这里是代码块
```

### 10.4 包含 HTML 标签的代码

```html
<div class="container">
    <p>这是 HTML 代码</p>
    <script>
        console.log("JavaScript in HTML");
    </script>
</div>
```

---

## 十一、列表中的复杂元素

### 11.1 列表项中包含代码块

1. 首先安装依赖：

   ```bash
   pip install requests
   ```

2. 然后导入模块：

   ```python
   import requests
   ```

3. 最后使用 API：

   ```python
   response = requests.get('https://api.example.com/data')
   data = response.json()
   ```

### 11.2 列表项中包含表格

- 第一个项：

| 列 1 | 列 2 |
|------|------|
| 值 1 | 值 2 |

- 第二个项：

| 列 A | 列 B |
|------|------|
| 值 A | 值 B |

### 11.3 列表项中包含块引用

- 重要信息：

  > 这是块引用的重要信息
  > 跨越多行

- 次要信息：

  > 这是另一个块引用

---

## 十二、段落和文本流

### 12.1 长段落测试

这是一个较长的段落，用来测试文本换行和段落格式在转换为 Word 时的表现。在 Markdown 中，段落是由空行分隔的连续文本行。这个转换工具应该能够正确处理各种长度的段落，包括包含多个句子和复杂标点符号的段落。段落中可以包含 **加粗**、*斜体*、`代码`、[链接](https://example.com) 等各种格式元素。当转换到 Word 时，这些格式应该保持完整和正确。

### 12.2 带有列表中断的段落

这是第一个段落。

- 列表项 1
- 列表项 2

这是第二个段落。

1. 有序项 1
2. 有序项 2

这是第三个段落，在列表之后。

---

## 十三、总结和最终测试

这个高级演示文档测试了以下特性：

- ✓ **多层级标题**（从 H1 到 H6）
- ✓ **复杂列表**（嵌套、混合、任务列表）
- ✓ **多种代码块**（Python、TypeScript、Java、C++、SQL）
- ✓ **复杂表格**（对齐、特殊字符、大型）
- ✓ **嵌套引用**（多层级）
- ✓ **数学公式**（行内和块级）
- ✓ **特殊字符**（Unicode、Emoji、符号）
- ✓ **HTML 元素**（下标、上标、高亮）
- ✓ **链接和脚注**（各种形式）
- ✓ **混合复杂内容**（列表中的表格、块引用等）

### 最后的说明

本文档的目的是充分测试 Markdown 到 Word 转换工具的能力，确保各种复杂格式都能正确转换。如果所有上述元素都能在生成的 Word 文档中正确显示，那就说明这个转换工具已经相当完善了。

希望这个演示有帮助！

---

**文档完成日期**：2026-01-09

**转换工具**：Markdown to Word Converter v1.0

**Pandoc 版本**：2.x+

**测试环境**：Windows + Python 3.x + Pandoc
