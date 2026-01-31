// TypeScript LSP 测试文件

interface User {
  id: number;
  name: string;
  email: string;
}

function greet(user: User): string {
  return `Hello, ${user.name}!`;
}

// 测试类型检查 - 故意写错类型
const user: User = {
  id: "123",  // 错误: 应该是 number
  name: "张三",
  email: "test@example.com"
};

// 测试未定义变量
console.log(undefinedVariable);

// 正确的用法
const validUser: User = {
  id: 1,
  name: "李四",
  email: "lisi@example.com"
};

console.log(greet(validUser));

// === 更多 LSP 测试 ===

// 测试: 参数类型错误
function add(a: number, b: number): number {
  return a + b;
}
add("1", 2);  // 错误: 第一个参数应该是 number

// 测试: 缺少必需属性
const incompleteUser: User = {
  id: 2,
  name: "王五"
  // 缺少 email 属性
};

// 测试: 多余属性
const extraUser: User = {
  id: 3,
  name: "赵六",
  email: "zhaoliu@example.com",
  age: 25  // 错误: User 接口没有 age 属性
};

// 测试: 返回类型不匹配
function getNumber(): number {
  return "not a number";  // 错误: 返回类型应该是 number
}

// 测试: 调用不存在的方法
validUser.sayHello();  // 错误: User 没有 sayHello 方法
