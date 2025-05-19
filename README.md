当然可以！以下是中文版本的 `README.md`，适用于你的 `excel-mcp` 项目：

---

### 📄 `README.md`

````markdown
# 📊 excel-mcp

`excel-mcp` 是一个基于 [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) 的服务器，使用 TypeScript 编写，具备读取 Excel 文件（`.xlsx` 格式）的功能，并通过 MCP 协议对接交互式工具。

---

## ✨ 功能特点

- 📁 支持读取本地 `.xlsx` 文件；
- 📑 解析 Excel 中的所有工作表（Sheet）；
- 🧾 提取每个表的表头信息；
- 🔍 展示每个表前 10 行数据；
- 🤖 可通过 MCP 协议进行交互式调用。

---

## 🚀 快速开始

### 1. 安装依赖

```bash
npm install
````

### 2. 构建项目

```bash
npm run build
```

### 3. 启动 MCP Server

```bash
./build/index.js
```

MCP Server 会通过标准输入/输出（stdio）监听来自 Model Context 的请求。

---

## 🛠 可用工具

### `read_excel`

读取并理解指定路径下的 Excel 文件。

#### 输入参数：

| 参数名        | 类型     | 描述                   |
| ---------- | ------ | -------------------- |
| `filePath` | string | Excel 文件的完整路径（.xlsx） |

#### 返回内容：

* 所有工作表名称；
* 每个 Sheet 的表头（第一行）；
* 每个 Sheet 的前 10 行数据（按行格式化展示）；

#### 示例请求：

```json
{
  "tool": "read_excel",
  "input": {
    "filePath": "/absolute/path/to/example.xlsx"
  }
}
```
