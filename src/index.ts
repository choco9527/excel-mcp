import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import xlsx from "xlsx";
import fs from "fs";
// 内存缓存，存储已读取的数据
const sessionCache = new Map<string, { filePath: string; data: { [sheetName: string]: (string | number)[][] } }>();

// 初始化 MCP Server
const server = new McpServer({
  name: "excel-mcp",
  version: "1.0.0",
});

// 工具：读取 Excel 或 CSV 文件
server.tool(
  "read_excel",
  "读取并理解一个 Excel 或 CSV 文件（包含表名/文件名、表头和前 10 行数据）",
  {
    filePath: z.string().describe(".xlsx 或 .csv 文件的路径"),
    sheetName: z.string().optional().describe("Excel 工作表名称，默认为第一个工作表"),
  },
  async ({ filePath, sheetName }) => {
    if (!fs.existsSync(filePath)) {
      return {
        content: [{ type: "text", text: `❌ 未找到文件: ${filePath}` }],
      };
    }

    try {
      const responseChunks: string[] = [];
      let data: { filePath: string; data: { [sheetName: string]: (string | number)[][] } };

      if (filePath.endsWith(".xlsx")) {
        const workbook = xlsx.readFile(filePath);
        const sheetNames = workbook.SheetNames;
        responseChunks.push(`📄 检测到工作表: ${sheetNames.join(", ")}`);

        // 如果未指定 sheetName，默认使用第一个工作表
        const targetSheet = sheetName && sheetNames.includes(sheetName) ? sheetName : sheetNames[0];
        data = { filePath, data: {} };

        // 读取指定工作表（或所有工作表以缓存）
        for (const name of sheetNames) {
          const sheet = workbook.Sheets[name];
          data.data[name] = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        }

        // 显示指定工作表的数据
        const sheetData = data.data[targetSheet];
        const headerRow = sheetData[0] || [];
        const previewRows = sheetData.slice(1, 11);
        responseChunks.push(`\n🔹 工作表: ${targetSheet}`);
        responseChunks.push(`表头: ${headerRow.join(", ")}`);
        responseChunks.push(`前 ${previewRows.length} 行数据:`);
        for (const row of previewRows) {
          responseChunks.push(row.map(cell => (cell === undefined ? "" : String(cell))).join(" | "));
        }
        responseChunks.push("---");
      } else if (filePath.endsWith(".csv")) {
        const workbook = xlsx.readFile(filePath, { raw: true });
        const sheetNames = workbook.SheetNames;
        data = { filePath, data: {} };
        // CSV文件只有一个sheet
        const sheet = workbook.Sheets[sheetNames[0]];
        data.data["default"] = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        const headerRow = data.data["default"][0] || [];
        const previewRows = data.data["default"].slice(1, 11);
        responseChunks.push(`📄 检测到 CSV 文件: ${filePath}`);
        responseChunks.push(`\n🔹 文件名: ${filePath}`);
        responseChunks.push(`表头: ${headerRow.join(", ")}`);
        responseChunks.push(`前 ${previewRows.length} 行数据:`);
        for (const row of previewRows) {
          responseChunks.push(row.map(cell => (cell === undefined ? "" : String(cell))).join(" | "));
        }
        responseChunks.push("---");
      } else {
        return {
          content: [{ type: "text", text: "❌ 仅支持 .xlsx 或 .csv 文件。" }],
        };
      }

      // 缓存数据
      sessionCache.set(filePath, data);

      // 添加提示，引导模型继续交互
      responseChunks.push(
        "\n✅ 文件读取完成！你可以直接描述你的需求，例如：‘请找出所有24号点了早餐的人’，我将自动为你生成并执行 Node.js 脚本来完成你的需求。无需手动编写代码，只需用自然语言告诉我你想要的数据处理结果。"
      );
      return {
        content: [{ type: "text", text: responseChunks.join("\n") }],
      };
    } catch (error) {
      console.error("读取文件时出错:", error);
      return {
        content: [{ type: "text", text: "❌ 读取文件失败。请确保该文件是有效的 .xlsx 或 .csv 文件。" }],
      };
    }
  }
);


// 启动服务主函数
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Excel MCP 服务已通过 stdio 启动。");
}

main().catch((err) => {
  console.error("致命错误:", err);
  process.exit(1);
});
