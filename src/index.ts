import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import xlsx from "xlsx";
import fs from "fs";

// 初始化 MCP Server
const server = new McpServer({
  name: "excel-mcp",
  version: "1.0.0",
  capabilities: {
    tools: {},
    resources: {},
  },
});

// 工具：读取 Excel 文件
server.tool(
  "read_excel",
  "读取并理解一个 Excel 文件（包含表名、表头和前 10 行数据）",
  {
    filePath: z.string().describe(".xlsx Excel 文件的路径"),
  },
  async ({ filePath }) => {
    // 检查文件是否存在
    if (!fs.existsSync(filePath)) {
      return {
        content: [
          {
            type: "text",
            text: `❌ 未找到文件: ${filePath}`,
          },
        ],
      };
    }

    try {
      // 读取 Excel 文件
      const workbook = xlsx.readFile(filePath);
      const sheetNames = workbook.SheetNames;
      const responseChunks: string[] = [];

      responseChunks.push(`📄 检测到工作表: ${sheetNames.join(", ")}`);

      for (const sheetName of sheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1 }) as (string | number)[][];

        const headerRow = jsonData[0] || [];
        const previewRows = jsonData.slice(1, 11);

        responseChunks.push(`\n🔹 工作表: ${sheetName}`);
        responseChunks.push(`表头: ${headerRow.join(", ")}`);
        responseChunks.push(`前 ${previewRows.length} 行数据:`);

        for (const row of previewRows) {
          responseChunks.push(row.map((cell) => (cell === undefined ? "" : String(cell))).join(" | "));
        }

        responseChunks.push("---");
      }

      return {
        content: [
          {
            type: "text",
            text: responseChunks.join("\n"),
          },
        ],
      };
    } catch (error) {
      console.error("读取 Excel 时出错:", error);
      return {
        content: [
          {
            type: "text",
            text: "❌ 读取 Excel 文件失败。请确保该文件是有效的 .xlsx 文件。",
          },
        ],
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
