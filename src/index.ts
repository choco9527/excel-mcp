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
        "\nℹ️ 你可以继续使用 'process_excel' 工具来提取特定列或过滤数据。例如，指定列索引、过滤值或工作表名称。"
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

// // 工具：处理已读取的 Excel/CSV 数据
// server.tool(
//   "process_excel",
//   "对已读取的 Excel 或 CSV 数据进行进一步处理（例如提取列、过滤行）",
//   {
//     filePath: z.string().describe("之前读取的 .xlsx 或 .csv 文件路径"),
//     sheetName: z.string().optional().describe("Excel 工作表名称，默认为第一个工作表或 CSV 的默认表"),
//     action: z.enum(["extract_column", "filter_rows"]).describe("操作类型：提取列或过滤行"),
//     columnIndex: z.number().optional().describe("要提取的列索引（从 0 开始），用于 extract_column"),
//     filterValue: z.string().optional().describe("过滤行的值，包含该值的行将被返回，用于 filter_rows"),
//   },
//   async ({ filePath, sheetName, action, columnIndex, filterValue }) => {
//     try {
//       // 检查缓存或重新读取文件
//       let data = sessionCache.get(filePath);
//       if (!data) {
//         if (!fs.existsSync(filePath)) {
//           return {
//             content: [{ type: "text", text: `❌ 未找到文件: ${filePath}` }],
//           };
//         }
//         if (filePath.endsWith(".xlsx") || filePath.endsWith(".csv")) {
//           const workbook = xlsx.readFile(filePath, filePath.endsWith(".csv") ? { raw: true } : undefined);
//           const sheetNames = workbook.SheetNames;
//           data = { filePath, data: {} };
//           for (const name of sheetNames) {
//             const sheet = workbook.Sheets[name];
//             data.data[name] = xlsx.utils.sheet_to_json(sheet, { header: 1 });
//           }
//           sessionCache.set(filePath, data);
//         } else {
//           return {
//             content: [{ type: "text", text: "❌ 仅支持 .xlsx 或 .csv 文件。" }],
//           };
//         }
//       }

//       // 选择目标sheet
//       const availableSheets = Object.keys(data.data);
//       let targetSheet = sheetName && availableSheets.includes(sheetName) ? sheetName : availableSheets[0];
//       if (!data.data[targetSheet]) {
//         return {
//           content: [{ type: "text", text: `❌ 工作表 ${sheetName} 不存在。可用工作表: ${availableSheets.join(", ")}` }],
//         };
//       }

//       const jsonData = data.data[targetSheet] as (string | number)[][];
//       const headerRow = jsonData[0] || [];
//       const rows = jsonData.slice(1);

//       // 预览数据（仅首次读取时展示）
//       if (!sessionCache.has(filePath)) {
//         const responseChunks: string[] = [];
//         responseChunks.push(`📄 检测到${filePath.endsWith(".csv") ? "CSV文件" : "工作表"}: ${filePath.endsWith(".csv") ? filePath : availableSheets.join(", ")}`);
//         responseChunks.push(`\n🔹 ${filePath.endsWith(".csv") ? "文件名" : "工作表"}: ${targetSheet}`);
//         responseChunks.push(`表头: ${headerRow.join(", ")}`);
//         const previewRows = jsonData.slice(1, 11);
//         responseChunks.push(`前 ${previewRows.length} 行数据:`);
//         for (const row of previewRows) {
//           responseChunks.push(row.map(cell => (cell === undefined ? "" : String(cell))).join(" | "));
//         }
//         responseChunks.push("---");
//         return {
//           content: [{ type: "text", text: responseChunks.join("\n") }],
//         };
//       }

//       // 根据 action 处理数据
//       const responseChunks: string[] = [];
//       if (action === "extract_column") {
//         if (columnIndex === undefined || columnIndex >= headerRow.length) {
//           return {
//             content: [{ type: "text", text: `❌ 无效的列索引: ${columnIndex}` }],
//           };
//         }
//         const columnData = rows.map(row => row[columnIndex] ?? "").filter(cell => cell !== "");
//         responseChunks.push(`🔹 提取列: ${headerRow[columnIndex]} (工作表: ${targetSheet})`);
//         responseChunks.push(columnData.join("\n"));
//       } else if (action === "filter_rows") {
//         if (!filterValue) {
//           return {
//             content: [{ type: "text", text: "❌ 缺少过滤值" }],
//           };
//         }
//         const filteredRows = rows.filter(row =>
//           row.some(cell => String(cell).includes(filterValue))
//         );
//         responseChunks.push(`🔹 过滤包含 "${filterValue}" 的行 (工作表: ${targetSheet}):`);
//         for (const row of filteredRows) {
//           responseChunks.push(row.map(cell => (cell === undefined ? "" : String(cell))).join(" | "));
//         }
//       }

//       return {
//         content: [{ type: "text", text: responseChunks.join("\n") }],
//       };
//     } catch (error) {
//       console.error("处理数据时出错:", error);
//       return {
//         content: [{ type: "text", text: "❌ 处理数据失败。" }],
//       };
//     }
//   }
// );

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
