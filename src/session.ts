import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ErrorCode,
  McpError,
  JSONRPCMessage
} from "@modelcontextprotocol/sdk/types.js";
import { getAccessToken } from "./auth";
import { Env } from "./types";
import { getDocumentBody, createDocument, searchDocuments, replaceSection, appendText, listSections, getDocumentInfo, findAndReplace, listDocuments, deleteSection, readSection, insertSection, renameSection } from "./google-api";
import { listSpreadsheets, readSheetRange, createSpreadsheet, writeSheetRange, appendSheetRows } from "./sheets-api";
import { docToMarkdown } from "./markdown-utils";
import { checkRateLimit, tokenTag } from "./utils";
import { CloudflareSSEServerTransport } from "./transports/sse";
import { StreamableHttpTransport } from "./transports/streamable";

export class MCPSession {
  // SSE transport state
  private server: Server | null = null;
  private transport: CloudflareSSEServerTransport | null = null;

  // Streamable HTTP transport state (lazy-initialised, lives for the DO's lifetime)
  private streamableServer: Server | null = null;
  private streamableTransport: StreamableHttpTransport | null = null;

  private state: DurableObjectState;
  private env: Env;

  constructor(state: DurableObjectState, env: Env) {
    this.state = state;
    this.env = env;
    console.log("[DO Lifecycle] Constructor called! Instance started.");
  }

  // ─── Tool definitions ────────────────────────────────────────────────────────

  private static readonly toolDefinitions = [
    {
      name: "read_document",
      description: "Read a Google Document and return its full content as Markdown",
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" }
        },
        required: ["documentId"]
      }
    },
    {
      name: "create_document",
      description: "Create a new, empty Google Document and return its ID",
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          title: { type: "string", description: "The title of the new Google Document" }
        },
        required: ["title"]
      }
    },
    {
      name: "search_documents",
      description: "Search Google Drive for Documents matching a query by title or content",
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          query: { type: "string", description: "The search query to match against document names and content" }
        },
        required: ["query"]
      }
    },
    {
      name: "edit_section",
      description: "Replace the content under a specific section header in a Google Document. The header itself is preserved; only the content beneath it is replaced.",
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" },
          headerText: { type: "string", description: "The exact header text identifying the section to replace" },
          newContent: { type: "string", description: "The Markdown content to insert beneath the header" }
        },
        required: ["documentId", "headerText", "newContent"]
      }
    },
    {
      name: "append_text",
      description: "Append Markdown content to the very end of a Google Document",
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" },
          newContent: { type: "string", description: "The Markdown content to append" }
        },
        required: ["documentId", "newContent"]
      }
    },
    {
      name: "list_sections",
      description: "List all section headers in a Google Document with their heading level and position. Use this before edit_section or delete_section to confirm exact header text.",
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" }
        },
        required: ["documentId"]
      }
    },
    {
      name: "get_document_info",
      description: "Return metadata for a Google Document: title, document ID, revision ID, last modified time, and file size. Does not return document body content.",
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" }
        },
        required: ["documentId"]
      }
    },
    {
      name: "find_and_replace",
      description: "Find all occurrences of a text string in a Google Document and replace them. Returns the number of replacements made.",
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" },
          findText: { type: "string", description: "The text to search for" },
          replaceText: { type: "string", description: "The text to replace each match with" },
          matchCase: { type: "boolean", description: "Whether the search is case-sensitive (default: false)" }
        },
        required: ["documentId", "findText", "replaceText"]
      }
    },
    {
      name: "list_documents",
      description: "List the user's most recently modified Google Documents from Drive. Returns up to 20 documents with their IDs, names, and last modified times.",
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {},
        required: []
      }
    },
    {
      name: "delete_section",
      description: "Permanently delete a section from a Google Document — the header and all content beneath it up to the next same-or-higher-level heading. This action cannot be undone.",
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" },
          headerText: { type: "string", description: "The exact header text of the section to delete" }
        },
        required: ["documentId", "headerText"]
      }
    },
    {
      name: "read_section",
      description: "Read the content of a specific section in a Google Document and return it as Markdown. Includes the heading and all content beneath it up to the next same-or-higher-level heading.",
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" },
          headerText: { type: "string", description: "The exact header text of the section to read" }
        },
        required: ["documentId", "headerText"]
      }
    },
    {
      name: "insert_section",
      description: "Insert a new section immediately before or after an existing section in a Google Document. The new content should include its own heading.",
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" },
          headerText: { type: "string", description: "The exact header text of the reference section" },
          position: { type: "string", enum: ["before", "after"], description: "Whether to insert before or after the reference section" },
          newContent: { type: "string", description: "The Markdown content to insert, including its own heading" }
        },
        required: ["documentId", "headerText", "position", "newContent"]
      }
    },
    {
      name: "list_spreadsheets",
      description: "List the user's available Google Spreadsheets from Drive. Returns up to 20 sheets with their IDs and names.",
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {},
        required: []
      }
    },
    {
      name: "read_sheet_range",
      description: "Read the values from a specific range in a Google Spreadsheet.",
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          spreadsheetId: { type: "string", description: "The ID of the Google Spreadsheet" },
          range: { type: "string", description: "The range to read, e.g., 'Sheet1!A1:B10'" }
        },
        required: ["spreadsheetId", "range"]
      }
    },
    {
      name: "create_spreadsheet",
      description: "Create a new Google Spreadsheet and return its ID.",
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          title: { type: "string", description: "The title of the new Spreadsheet" }
        },
        required: ["title"]
      }
    },
    {
      name: "write_sheet_range",
      description: "Write values to a specific range in a Google Spreadsheet. Overwrites existing values.",
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          spreadsheetId: { type: "string", description: "The ID of the Google Spreadsheet" },
          range: { type: "string", description: "The range to write, e.g., 'Sheet1!A1:B10'" },
          values: { type: "array", items: { type: "array", items: { type: "string" } }, description: "The 2D array of values to write" }
        },
        required: ["spreadsheetId", "range", "values"]
      }
    },
    {
      name: "append_sheet_rows",
      description: "Append rows of values to the end of a sheet in a Google Spreadsheet.",
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          spreadsheetId: { type: "string", description: "The ID of the Google Spreadsheet" },
          range: { type: "string", description: "The range where to append, e.g., 'Sheet1!A1'" },
          values: { type: "array", items: { type: "array", items: { type: "string" } }, description: "The 2D array of values to append" }
        },
        required: ["spreadsheetId", "range", "values"]
      }
    },
    {
      name: "rename_section",
      description: "Rename the heading of an existing section in a Google Document. The heading level (H1, H2, etc.) is preserved; only the text changes.",
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
      inputSchema: {
        type: "object" as const,
        properties: {
          documentId: { type: "string", description: "The ID or full URL of the Google Document" },
          headerText: { type: "string", description: "The exact current heading text of the section to rename" },
          newHeaderText: { type: "string", description: "The new heading text" }
        },
        required: ["documentId", "headerText", "newHeaderText"]
      }
    }
  ];

  // ─── Argument validation ──────────────────────────────────────────────────��──

  private static requireString(args: Record<string, unknown>, key: string, maxLen = 10_000): string {
    const val = args[key];
    if (typeof val !== "string" || val.trim() === "") {
      throw new McpError(ErrorCode.InvalidParams, `"${key}" must be a non-empty string`);
    }
    if (val.length > maxLen) {
      throw new McpError(ErrorCode.InvalidParams, `"${key}" exceeds maximum length of ${maxLen} characters`);
    }
    return val;
  }

  // ─── Tool call dispatcher (shared between SSE and Streamable HTTP) ───────────

  private async dispatchToolCall(name: string, args: Record<string, unknown>): Promise<{ content: { type: string; text: string }[] }> {
    const userToken = await this.state.storage.get<string>("userToken");
    if (!userToken) throw new McpError(ErrorCode.InvalidRequest, "Session not initialised");

    const accessToken = await getAccessToken(userToken, this.env);
    if (!accessToken) throw new McpError(ErrorCode.InvalidRequest, "Not authenticated or token expired");

    if (!await checkRateLimit(this.env.TOKENS, `google:${userToken}`, 100, 60)) {
      throw new McpError(ErrorCode.InvalidRequest, "Google API rate limit exceeded. Please wait before retrying.");
    }

    const req = MCPSession.requireString;

    switch (name) {
      case "read_document": {
        const doc = await getDocumentBody(req(args, "documentId"), accessToken);
        return { content: [{ type: "text", text: docToMarkdown(doc) }] };
      }
      case "create_document": {
        const documentId = await createDocument(req(args, "title", 1000), accessToken);
        return { content: [{ type: "text", text: `Document created with ID: ${documentId}` }] };
      }
      case "search_documents": {
        const results = await searchDocuments(req(args, "query", 500), accessToken);
        return { content: [{ type: "text", text: JSON.stringify(results, null, 2) }] };
      }
      case "edit_section": {
        const headerText = req(args, "headerText", 500);
        await replaceSection(req(args, "documentId"), headerText, req(args, "newContent"), accessToken);
        return { content: [{ type: "text", text: `Section "${headerText}" updated successfully.` }] };
      }
      case "append_text": {
        await appendText(req(args, "documentId"), req(args, "newContent"), accessToken);
        return { content: [{ type: "text", text: "Text successfully appended to the document." }] };
      }
      case "list_sections": {
        const doc = await getDocumentBody(req(args, "documentId"), accessToken);
        const sections = listSections(doc);
        if (sections.length === 0) {
          return { content: [{ type: "text", text: "No sections (headings) found in this document." }] };
        }
        const formatted = sections
          .map(s => `${"  ".repeat(s.level - 1)}${"#".repeat(s.level)} ${s.text}`)
          .join("\n");
        return { content: [{ type: "text", text: formatted }] };
      }
      case "get_document_info": {
        const info = await getDocumentInfo(req(args, "documentId"), accessToken);
        return { content: [{ type: "text", text: JSON.stringify(info, null, 2) }] };
      }
      case "find_and_replace": {
        const matchCase = typeof args.matchCase === "boolean" ? args.matchCase : false;
        const count = await findAndReplace(
          req(args, "documentId"),
          req(args, "findText", 10_000),
          req(args, "replaceText", 10_000),
          matchCase,
          accessToken
        );
        return { content: [{ type: "text", text: `Replaced ${count} occurrence${count !== 1 ? "s" : ""}.` }] };
      }
      case "list_documents": {
        const result = await listDocuments(accessToken);
        const files = result.files ?? [];
        if (files.length === 0) {
          return { content: [{ type: "text", text: "No documents found." }] };
        }
        return { content: [{ type: "text", text: JSON.stringify(files, null, 2) }] };
      }
      case "delete_section": {
        const headerText = req(args, "headerText", 500);
        await deleteSection(req(args, "documentId"), headerText, accessToken);
        return { content: [{ type: "text", text: `Section "${headerText}" deleted successfully.` }] };
      }
      case "read_section": {
        const headerText = req(args, "headerText", 500);
        const content = await readSection(req(args, "documentId"), headerText, accessToken);
        return { content: [{ type: "text", text: content }] };
      }
      case "insert_section": {
        const headerText = req(args, "headerText", 500);
        const positionRaw = req(args, "position", 10);
        if (positionRaw !== "before" && positionRaw !== "after") {
          throw new McpError(ErrorCode.InvalidParams, '"position" must be "before" or "after"');
        }
        await insertSection(req(args, "documentId"), headerText, positionRaw, req(args, "newContent"), accessToken);
        return { content: [{ type: "text", text: `New section inserted ${positionRaw} "${headerText}".` }] };
      }
      case "rename_section": {
        const headerText = req(args, "headerText", 500);
        const newHeaderText = req(args, "newHeaderText", 500);
        await renameSection(req(args, "documentId"), headerText, newHeaderText, accessToken);
        return { content: [{ type: "text", text: `Section renamed from "${headerText}" to "${newHeaderText}".` }] };
      }
      case "list_spreadsheets": {
        const result = await listSpreadsheets(userToken, this.env);
        return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
      }
      case "read_sheet_range": {
        const result = await readSheetRange(userToken, this.env, req(args, "spreadsheetId"), req(args, "range"));
        return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
      }
      case "create_spreadsheet": {
        const result = await createSpreadsheet(userToken, this.env, req(args, "title"));
        return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
      }
      case "write_sheet_range": {
        const result = await writeSheetRange(userToken, this.env, req(args, "spreadsheetId"), req(args, "range"), args.values);
        return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
      }
      case "append_sheet_rows": {
        const result = await appendSheetRows(userToken, this.env, req(args, "spreadsheetId"), req(args, "range"), args.values);
        return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
      }
      default:
        throw new McpError(ErrorCode.MethodNotFound, `Tool not found: ${name}`);
    }
  }

  // ─── Shared handler registration ─────────────────────────────────────────────

  private registerHandlers(server: Server) {
    server.setRequestHandler(ListToolsRequestSchema, async () => {
      this.log("info", "tools/list request received");

      const userToken = await this.state.storage.get<string>("userToken");
      if (!userToken) throw new McpError(ErrorCode.InvalidRequest, "Session not initialised");
      const valid = await this.env.TOKENS.get(userToken);
      if (!valid) throw new McpError(ErrorCode.InvalidRequest, "Session revoked or expired");

      return { tools: MCPSession.toolDefinitions };
    });

    server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        const rawArgs = request.params.arguments;
        if (rawArgs !== undefined && rawArgs !== null &&
            (typeof rawArgs !== "object" || Array.isArray(rawArgs))) {
          throw new McpError(ErrorCode.InvalidParams, "Tool arguments must be a plain object");
        }
        return await this.dispatchToolCall(
          request.params.name,
          (rawArgs ?? {}) as Record<string, unknown>
        );
      } catch (error: any) {
        if (error instanceof McpError) throw error;
        console.error(`MCP Tool Error [${request.params.name}]:`, error.message, error.stack);
        throw new McpError(ErrorCode.InternalError, "An internal error occurred. Please try again.");
      }
    });
  }

  // ─── SSE server lifecycle ─────────────────────────────────────────────────────
  // Two separate Server instances are intentional: the MCP SDK binds a Server to
  // exactly one transport at a time. SSE uses a persistent stream (re-created on
  // each new client connection), while Streamable HTTP is request-response (lives
  // for the DO lifetime). Both call the same registerHandlers(), so adding a tool
  // requires only one change in toolDefinitions + dispatchToolCall.

  private initializeSseServer(): Server {
    const server = new Server(
      { name: "google-docs-mcp", version: "1.0.0" },
      { capabilities: { tools: { listChanged: true } } }
    );
    this.registerHandlers(server);
    this.server = server;
    return server;
  }

  // ─── Streamable HTTP server lifecycle ─────────────────────────────────────────

  private async ensureStreamableServer() {
    if (this.streamableServer && this.streamableTransport) return;

    this.streamableServer = new Server(
      { name: "google-docs-mcp", version: "1.0.0" },
      { capabilities: { tools: { listChanged: true } } }
    );
    this.registerHandlers(this.streamableServer);
    this.streamableTransport = new StreamableHttpTransport();
    await this.streamableServer.connect(this.streamableTransport);
    this.log("info", "Streamable HTTP server initialised");
  }

  // ─── Streamable HTTP request handler ─────────────────────────────────────────

  private async handleStreamablePost(request: Request, userToken: string): Promise<Response> {
    await this.ensureStreamableServer();

    const MAX_BODY = 512 * 1024;

    let rawBody: string;
    try {
      rawBody = await request.text();
    } catch {
      return new Response(
        JSON.stringify({ jsonrpc: "2.0", id: null, error: { code: -32700, message: "Parse error" } }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    if (new TextEncoder().encode(rawBody).length > MAX_BODY) {
      return new Response(
        JSON.stringify({ jsonrpc: "2.0", id: null, error: { code: -32600, message: "Payload too large" } }),
        { status: 413, headers: { "Content-Type": "application/json" } }
      );
    }

    let body: JSONRPCMessage;
    try {
      body = JSON.parse(rawBody) as JSONRPCMessage;
    } catch {
      return new Response(
        JSON.stringify({ jsonrpc: "2.0", id: null, error: { code: -32700, message: "Parse error" } }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    const response = await this.streamableTransport!.processMessage(body);

    if (response === null) {
      return new Response(null, { status: 202 });
    }

    return new Response(JSON.stringify(response), {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        "Mcp-Session-Id": await tokenTag(userToken, 16),
      }
    });
  }

  // ─── Logging helper ───────────────────────────────────────────────────────────

  private log(level: "debug" | "info" | "warning" | "error", message: string, data?: any) {
    const logData = data ? `${message} ${JSON.stringify(data)}` : message;
    if (level === "error") {
      console.error(`[MCP ${level}]`, logData);
    } else if (level === "warning") {
      console.warn(`[MCP ${level}]`, logData);
    } else {
      console.log(`[MCP ${level}]`, logData);
    }
  }

  // ─── Durable Object fetch handler ────────────────────────────────────────────

  async fetch(request: Request): Promise<Response> {
    try {
      const url = new URL(request.url);
      this.log("debug", `DO request: ${request.method} ${url.pathname}`);

      // ── Streamable HTTP (MCP 2025-03-26) ──
      if (url.pathname.endsWith("/streamable")) {
        const userToken = url.searchParams.get("userToken");
        if (!userToken) return new Response("Missing userToken", { status: 401 });

        if (request.method === "GET") {
          return new Response("Use POST for Streamable HTTP", { status: 405 });
        }

        if (request.method === "POST") {
          const storedToken = await this.state.storage.get<string>("userToken");
          if (storedToken !== userToken) {
            await this.state.storage.put("userToken", userToken);
          }
          return this.handleStreamablePost(request, userToken);
        }

        return new Response("Method Not Allowed", { status: 405 });
      }

      // ── Legacy SSE transport ──
      if (url.pathname.endsWith("/sse")) {
        const userToken = url.searchParams.get("userToken");
        if (userToken) {
          const valid = await this.env.TOKENS.get(userToken);
          if (!valid) return new Response("Invalid or revoked token", { status: 401 });
          await this.state.storage.put("userToken", userToken);
          this.log("info", "Stored userToken for SSE session");
        }

        if (this.transport && this.server) {
          this.log("info", "Closing existing SSE transport");
          try { await this.server.close(); } catch {}
          this.transport = null;
          this.server = null;
        }

        const server = this.initializeSseServer();

        const messagesEndpoint = userToken
          ? `${url.origin}/mcp/messages?token=${userToken}`
          : `${url.origin}/mcp/messages`;

        this.transport = new CloudflareSSEServerTransport(messagesEndpoint);
        this.transport.onclose = () => this.log("info", "SSE transport closed");

        await server.connect(this.transport);
        this.log("info", "SSE server connected");
        return this.transport.handleSseRequest();
      }

      if (url.pathname.endsWith("/messages")) {
        if (!this.transport || !this.server) {
          this.log("error", "POST /messages: no active SSE session");
          return new Response("No active SSE session. Please re-establish the SSE connection.", { status: 400 });
        }
        return this.transport.handlePostRequest(request);
      }

      this.log("warning", `Path not found in DO: ${url.pathname}`);
      return new Response("Not Found", { status: 404 });

    } catch (e: any) {
      this.log("error", `DO Fetch Error: ${e.message}`, { stack: e.stack });
      return new Response(JSON.stringify({ error: "Internal Server Error" }), {
        status: 500,
        headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }
      });
    }
  }
}
