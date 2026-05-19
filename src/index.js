#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import {
  getLoopByShareUrl,
  getLoopByItemId,
  listLoopContainers,
  listLoopFilesInDrive,
} from "./services/loopService.js";

// Credentials are injected as environment variables by the MCP gateway
// (stored in server_parameters, never passed through the LLM)
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

// Crea il server
const server = new Server(
  {
    name: "mcp-loop",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Lista dei tool disponibili
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: "getLoopByShareUrl",
        description:
          "Retrieve the HTML content of a Microsoft Loop file (.loop / .fluid) starting from its sharing URL. Works for Loop files stored anywhere: OneDrive, SharePoint sites, and SharePoint Embedded (Loop app workspaces).",
        inputSchema: {
          type: "object",
          properties: {
            shareUrl: {
              type: "string",
              description:
                "The full sharing URL of the Loop file (the link you would normally open in a browser).",
            },
            includeText: {
              type: "boolean",
              description:
                "If true, also returns a plain-text version of the content (HTML stripped). Default: true.",
            },
          },
          required: ["shareUrl"],
        },
      },
      {
        name: "getLoopByItemId",
        description:
          "Retrieve the HTML content of a Microsoft Loop file (.loop / .fluid) using its driveId and itemId directly. Use this when you already know the file identifiers (faster than resolving from a share URL).",
        inputSchema: {
          type: "object",
          properties: {
            driveId: { type: "string", description: "The drive ID containing the Loop file" },
            itemId: { type: "string", description: "The item ID of the Loop file" },
            includeText: {
              type: "boolean",
              description:
                "If true, also returns a plain-text version of the content (HTML stripped). Default: true.",
            },
          },
          required: ["driveId", "itemId"],
        },
      },
      {
        name: "listLoopContainers",
        description:
          "Discover all Microsoft Loop workspaces (drives) accessible to the app. Returns each workspace with its name, driveId, and number of Loop files. Use the returned driveId with listLoopFilesInDrive to enumerate files.",
        inputSchema: {
          type: "object",
          properties: {},
          required: [],
        },
      },
      {
        name: "listLoopFilesInDrive",
        description:
          "List all Loop files (.loop and .fluid) inside a specific drive (works for OneDrive, SharePoint, and SharePoint Embedded container drives).",
        inputSchema: {
          type: "object",
          properties: {
            driveId: { type: "string", description: "The drive ID to scan" },
            path: {
              type: "string",
              description:
                "Optional path inside the drive (default: 'root'). Use 'root' or a path like 'Folder/Subfolder'.",
            },
          },
          required: ["driveId"],
        },
      },
    ],
  };
});

// Gestione delle chiamate ai tool
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  switch (name) {
    case "getLoopByShareUrl":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              await getLoopByShareUrl(
                TENANT_ID,
                CLIENT_ID,
                CLIENT_SECRET,
                args.shareUrl,
                args.includeText !== false
              )
            ),
          },
        ],
      };

    case "getLoopByItemId":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              await getLoopByItemId(
                TENANT_ID,
                CLIENT_ID,
                CLIENT_SECRET,
                args.driveId,
                args.itemId,
                args.includeText !== false
              )
            ),
          },
        ],
      };

    case "listLoopContainers":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              await listLoopContainers(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
            ),
          },
        ],
      };

    case "listLoopFilesInDrive":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              await listLoopFilesInDrive(
                TENANT_ID,
                CLIENT_ID,
                CLIENT_SECRET,
                args.driveId,
                args.path
              )
            ),
          },
        ],
      };

    default:
      throw new Error(`Tool sconosciuto: ${name}`);
  }
});

// Avvia il server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Server MCP Loop avviato");
}

main().catch((error) => {
  console.error("Errore:", error);
  process.exit(1);
});
