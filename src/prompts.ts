import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

export function registerPrompts(server: McpServer): void {
  server.registerPrompt(
    "onenote_report",
    {
      title: "Build a report from OneNote",
      description: "Gathers notes (text and screenshots) from a notebook or section and turns them into a report.",
      argsSchema: {
        scope: z.string().describe('Notebook and/or section to report on, e.g. "Work / Weekly status".'),
        period: z.string().optional().describe('Time window, e.g. "last 7 days" or "September 2026".'),
        focus: z.string().optional().describe("What the report should emphasise or answer."),
        destination: z
          .string()
          .optional()
          .describe('Where to put the report: "chat" (default) or a OneNote "Notebook / Section" to create a page in.'),
      },
    },
    ({ scope, period, focus, destination }) => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: [
              `Create a report from my OneNote notes.`,
              `Scope: ${scope}`,
              `Period: ${period ?? "all recent notes"}`,
              focus ? `Focus: ${focus}` : "",
              `Destination: ${destination ?? "chat"}`,
              "",
              "Steps:",
              "1. Use onenote_list_notebooks (or onenote_list_sections) to find the notebook/section IDs for the scope.",
              `2. Use onenote_collect_pages with that scope${period ? " and a modified_since date matching the period" : ""} to read the pages.`,
              "3. For pages whose screenshots matter (dashboards, tables, error messages, charts), read them with onenote_get_page, " +
                "and zoom into dense screenshots with onenote_get_page_image and a region. Pull exact figures from them.",
              "4. Write the report: a short executive summary, key findings with figures, decisions and action items " +
                "(owner and date where known), risks or open questions, and a sources list naming each page title " +
                "(and image number for facts taken from screenshots). Flag anything you could not read clearly.",
              destination && destination.toLowerCase() !== "chat"
                ? `5. Save it with onenote_create_page in "${destination}" (use notebook_name/section_name; ask before creating a new section). ` +
                  "Use markdown with tables for figures, and embed key screenshots with ![caption](onenote-image://PAGE_ID/N)."
                : "5. Show the report here in chat, then ask whether to save it to OneNote.",
            ]
              .filter(Boolean)
              .join("\n"),
          },
        },
      ],
    }),
  );

  server.registerPrompt(
    "onenote_extract_screenshots",
    {
      title: "Extract data from OneNote screenshots",
      description: "Reads every screenshot on a page and transcribes the text, tables and figures in them.",
      argsSchema: {
        page: z.string().describe("Page title (or page ID) to process."),
        instructions: z.string().optional().describe("What to extract, e.g. 'all error codes' or 'the KPI table'."),
      },
    },
    ({ page, instructions }) => ({
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: [
              `Extract the information in the screenshots on my OneNote page "${page}".`,
              instructions ? `Focus on: ${instructions}` : "",
              "",
              "Steps:",
              "1. Find the page with onenote_search_pages (or use it directly if it is an ID).",
              "2. Read it with onenote_get_page. For each image, if any text is too small to read with certainty, " +
                "call onenote_get_page_image with a region (e.g. quadrants) to zoom in.",
              "3. For each image, give: what it shows, a faithful transcription of its text, and any tables as markdown tables.",
              "4. Mark uncertain readings with [?] rather than guessing. Finish with a short summary of the key facts.",
            ]
              .filter(Boolean)
              .join("\n"),
          },
        },
      ],
    }),
  );
}
