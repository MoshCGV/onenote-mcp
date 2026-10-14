// End-to-end test: runs the built MCP server over stdio against a local mock of Microsoft Graph.
import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import http from "node:http";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import sharp from "sharp";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const serverEntry = path.join(here, "..", "dist", "index.js");
const TOKEN = "test-token";

let server;
let port;
let tmp;
let client;
const requests = [];
let r2Hits = 0;

async function textPng(width, height, label) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
    <rect width="100%" height="100%" fill="#ffffff"/>
    <rect x="20" y="20" width="${width - 40}" height="${Math.round(height / 5)}" fill="#1e4e79"/>
    <text x="40" y="${Math.round(height / 2)}" font-size="${Math.round(height / 12)}" font-family="Arial" fill="#000">${label}</text>
  </svg>`;
  return sharp(Buffer.from(svg)).png().toBuffer();
}

function root() {
  return `http://127.0.0.1:${port}/v1.0/me/onenote`;
}

function page1Html() {
  const R = root();
  return `<html lang="en-US"><head><title>Sprint 42 status</title><meta name="created" content="2026-10-10T09:00:00.0000000" /></head>
<body data-absolute-enabled="true" style="font-family:Calibri;font-size:11pt">
<div style="position:absolute;left:48px;top:300px;width:624px"><p>Second outline: deploy notes</p>
<img alt="Dashboard screenshot" width="1500" height="1000" src="${R}/resources/r1-web/$value" data-src-type="image/png" data-fullres-src="${R}/resources/r1/$value" data-fullres-src-type="image/png" /></div>
<div data-id="_default" style="position:absolute;left:48px;top:120px;width:624px">
<h1 style="font-size:16pt;color:#1e4e79">Summary</h1>
<p>Velocity was <span style="font-weight:bold">42 points</span> this sprint.</p>
<p data-tag="to-do">Follow up with QA</p>
<p data-tag="to-do:completed">Ship build 7.1</p>
<table style="border:0px"><tr><td style="border:0px">Metric</td><td>Value</td></tr><tr><td>Bugs</td><td>7 | open</td></tr></table>
<ul><li><span data-tag="important">Risk: API quota</span></li></ul>
<img src="${R}/resources/r2/$value" data-src-type="image/jpeg" data-fullres-src="${R}/resources/r2/$value" data-fullres-src-type="image/jpeg" width="300" height="200" />
<object data="${R}/resources/f1/$value" data-attachment="report.pdf" type="application/pdf"></object>
</div></body></html>`;
}

function page2Html() {
  return `<html><head><title>Retro</title></head><body data-absolute-enabled="true">
<div data-id="_default" style="position:absolute;left:48px;top:120px;width:624px">
<p>Retro notes: the database migration was slow.</p>
<img src="https://evil.example.com/steal.png" data-fullres-src="https://evil.example.com/steal.png" />
</div></body></html>`;
}

const notebooks = [
  {
    id: "nb-1",
    displayName: "Work",
    isDefault: true,
    lastModifiedDateTime: "2026-10-12T10:00:00Z",
    sections: [{ id: "sec-1", displayName: "Sprint notes" }],
    sectionGroups: [
      { id: "sg-1", displayName: "Archive", sections: [{ id: "sec-2", displayName: "2025" }], sectionGroups: [] },
    ],
  },
];

const pages = [
  {
    id: "pg-1",
    title: "Sprint 42 status",
    createdDateTime: "2026-10-10T09:00:00Z",
    lastModifiedDateTime: "2026-10-12T09:00:00Z",
    links: { oneNoteWebUrl: { href: "https://onenote.example/pg-1" } },
    parentSection: { id: "sec-1", displayName: "Sprint notes" },
    parentNotebook: { id: "nb-1", displayName: "Work" },
  },
  {
    id: "pg-2",
    title: "Retro",
    createdDateTime: "2026-10-01T09:00:00Z",
    lastModifiedDateTime: "2026-10-02T09:00:00Z",
    parentSection: { id: "sec-1", displayName: "Sprint notes" },
    parentNotebook: { id: "nb-1", displayName: "Work" },
  },
];

before(async () => {
  tmp = await fs.mkdtemp(path.join(os.tmpdir(), "onenote-mcp-test-"));
  const big = await textPng(3000, 2000, "ERROR 504 at checkout service");
  const web = await sharp(big).resize(1500).png().toBuffer();
  const small = await sharp(await textPng(300, 200, "chart")).jpeg().toBuffer();

  server = http.createServer(async (req, res) => {
    const chunks = [];
    for await (const c of req) chunks.push(c);
    const body = Buffer.concat(chunks);
    const url = new URL(req.url, "http://x");
    const rel = decodeURIComponent(url.pathname.replace(/^\/v1\.0\/me\/onenote\//, ""));
    requests.push({ method: req.method, rel, query: url.search, headers: req.headers, body });

    const json = (status, obj) => {
      res.writeHead(status, { "content-type": "application/json" });
      res.end(JSON.stringify(obj));
    };
    if (req.headers.authorization !== `Bearer ${TOKEN}`) return json(401, { error: { code: "401", message: "bad token" } });

    const filter = url.searchParams.get("$filter") ?? "";
    const titleMatch = /contains\(tolower\(title\),'([^']*)'\)/.exec(filter);
    const filterPages = (list) => (titleMatch ? list.filter((p) => p.title.toLowerCase().includes(titleMatch[1])) : list);

    if (req.method === "GET" && rel === "notebooks") return json(200, { value: notebooks });
    if (req.method === "GET" && rel === "notebooks/nb-1") return json(200, notebooks[0]);
    if (req.method === "GET" && rel === "sections") return json(200, { value: [{ id: "sec-1", displayName: "Sprint notes", parentNotebook: { id: "nb-1", displayName: "Work" } }] });
    if (req.method === "GET" && (rel === "sections/sec-1/pages" || rel === "pages")) return json(200, { value: filterPages(pages) });
    if (req.method === "GET" && rel === "sections/sec-2/pages") return json(200, { value: [] });
    const pageMeta = /^pages\/([^/]+)$/.exec(rel);
    if (req.method === "GET" && pageMeta) {
      const p = pages.find((x) => x.id === pageMeta[1]);
      return p ? json(200, p) : json(404, { error: { code: "20102", message: "not found" } });
    }
    const content = /^pages\/([^/]+)\/content$/.exec(rel);
    if (content && req.method === "GET") {
      res.writeHead(200, { "content-type": "text/html" });
      return res.end(content[1] === "pg-1" ? page1Html() : page2Html());
    }
    if (content && req.method === "PATCH") {
      res.writeHead(204);
      return res.end();
    }
    if (req.method === "POST" && rel === "sections/sec-1/pages") {
      return json(201, { id: "pg-new", title: "created", links: { oneNoteWebUrl: { href: "https://onenote.example/pg-new" } } });
    }
    if (req.method === "POST" && rel === "notebooks/nb-1/sections") {
      return json(201, { id: "sec-new", displayName: JSON.parse(body.toString()).displayName });
    }
    const resource = /^resources\/([^/]+)\/\$value$/.exec(rel);
    if (req.method === "GET" && resource) {
      const id = resource[1];
      if (id === "r2" && r2Hits++ === 0) {
        res.writeHead(429, { "retry-after": "1" });
        return res.end();
      }
      const map = { r1: [big, "image/png"], "r1-web": [web, "image/png"], r2: [small, "image/jpeg"], f1: [Buffer.from("%PDF-1.4 fake"), "application/pdf"] };
      const hit = map[id];
      if (!hit) return json(404, { error: { code: "20102", message: "no resource" } });
      res.writeHead(200, { "content-type": hit[1] });
      return res.end(hit[0]);
    }
    json(404, { error: { code: "404", message: `no route ${req.method} ${rel}` } });
  });
  await new Promise((r) => server.listen(0, "127.0.0.1", r));
  port = server.address().port;

  client = new Client({ name: "e2e", version: "1.0.0" });
  await client.connect(
    new StdioClientTransport({
      command: process.execPath,
      args: [serverEntry],
      env: {
        PATH: process.env.PATH,
        HOME: tmp,
        ONENOTE_ACCESS_TOKEN: TOKEN,
        ONENOTE_GRAPH_ROOT: `http://127.0.0.1:${port}/v1.0`,
        ONENOTE_DOWNLOAD_DIR: path.join(tmp, "downloads"),
        ONENOTE_CACHE_DIR: path.join(tmp, "cache"),
      },
      stderr: "pipe",
    }),
  );
});

after(async () => {
  await client?.close();
  server?.close();
  if (tmp) await fs.rm(tmp, { recursive: true, force: true });
});

const textOf = (r) => r.content.filter((c) => c.type === "text").map((c) => c.text).join("\n");
const imagesOf = (r) => r.content.filter((c) => c.type === "image");

test("lists tools and prompts", async () => {
  const { tools } = await client.listTools();
  const names = tools.map((t) => t.name).sort();
  assert.deepEqual(names, [
    "onenote_collect_pages",
    "onenote_create_notebook",
    "onenote_create_page",
    "onenote_create_section",
    "onenote_download_resources",
    "onenote_get_page",
    "onenote_get_page_image",
    "onenote_list_notebooks",
    "onenote_list_pages",
    "onenote_list_sections",
    "onenote_search_pages",
    "onenote_sign_in",
    "onenote_sign_out",
    "onenote_update_page",
  ]);
  const { prompts } = await client.listPrompts();
  assert.deepEqual(prompts.map((p) => p.name).sort(), ["onenote_extract_screenshots", "onenote_report"]);
  const prompt = await client.getPrompt({ name: "onenote_report", arguments: { scope: "Work / Sprint notes", period: "last 7 days" } });
  assert.match(prompt.messages[0].content.text, /onenote_collect_pages/);
});

test("lists notebooks with nested sections", async () => {
  const r = await client.callTool({ name: "onenote_list_notebooks", arguments: {} });
  const t = textOf(r);
  assert.match(t, /Notebook: Work \(default\)/);
  assert.match(t, /Section: Sprint notes \| id: sec-1/);
  assert.match(t, /Section group: Archive/);
  assert.match(t, /Section: 2025 \| id: sec-2/);
});

test("lists pages with an OData title filter", async () => {
  const r = await client.callTool({ name: "onenote_list_pages", arguments: { section_id: "sec-1", title_contains: "Sprint" } });
  assert.match(textOf(r), /Sprint 42 status/);
  assert.doesNotMatch(textOf(r), /Retro/);
  const req = requests.findLast((x) => x.rel === "sections/sec-1/pages");
  assert.match(decodeURIComponent(req.query), /\$filter=contains\(tolower\(title\),'sprint'\)/);
});

test("reads a page as markdown in visual order, with screenshots as images", async () => {
  const r = await client.callTool({ name: "onenote_get_page", arguments: { page_id: "pg-1" } });
  assert.ok(!r.isError, textOf(r));
  const t = textOf(r);
  assert.ok(t.indexOf("Summary") < t.indexOf("Second outline"), "top outline should come first");
  assert.match(t, /Velocity was \*\*42 points\*\* this sprint/);
  assert.match(t, /\[ \] Follow up with QA/);
  assert.match(t, /\[x\] Ship build 7\.1/);
  assert.match(t, /\| Metric \| Value \|\n\| --- \| --- \|\n\| Bugs \| 7 \\\| open \|/);
  assert.match(t, /\[important\] Risk: API quota/);
  assert.match(t, /\[Image 1 \(300x200\)\]/);
  assert.match(t, /\[Image 2: Dashboard screenshot \(1500x1000\)\]/);
  assert.match(t, /\[Attachment 1: report\.pdf \(application\/pdf\)\]/);
  assert.match(t, /Open in OneNote: https:\/\/onenote\.example\/pg-1/);

  const imgs = imagesOf(r);
  assert.equal(imgs.length, 2);
  for (const img of imgs) {
    const meta = await sharp(Buffer.from(img.data, "base64")).metadata();
    assert.ok(Math.max(meta.width, meta.height) <= 1024, `image within 1024px, got ${meta.width}x${meta.height}`);
  }
  const big = await sharp(Buffer.from(imgs[1].data, "base64")).metadata();
  assert.equal(big.width, 1024);
  assert.match(t, /original 3000x2000/);
  assert.equal(r2Hits >= 2, true, "429 on the first resource request should be retried");
});

test("zooms into a screenshot region at high resolution", async () => {
  const r = await client.callTool({
    name: "onenote_get_page_image",
    arguments: { page_id: "pg-1", image_index: 2, region: { x: 0, y: 0.4, width: 0.5, height: 0.25 } },
  });
  assert.ok(!r.isError, textOf(r));
  const [img] = imagesOf(r);
  const meta = await sharp(Buffer.from(img.data, "base64")).metadata();
  assert.equal(meta.width, 1500);
  assert.equal(meta.height, 500);
  assert.match(textOf(r), /Original 3000x2000/);
});

test("never sends the token to non-Graph hosts", async () => {
  const r = await client.callTool({ name: "onenote_get_page_image", arguments: { page_id: "pg-2", image_index: 1 } });
  assert.equal(r.isError, true);
  assert.match(textOf(r), /Refusing to send credentials to evil\.example\.com/);
});

test("searches page text", async () => {
  const r = await client.callTool({ name: "onenote_search_pages", arguments: { query: "database migration", section_id: "sec-1" } });
  const t = textOf(r);
  assert.match(t, /Retro/);
  assert.match(t, /the database migration was slow/);
  assert.doesNotMatch(t, /Sprint 42 status/);
});

test("collects pages oldest to newest for reports", async () => {
  const r = await client.callTool({ name: "onenote_collect_pages", arguments: { section_id: "sec-1", include_images: true, max_images_total: 1 } });
  const t = textOf(r);
  assert.ok(t.indexOf("## Retro") < t.indexOf("## Sprint 42 status"));
  assert.equal(imagesOf(r).length, 1);
});

function parseMultipart(req) {
  const boundary = /boundary=(.+)$/.exec(req.headers["content-type"])[1];
  return req.body
    .toString("latin1")
    .split(`--${boundary}`)
    .slice(1, -1)
    .map((chunk) => {
      const [head, ...rest] = chunk.split("\r\n\r\n");
      return {
        name: /name="([^"]+)"/.exec(head)[1],
        type: /Content-Type: ([^\r\n]+)/i.exec(head)[1],
        body: rest.join("\r\n\r\n").replace(/\r\n$/, ""),
      };
    });
}

test("creates a report page with tables, to-dos, local and copied screenshots", async () => {
  const shot = path.join(tmp, "local shot.png");
  await fs.writeFile(shot, await textPng(2000, 1000, "local screenshot"));
  const md = [
    "## Highlights",
    "",
    "| Metric | Value |",
    "| --- | ---: |",
    "| Bugs | 7 |",
    "",
    "- [ ] Follow up with QA",
    "- [x] Ship build",
    "",
    "```",
    "  indented code",
    "```",
    "",
    `![Local](<${shot}>)`,
    "",
    "![From OneNote](onenote-image://pg-1/2)",
  ].join("\n");
  const r = await client.callTool({
    name: "onenote_create_page",
    arguments: { title: "Weekly report <draft> été", content: md, notebook_name: "work", section_name: "Sprint notes" },
  });
  assert.ok(!r.isError, textOf(r));
  assert.match(textOf(r), /Page id: pg-new/);

  const req = requests.findLast((x) => x.method === "POST" && x.rel === "sections/sec-1/pages");
  const parts = parseMultipart(req);
  assert.deepEqual(parts.map((p) => p.name), ["Presentation", "img1", "img2"]);
  const html = parts[0].body;
  assert.match(html, /<title>Weekly report &lt;draft&gt; &#xe9;t&#xe9;<\/title>/);
  assert.ok(/^[\x00-\x7f]*$/.test(html), "presentation HTML should be ASCII-only");
  assert.match(html, /<table border="1">/);
  assert.match(html, /<td><b>Metric<\/b><\/td>/);
  assert.match(html, /style="text-align:right"/);
  assert.doesNotMatch(html, /<th|<thead|<tbody|<input/);
  assert.match(html, /<li data-tag="to-do">/);
  assert.match(html, /<li data-tag="to-do:completed">/);
  assert.match(html, /font-family:Consolas.*&#xa0;&#xa0;indented code/);
  assert.match(html, /\n<img src="name:img1" alt="Local" width="960" height="480" \/>/);
  assert.match(html, /\n<img src="name:img2" alt="From OneNote" width="960" height="640" \/>/);
  assert.match(html, /<li data-tag="to-do">Follow up with QA<\/li>/);
  assert.equal(parts[1].type, "image/png");
  assert.equal(parts[2].type, "image/png");
});

test("creates a missing section when asked", async () => {
  const r = await client.callTool({
    name: "onenote_create_page",
    arguments: { title: "x", content: "hi", notebook_name: "Work", section_name: "Reports", create_section_if_missing: true },
  });
  // The mock only accepts page POSTs to sec-1, so the page step fails; the section must have been created first.
  assert.ok(requests.some((x) => x.method === "POST" && x.rel === "notebooks/nb-1/sections"));
  assert.match(textOf(r), /no route POST sections\/sec-new\/pages/);
});

test("appends to a page and rejects web images on update", async () => {
  const ok = await client.callTool({ name: "onenote_update_page", arguments: { page_id: "pg-1", content: "**Update:** fixed" } });
  assert.ok(!ok.isError, textOf(ok));
  const req = requests.findLast((x) => x.method === "PATCH");
  const cmds = JSON.parse(req.body.toString());
  assert.deepEqual(cmds, [{ target: "body", action: "append", content: "<p><strong>Update:</strong> fixed</p>\n" }]);

  const bad = await client.callTool({
    name: "onenote_update_page",
    arguments: { page_id: "pg-1", content: "![x](https://example.com/a.png)" },
  });
  assert.equal(bad.isError, true);
  assert.match(textOf(bad), /does not download web images/);

  const title = await client.callTool({ name: "onenote_update_page", arguments: { page_id: "pg-1", target: "title", action: "replace", content: "New name" } });
  assert.match(textOf(title), /title changed/);
});

test("downloads full-resolution images and attachments locally", async () => {
  const r = await client.callTool({ name: "onenote_download_resources", arguments: { page_id: "pg-1", all_images: true, all_attachments: true } });
  assert.ok(!r.isError, textOf(r));
  const files = [...textOf(r).matchAll(/^- (.+)$/gm)].map((m) => m[1]);
  assert.equal(files.length, 3);
  for (const f of files) assert.ok(f.startsWith(path.join(tmp, "downloads")), f);
  const meta = await sharp(await fs.readFile(files.find((f) => f.endsWith("image-2.png")))).metadata();
  assert.equal(meta.width, 3000);
  assert.ok(files.some((f) => f.endsWith("report.pdf")));

  const escape = await client.callTool({ name: "onenote_download_resources", arguments: { page_id: "pg-1", all_images: true, subfolder: "../../etc" } });
  const escaped = [...textOf(escape).matchAll(/^- (.+)$/gm)].map((m) => m[1]);
  for (const f of escaped) assert.ok(f.startsWith(path.join(tmp, "downloads") + path.sep), f);
});

test("explains setup when no client ID is configured", async () => {
  const c = new Client({ name: "e2e-noauth", version: "1.0.0" });
  await c.connect(
    new StdioClientTransport({
      command: process.execPath,
      args: [serverEntry],
      env: { PATH: process.env.PATH, HOME: tmp, ONENOTE_CACHE_DIR: path.join(tmp, "cache2") },
      stderr: "pipe",
    }),
  );
  const r = await c.callTool({ name: "onenote_list_notebooks", arguments: {} });
  assert.equal(r.isError, true);
  assert.match(textOf(r), /not configured yet/);
  assert.match(textOf(r), /onenote-mcp configure --client-id/);
  await c.close();
});

test("configure saves settings that the server and CLI use, with env overrides", async () => {
  const { execFileSync } = await import("node:child_process");
  const cacheDir = path.join(tmp, "cfg");
  const env = { PATH: process.env.PATH, HOME: tmp, ONENOTE_CACHE_DIR: cacheDir };
  const run = (args, extra = {}) => execFileSync(process.execPath, [serverEntry, ...args], { env: { ...env, ...extra }, encoding: "utf8" });

  run(["configure", "--client-id", "aaaaaaaa-0000-0000-0000-000000000001", "--tenant", "contoso.com", "--auth=device"]);
  const file = path.join(cacheDir, "config.json");
  const saved = JSON.parse(await fs.readFile(file, "utf8"));
  assert.deepEqual(saved, { clientId: "aaaaaaaa-0000-0000-0000-000000000001", tenantId: "contoso.com", authFlow: "device" });
  assert.equal((await fs.stat(file)).mode & 0o777, 0o600);

  run(["configure", "--scopes", "Notes.ReadWrite.All"]);
  const status = run(["status"]);
  assert.match(status, /Client ID: +aaaaaaaa-0000-0000-0000-000000000001/);
  assert.match(status, /Tenant: +contoso\.com/);
  assert.match(status, /Sign-in: +device/);
  assert.match(status, /Scopes: +Notes\.ReadWrite\.All/);

  const overridden = run(["status"], { ONENOTE_TENANT_ID: "fabrikam.com", ONENOTE_AUTH_FLOW: "browser" });
  assert.match(overridden, /Tenant: +fabrikam\.com/);
  assert.match(overridden, /Sign-in: +browser/);

  assert.throws(() => run(["configure", "--auth", "carrier-pigeon"]), /--auth must be/);
});
