# OneNote MCP server

An MCP server that connects Claude to Microsoft OneNote through Microsoft Graph. Claude can
browse notebooks, read pages, **see the screenshots and images on them**, zoom into small text,
gather many pages for a report, and write reports back to OneNote.

It runs locally on your computer and works with personal Microsoft accounts and with company
(Microsoft 365 work or school) accounts.

## What it can do

- **Browse:** list notebooks, section groups, sections and pages; filter pages by title and date.
- **Search:** find pages by title (server-side) and by page text (scans recent pages in a scope).
- **Read pages:** page text comes back as clean markdown in visual reading order. Tables, to-do
  checkboxes, note tags, links and attachments are kept, and each image is marked `[Image N]`.
- **Read screenshots:** images are returned as real images that Claude can look at. Large
  screenshots are downscaled for the page view. A zoom tool returns one image at high resolution or
  crops a region, so small text in a full-screen capture stays legible.
- **Report:** collect every page in a section for a date range (text plus optional images), oldest
  to newest, ready to summarise.
- **Write:** create pages from markdown, append to or edit existing pages, create sections and
  notebooks. Reports can embed local screenshots or copy images from other OneNote pages.
- **Export:** save original-resolution screenshots and attachments to a local folder, for example to
  put them in a Word or PDF report.

## Quick start

1. [Register an app in Microsoft Entra](#1-register-an-app-in-microsoft-entra) and copy its
   client ID and tenant ID.
2. Install:

   ```bash
   git clone https://github.com/joshmaster2165/onenote-mcp.git
   ```

   ```bash
   cd onenote-mcp && npm install && npm run build
   ```

3. Save your app details and sign in:

   ```bash
   node dist/index.js configure --client-id <application-client-id> --tenant <tenant-id-or-domain>
   ```

   ```bash
   node dist/index.js login
   ```

4. [Add the server to Claude](#3-add-it-to-claude).

Requires Node.js 18.17 or newer.

## Setup in detail

### 1. Register an app in Microsoft Entra

OneNote access goes through Microsoft Graph, which needs an app registration. It takes about five
minutes. In many companies only IT can do this; see [Company accounts](#company-work-or-school-accounts)
for a request you can send them.

1. Go to <https://entra.microsoft.com>, then **Identity → Applications → App registrations →
   New registration**.
2. Name it, for example `OneNote MCP`.
3. Under **Supported account types**:
   - For a company account, choose **Accounts in this organizational directory only**.
   - For a personal Outlook/Hotmail account, choose **Accounts in any organizational directory
     and personal Microsoft accounts**.
4. Under **Redirect URI**, pick **Public client/native (mobile & desktop)** and enter
   `http://localhost`. Click **Register**.
5. Open **Authentication**. Under **Advanced settings**, set **Allow public client flows** to
   **Yes** and save. This is only needed for device-code sign-in, but it does no harm.
6. Open **API permissions → Add a permission → Microsoft Graph → Delegated permissions** and add
   **Notes.ReadWrite.All**. Use **Notes.ReadWrite** instead if you only need your own notebooks.
   If your company requires it, click **Grant admin consent** (admins only).
7. From **Overview**, copy the **Application (client) ID** and the **Directory (tenant) ID**.

### 2. Configure and sign in

```bash
node dist/index.js configure --client-id <application-client-id> --tenant <directory-tenant-id> --scopes Notes.ReadWrite.All
```

This writes `~/.onenote-mcp/config.json`, which is readable only by you. You can use your company
domain, such as `contoso.com`, instead of the tenant ID. Personal accounts use `--tenant consumers`.

```bash
node dist/index.js login
```

A browser window opens. Sign in with your work account and approve the permissions. The token is
cached in `~/.onenote-mcp/msal-cache.json` and refreshes automatically. You can skip this step: the
first OneNote request in Claude opens the same browser sign-in.

If your computer can't open a browser, use device-code sign-in instead. It shows a code to enter at
<https://microsoft.com/devicelogin> on any device:

```bash
node dist/index.js login --device
```

Check your setup at any time:

```bash
node dist/index.js status
```

### 3. Add it to Claude

Use the absolute path to `dist/index.js` in your clone. Because the app details are in the settings
file, the MCP config needs no secrets or environment variables.

**Claude Code and the Claude desktop app's Code tab:**

```bash
claude mcp add onenote --scope user -- node /absolute/path/to/onenote-mcp/dist/index.js
```

**Claude desktop app chat:** edit `~/Library/Application Support/Claude/claude_desktop_config.json`
(on Windows, `%APPDATA%\Claude\claude_desktop_config.json`), then quit and reopen Claude:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "/absolute/path/to/node",
      "args": ["/absolute/path/to/onenote-mcp/dist/index.js"]
    }
  }
}
```

Use the full path from `which node` for `command`. The desktop app doesn't always inherit your
shell's `PATH`.

## Company (work or school) accounts

### What to ask IT for

If you can't register apps yourself, send your IT or Microsoft 365 admin something like this:

> Please create an app registration for a local OneNote tool I use with Claude:
> - Supported account types: this organization only
> - Redirect URI: Public client/native (mobile & desktop), `http://localhost`
> - Allow public client flows: Yes
> - Microsoft Graph delegated permission: `Notes.ReadWrite.All` (or `Notes.ReadWrite`), with admin
>   consent if our tenant requires it
> - No client secret or certificate is needed. It signs in as me and can reach only what I can.
>
> Please send me the Application (client) ID and Directory (tenant) ID.

### Notebooks in Teams or SharePoint

By default the server reads notebooks in your own OneDrive (`me/onenote`). Team notebooks live in
a SharePoint site or a Microsoft 365 group, so you point the server at that location:

```bash
node dist/index.js configure --base-path "sites/<site-id>/onenote"
```

To find the IDs, sign in to [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
with your work account and run one of these:

- `GET https://graph.microsoft.com/v1.0/me/joinedTeams` lists your Teams. Each `id` is a group ID,
  used as `groups/<group-id>/onenote`.
- `GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Engineering` returns the
  site's `id`, used as `sites/<site-id>/onenote`. Change the host and path to match the notebook's
  SharePoint address.

To use your own notebooks and a team notebook together, add the server to Claude twice with
different names. The second one sets its location with an environment variable:

```bash
claude mcp add onenote-team --scope user -e ONENOTE_BASE_PATH=sites/<site-id>/onenote -- node /absolute/path/to/onenote-mcp/dist/index.js
```

Both share one sign-in. Use the `Notes.ReadWrite.All` permission for notebooks you don't own.

### Troubleshooting sign-in

Errors show Microsoft's `AADSTS` code and a suggested fix. The common ones:

| Error | Meaning and fix |
| --- | --- |
| AADSTS700016, AADSTS50059 | Wrong client ID or tenant. Check both with `node dist/index.js status`. |
| AADSTS50011, AADSTS500113 | Redirect URI missing. Add `http://localhost` under Mobile and desktop applications. |
| AADSTS7000218 | Public client flows are off. Turn on "Allow public client flows". |
| AADSTS53003 | Blocked by a Conditional Access policy. Use browser sign-in (the default) on a company-managed device, or ask IT. |
| AADSTS65001, AADSTS90094 | Admin consent required. Ask IT to grant consent for the Notes permission. |
| AADSTS50105 | You aren't assigned to the app. Ask IT to assign you. |
| 403 from Graph | The permission is missing or not consented, or the notebook isn't shared with you. |

## Things to ask Claude

- "List my OneNote notebooks."
- "Read the page *Sprint 42 status* and tell me what the dashboard screenshot shows."
- "Zoom into the error log in image 2 on that page and transcribe the error codes."
- "Summarise everything in *Work / Meeting notes* from the last 7 days as a status report."
- "Write that report to a new page in *Work / Reports*, with the latency chart screenshot embedded."
- "Append today's action items to the bottom of that page."
- "Save all screenshots from that page to my Downloads folder."

The server also provides two ready-made prompts. **onenote_report** builds a report from a
notebook or section. **onenote_extract_screenshots** transcribes the text and tables in a page's
screenshots.

## Tools

| Tool | What it does |
| --- | --- |
| `onenote_sign_in` / `onenote_sign_out` | Start sign-in, or clear the cached sign-in |
| `onenote_list_notebooks` | Notebooks with their section groups and sections |
| `onenote_list_sections` | Sections in a notebook or section group, filterable by name |
| `onenote_list_pages` | Pages in a section, a notebook or everywhere; filter by title and modified date |
| `onenote_search_pages` | Title search plus a text scan of recent pages |
| `onenote_get_page` | Page as markdown, with images inline; `format=html_with_ids` returns element IDs for edits |
| `onenote_get_page_image` | One image at high resolution, optionally cropped to a region |
| `onenote_collect_pages` | Many pages at once for reports, oldest to newest |
| `onenote_create_page` | New page from markdown or HTML, by section ID or by notebook and section name |
| `onenote_update_page` | Append, prepend, insert, replace an element, or rename the page |
| `onenote_create_section` / `onenote_create_notebook` | Create containers for reports |
| `onenote_download_resources` | Save original images and attachments locally |

There is no delete tool, on purpose.

### Writing content

`content` is markdown by default. Headings, bold, italic, strikethrough, links, bullet and numbered
lists, task lists (these become OneNote to-do checkboxes), tables, code blocks and quotes are
converted to HTML that OneNote accepts. Images can come from:

- a local file: `![Latency chart](/Users/me/Desktop/latency.png)`. Wrap paths that contain spaces
  in angle brackets: `![x](</Users/me/My Shots/a.png>)`
- another OneNote page: `![Dashboard](onenote-image://PAGE_ID/2)` copies image 2 from that page
- the web: `![Logo](https://example.com/logo.png)`, for new pages only

Images wider than 960 px are shown at 960 px on the page, and the full-resolution file is kept.

## Configuration

Settings come from `~/.onenote-mcp/config.json` (written by `configure`). Environment variables
override it. See [.env.example](.env.example) for the full list.

| Setting | `configure` option | Environment variable | Default |
| --- | --- | --- | --- |
| App (client) ID | `--client-id` | `ONENOTE_CLIENT_ID` | (required) |
| Tenant | `--tenant` | `ONENOTE_TENANT_ID` | `common` |
| Sign-in method | `--auth browser\|device` | `ONENOTE_AUTH_FLOW` | `browser` |
| Permissions | `--scopes` | `ONENOTE_SCOPES` | `Notes.ReadWrite` |
| OneNote location | `--base-path` | `ONENOTE_BASE_PATH` | `me/onenote` |
| Download folder | `--download-dir` | `ONENOTE_DOWNLOAD_DIR` | `~/Downloads/OneNote MCP` |

Other environment variables: `ONENOTE_CACHE_DIR` (settings and token location),
`ONENOTE_NO_BROWSER=1` (print the sign-in link instead of opening a browser),
`ONENOTE_MAX_RESPONSE_IMAGE_BYTES` (default 700000) and `ONENOTE_MAX_IMAGES` (default 6).

**Response size.** Image-heavy responses can hit client limits. Claude Code caps MCP output with
`MAX_MCP_OUTPUT_TOKENS` (default 25,000). Raise that, or lower `max_images` or `image_max_px` on the
call. The server keeps each response's images within `ONENOTE_MAX_RESPONSE_IMAGE_BYTES` and says
which images it skipped, so Claude can fetch them one at a time.

## Security and privacy

- Uses delegated sign-in only, so Claude acts as you and can reach only what your account can.
  Microsoft no longer supports app-only access to the OneNote API.
- No client secret is used or stored. Sign-in uses the standard authorization-code flow with PKCE,
  or the device-code flow.
- The access token is sent only to `graph.microsoft.com`. Image links on a page that point
  elsewhere are refused, not fetched with your token.
- The settings and token cache files are created with owner-only permissions.
- Local files can be uploaded only as images, and only after they decode as a real image.
- Downloads are written only inside the download folder.
- Page content is labelled as data, and the server tells Claude not to follow instructions found
  inside notes.

## Limitations

- Microsoft Graph has no full-text OneNote search. Text search scans the most recently modified
  pages in the scope you give (60 by default), so narrow it to a section or notebook for
  thoroughness.
- Ink drawings and handwriting aren't returned by the Graph API. OneNote's own image OCR text isn't
  exposed either, so Claude reads screenshots visually instead.
- Some image formats pasted from Office, such as EMF and WMF, can't be decoded. They're listed but
  not shown.
- Each server instance reads one OneNote location. Add a second instance for a team notebook.
- Accounts with very large OneNote libraries can get error 10008 on cross-notebook queries. Pass a
  section or notebook to avoid it.

## Development

```bash
npm test
```

This builds the server and runs an end-to-end suite. The suite starts the real server over stdio,
points it at a local mock of Microsoft Graph, and checks reading, image handling, search, report
collection, page creation with uploaded images, updates, downloads, settings and the credential
host check. It needs no Microsoft account or network access.

Project layout:

- `src/index.ts`: entry point and CLI (`configure`, `login`, `status`, `logout`)
- `src/config.ts`: settings file and environment variables
- `src/auth.ts`: MSAL browser and device-code sign-in, and the token cache
- `src/graph.ts`: Graph HTTP client with retries, paging and the host allow-list
- `src/onenote.ts`: OneNote operations and caching
- `src/html.ts`: conversion from page HTML to markdown, and from markdown to OneNote HTML
- `src/images.ts`: resizing, cropping and compression for vision, plus upload validation
- `src/content.ts`: resolves image sources in outgoing content into upload parts
- `src/tools.ts` and `src/prompts.ts`: MCP tool and prompt definitions

## License

MIT
