# M365 FAQ Code App - Beginner Setup and Deployment Guide

This guide shows how to create, build, and publish the M365 FAQ Code App in Power Apps. It is written like a step-by-step beginner walkthrough.

## App Features

This app is packed with modern UI/UX enhancements out of the box:

- **Clean Header:** Minimal text-only header (no logo) with the app title and an always-visible search bar.
- **Dynamic Layout:** Left sidebar for main categories and top tabs for subcategories.
- **Accordion & Modal UI:** Configure answers to expand smoothly inline (accordion) or open in a focused modal overlay — controlled by a single config property.
- **Pinned Items:** Pinned FAQs appear first with a ★ star indicator.
- **Markdown Support:** Answers fully support markdown rendering (bold, italics, lists, links, tables, inline code, etc.).
- **SharePoint Attachments:** Detects items with attachments and renders clickable file links directly in the answer.
- **Copy to Clipboard:** A one-click "Copy" button formats the question, answer, and attachments for easy sharing over Teams or email.
- **Search Highlighting:** Live visual highlighting of matching search terms within FAQ titles.
- **Pagination:** Configurable per-page limit handles large datasets gracefully; pagination controls reset automatically on filter changes.
- **Mobile Responsive:** A slide-out hamburger sidebar on screens under 980px wide.
- **Beautiful Empty States:** Illustrated empty state with a "Clear Search" button when a search yields no results.

## Prerequisites

- Windows or macOS machine
- Node.js installed (version 22 or later is recommended - earlier versions may have compatibility issues)
- Power Platform CLI (`pac`) installed (version 2.7.4 or later)
- Access to a Microsoft 365 Power Apps environment
- Access to a SharePoint site with a list or data source
- A Microsoft account with permission to create apps and use Power Apps
- SharePoint permissions: at least **Contribute** access on the target list and **Site Member** or higher on the site; if you need to create the list, **Site Owner** or **List Owner** is required

## Step 0: Create SharePoint List (Data Source)

Before creating the app, you need a SharePoint list to store the FAQ data.

### Option 1: Create SharePoint List Manually

1. Go to your SharePoint site: `https://vishtechtalk.sharepoint.com/sites/DemoSite`
2. Click **Site Contents** → **New** → **List**
3. Name it: `M365 FAQs`
4. Add these columns:

| Column Name | Type | Required | Notes |
|-------------|------|----------|-------|
| Title | Single line of text | Yes | The FAQ question |
| Answer | Multiple lines of text | Yes | The FAQ answer |
| Category | Choice | Yes | Dropdown: Account, Teams, Outlook, SharePoint, OneDrive, Other |
| SubCategory | Choice | No | Dropdown: Security, Setup, Usage, Troubleshooting |
| Active | Yes/No | Yes | Default: Yes. Only active items are shown in the app |
| Pinned | Yes/No | No | If Yes, the item appears first in its category/subcategory |

5. Add sample data:
   - Title: "How to reset password?"
   - Answer: "Go to account settings and click 'Reset Password'"
   - Category: Account
   - SubCategory: Security
   - Active: Yes

### Option 2: Import from CSV

1. Create a CSV file with this structure:

```csv
Title,Answer,Category,SubCategory,Active,Pinned
"How to reset password?","Go to account settings and click 'Reset Password'","Account","Security","Yes","No"
"How to create a team?","Click Teams app, then 'Create team'","Teams","Setup","Yes","No"
```

2. In SharePoint, create a new list and import the CSV:
   - **Site Contents** → **New** → **List**
   - Choose **From Excel**
   - Upload your CSV file
   - Map columns appropriately

> Note: The app only displays records where `Active` is set to `Yes`.
> Pinned items are shown first within their selected category/subcategory.

## Step 1: Create Power Apps Environment Connection

### 1.1 Create Shared Connection in Power Apps

1. Go to https://make.powerapps.com
2. Select your environment (e.g., "DemoEnv")
3. Click **Data** → **Connections** → **New connection**
4. Search for **SharePoint** and select it
5. Choose **Connect directly (cloud services)**
6. Sign in with your Microsoft account
7. Select your SharePoint site: `https://vishtechtalk.sharepoint.com/sites/DemoSite`
8. Click **Create**

### 1.2 Verify Connection

- Go to **Data** → **Connections**
- You should see "SharePoint" connection listed
- Test the connection by clicking it

## Step 2: Open the Project

Open the folder in VS Code or another code editor. In this repo, the main files are:

- `package.json`
- `vite.config.ts`
- `power.config.json`
- `src/App.tsx`
- `src/App.css`
- `src/fieldMap.ts`

### App UI Overview
- The app has a slim top header with the app title and a search bar.
- The left sidebar shows all FAQ categories. On mobile, the sidebar is hidden by default and revealed via a **hamburger menu** button in the header.
- The right content panel shows the selected category title, an `All` tab plus subcategory tabs, and the paginated FAQ cards.
- Only FAQs with `Active = Yes` are displayed.
- `Pinned` items appear first within their selected category/subcategory with a ★ star indicator.
- Each FAQ card expands inline (accordion) or opens a modal depending on the `uiMode` configuration.
- Answers support **Markdown** formatting (bold, lists, links, etc.).
- Items with SharePoint attachments display clickable file links below the answer.

### App Configuration (`src/config.ts`)
All global settings are stored in `src/config.ts`. Edit this file to control app behaviour:

```typescript
export const appConfig = {
  maxItemsPerPage: 5,       // Number of FAQs shown per page before pagination appears
  uiMode: "accordion",     // "accordion" = answers expand inline | "modal" = answers open in a popup
};
```

| Property | Type | Default | Description |
|---|---|---|---|
| `maxItemsPerPage` | `number` | `5` | Controls how many FAQs are shown per page. |
| `uiMode` | `"accordion"` \| `"modal"` | `"accordion"` | Switches between inline accordion view and the full-screen modal view. |

### SharePoint Field Mapping (`src/fieldMap.ts`)
- The app resolves SharePoint column names using the mapping in `src/fieldMap.ts`.
- If your SharePoint list uses a custom internal field name for any column (especially `SubCategory`), add that internal name to the corresponding array.
- Supported fields: `id`, `title`, `answer`, `category`, `subCategory`, `active`, `pinned`.

> **Note on Attachments:** The Power Apps SharePoint connector does not include attachment file URLs in the list response payload. The app detects the `{HasAttachments}` flag and generates a link to the SharePoint item's display page as a fallback. To show a direct download link for a specific known file, hardcode the attachment URL in `src/App.tsx` inside the `normalizeFaq` function.

## Step 3: Install Dependencies

In a terminal, run:

```bash
cd /path/to/m365-faq-code-app
npm install
```

## Step 4: Validate Node.js Version

Make sure Node.js is installed and the version is supported.

```bash
node -v
npm -v
```

If you have issues, uninstall and reinstall Node.js from the official site:

- macOS: https://nodejs.org/
- Windows: https://nodejs.org/

Use a stable LTS version like **18.x** or **20.x**.

## Step 5: Authenticate Power Apps CLI

Make sure `pac` is installed. Then log in to Power Apps:

```bash
pac auth create
```

Verify authentication and environments:

```bash
pac auth list
pac org list
```

This shows the environments available.

## Step 6: Choose the Correct Environment

You must use the environment where the SharePoint data source is available.

For example:

- `DemoEnv` with environment ID `21d09661-3b00-e734-a311-2aeaa306a481`
- `VibeDev` with environment ID `1441687e-e0a7-e4b3-a84d-747b95067a0c`

If `pac` is using the wrong profile, select the correct one using the index from `pac auth list`:

```bash
pac auth select --index 1
```

Then check again with:

```bash
pac org list
```

## Step 7: Update Power Apps Configuration

### 7.1 Update power.config.json

Edit `power.config.json` to match your environment and SharePoint list:

```json
{
  "version": "1.0",
  "appId": null,
  "appDisplayName": "M365 FAQ Code App",
  "region": "prod",
  "environmentId": "21d09661-3b00-e734-a311-2aeaa306a481",
  "description": "Microsoft 365 FAQ app",
  "buildPath": "./dist",
  "buildEntryPoint": "index.html",
  "localAppUrl": "http://localhost:3000",
  "logoPath": "Default",
  "connectionReferences": {
    "219d4846-2589-45b6-a72b-30ca2ad126aa": {
      "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "displayName": "SharePoint",
      "dataSources": [
        "m365_faqs"
      ],
      "dataSets": {
        "https://vishtechtalk.sharepoint.com/sites/DemoSite": {
          "dataSources": {
            "m365_faqs": {
              "tableName": "42f629c9-09bc-4de5-a7f6-d3b92f5f2265"
            }
          }
        }
      }
    }
  },
  "databaseReferences": {}
}
```

Replace:
- `environmentId`: Your environment ID from `pac org list`
- `https://vishtechtalk.sharepoint.com/sites/DemoSite`: Your SharePoint site URL
- `tableName`: Your SharePoint list ID (found in list settings)

### Connection reference details
- The top-level key under `connectionReferences` is the connection reference ID (a GUID generated by Power Apps).
- The `id` inside the connection reference object is the connector API path, for example `/providers/Microsoft.PowerApps/apis/shared_sharepointonline` for SharePoint.
- The `displayName` field is the connection name shown in Power Apps, for example `SharePoint`.

So the correct mapping is:
- `connectionReferences` key = connection reference GUID
- `displayName` = connection name in Power Apps UI
- `id` = connector API reference

If you update this file manually, you should typically keep the existing GUID key and only change the `displayName` or `tableName` if needed.

### 7.2 Find SharePoint List ID

1. Go to your SharePoint list
2. Click **List settings** → **List information**
3. Copy the List ID (looks like: `42f629c9-09bc-4de5-a7f6-d3b92f5f2265`)

## Step 8: Build the App Locally

Before publishing, build the project locally:

```bash
npm run build
```

If the build succeeds, the `dist/` folder is generated.

## Step 9: Initialize the Code App

The project should already have Power Apps code app configuration in `power.config.json`.

If initialization is needed, use:

```bash
pac code init --displayName "M365 FAQ Code App" --description "Microsoft 365 FAQ app" --buildPath dist --fileEntryPoint index.html --appUrl http://localhost:3000 --logoPath Default --region prod
```

If the config already exists, you do not need to run this command again.

## Step 10: Publish the App to Power Apps

Once the app is built and the correct environment is selected, publish with:

```bash
pac code push --solutionName "M365 FAQ Code App"
```

That pushes the current `dist/` build into the active Power Apps environment.

## Step 11: Verify the App

Open the app URL returned by the `pac code push` command:

```text
https://apps.powerapps.com/play/e/<environment-id>/app/<app-id>
```

The FAQ app should load with the updated layout.

## Common Issues and Fixes

### 1. `pac pcf push` error

If you run `pac pcf push`, you will get:

```text
Error: Power Apps component framework project file with extension pcfproj was not found
```

That error occurs because this is a Code app, not a PCF project. The correct command is:

```bash
pac code push --solutionName "M365 FAQ Code App"
```

### 2. Wrong environment selected

If the app is pushed to the wrong environment, use `pac auth list` and `pac org list` to confirm the active environment, then switch using:

```bash
pac auth select --index <correct-index>
```

### 3. Node.js issues

If Node.js is broken or incompatible, uninstall it and reinstall the latest stable LTS version.

### 4. Code app feature not enabled

If code app options are missing in Power Apps, ensure your environment supports code apps and that the code app feature is enabled. This may require admin permissions or a developer environment.

### 5. SharePoint connection issues

- Verify the SharePoint site URL is correct
- Ensure you have permissions to access the site and list
- Check that the list ID in `power.config.json` matches your SharePoint list

### 6. Data not loading

- Check browser console for errors (F12)
- Verify the SharePoint list has data and the column names match the code
- Ensure the connection reference in `power.config.json` is correct



## Quick Commands

```bash
cd /path/to/m365-faq-code-app
npm install
npm run build
pac auth list
pac org list
pac auth select --index 1
pac code push --solutionName "M365 FAQ Code App"
```

## Notes for Beginners

- Always build the app first before publishing.
- Use the correct Power Apps environment.
- Do not confuse `pac pcf` with `pac code`.
- If you see errors, read the messages carefully; they usually indicate the wrong command or wrong environment.

## Final Result

This app is a Power Apps Code app connected to a SharePoint FAQ data source. It is now published in the selected environment with a polished modern UI.
