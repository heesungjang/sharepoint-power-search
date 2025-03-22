# Power Search for SharePoint

![SharePoint Framework](https://img.shields.io/badge/SPFx-1.20.0-green.svg)
![FluentUI](https://img.shields.io/badge/FluentUI-9.56.1-blue.svg)
![License](https://img.shields.io/github/license/mashape/apistatus.svg)

A modern, keyboard-accessible search experience for SharePoint. This extension provides a clean, minimal search interface with powerful filtering capabilities.

## Features

- **Keyboard-first approach**: Open search with `Ctrl+K` or `Cmd+K` (Mac)
- **Filtering options**:
  - Filter by file type (docx, pptx, xlsx, pdf, pages, images)
  - Sort by relevance, newest first, or oldest first
- **View SharePoint pages in-place** without leaving the search interface
- **Recent searches history** for quick access to previous queries
- **Spotlight-like categories** to organize search results
- **Built with modern SPFx and FluentUI v9**

## Architecture

This extension is built as a SharePoint Framework (SPFx) Application Customizer that integrates with the SharePoint Search API. Key technologies used:

- SharePoint Framework (SPFx) 1.20.0
- React 17
- FluentUI v9 (React Components)
- PnPjs v4
- React Query for data fetching and caching

## Prerequisites

- SharePoint Online tenant
- Node.js v18.17.1 (LTS)
- SharePoint Framework compatible development environment

## Installation

### From Source

1. Clone this repository

```bash
git clone https://github.com/your-username/power-search.git
cd power-search
```

2. Install dependencies

```bash
npm install
```

3. Build and test the solution locally

```bash
npm run serve
```

4. Build the solution

```bash
npm run build
```

5. Package the solution file

```bash
npm run package
```

5. Deploy the solution to your SharePoint App Catalog
   - Upload the `.sppkg` file from the `sharepoint/solution` folder to your App Catalog
   - Deploy the solution globally or to specific sites

### From App Catalog

1. Navigate to your SharePoint App Catalog
2. Upload the `power-search.sppkg` file
3. Deploy the solution globally or to specific sites

## Configuration

No special configuration is required. Once deployed, the Power Search extension will automatically be available on all pages within the site(s) where it's activated.

## Usage

1. Press `Ctrl+K` (or `Cmd+K` on Mac) from any SharePoint page
2. Type your search query
3. Use filter buttons to narrow results by file type
4. Choose sorting methods to organize results
5. Click on a result to view it within the search interface
6. Open the document in a new tab using the external link button

## Development

This project uses SPFx fast-serve for development:

```bash
npm run serve
```

To build the solution:

```bash
npm run build
```
