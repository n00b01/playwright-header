# Standalone Header Menu Link Validator

This folder is intentionally separate from the Enterprise Audit Dashboard app.
It contains a standalone validator that checks header/menu links on target pages and exports Excel and HTML reports.

## Purpose
- load a page
- collect only header/menu links
- resolve relative URLs
- test HTTP response status
- generate XLSX + HTML output
- keep all logic separate from the existing dashboard app

## Run

```powershell
cd 'D:\TestAppv2\standalone-validator'
$env:HEADER_MENU_LINKS_URL='https://www.syncfusion.com/'; npm start
```

```powershell
cd 'D:\TestAppv2\standalone-validator'
$env:HEADER_MENU_LINKS_URL='https://example.com'; npm start
```

To explicitly allow one or more domains, provide a comma-separated list. Exact domains and their subdomains are allowed; staging, development, QA, test, localhost, and IP-based hosts are always rejected:

```powershell
$env:HEADER_MENU_LINKS_URL='https://www.syncfusion.com/'; $env:ALLOWED_DOMAINS='syncfusion.com,cdn.syncfusion.com'; npm start
```

```powershell
cd 'D:\TestAppv2\standalone-validator'
node .\header-menu-link-validator.js
```

## Output
Reports are written under:

```text
./reports/header-menu-links/
```

This folder does not modify the existing dashboard project or app files.

## Validation rules

- Every discovered URL must use `https://`.
- Every discovered URL must belong to an allowed domain. When `ALLOWED_DOMAINS` is not set, the source page domains and their subdomains are allowed.
- Staging, development, QA, test, localhost, and IP-based domains are reported as failures.
- The HTML report supports text search, status filtering, and column sorting. Report files include the date and time in their names.
