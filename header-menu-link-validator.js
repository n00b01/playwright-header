// Standalone validator kept outside the Enterprise Audit Dashboard app.
// This file is intentionally independent and does not modify app source files.

import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';
import { chromium } from '@playwright/test';

const spinnerFrames = ['|', '/', '-', '\\'];

function logStep(message) {
  const timestamp = new Date().toLocaleTimeString();
  console.log(`[${timestamp}] ${message}`);
}

function startSpinner(label) {
  let frame = 0;
  const timer = setInterval(() => {
    process.stdout.write(`\r${label} ${spinnerFrames[frame % spinnerFrames.length]}`);
    frame += 1;
  }, 120);

  return () => {
    clearInterval(timer);
    process.stdout.write('\r\x1b[2K');
  };
}

const config = {
  sourceUrls: [
    'https://example.com'
  ],
  selectors: [
    'header a[href]',
    'header button',
    'header [role="button"]',
    'nav a[href]',
    'nav button',
    'nav [role="button"]',
    '[role="navigation"] a[href]',
    '[role="navigation"] button',
    '[role="navigation"] [role="button"]',
    '.menu a[href]',
    '.menu button',
    '.menu [role="button"]',
    '.header-menu a[href]',
    '.header-menu button',
    '.header-menu [role="button"]'
  ],
  httpTimeout: 20000,
  pageTimeout: 30000,
  outputDir: path.resolve(process.cwd(), 'reports', 'header-menu-links'),
  allowedDomains: []
};

const HEADER_MENU_LINKS_URL = process.env.HEADER_MENU_LINKS_URL;
if (HEADER_MENU_LINKS_URL) {
  config.sourceUrls = [HEADER_MENU_LINKS_URL];
}

const ALLOWED_DOMAINS = process.env.ALLOWED_DOMAINS;
if (ALLOWED_DOMAINS) {
  config.allowedDomains = ALLOWED_DOMAINS.split(',').map((domain) => domain.trim().toLowerCase()).filter(Boolean);
}

function normalizedHostname(value) {
  return String(value || '').trim().toLowerCase().replace(/^\.+|\.+$/g, '');
}

function isAllowedDomain(value, sourceUrls) {
  let hostname;
  try {
    hostname = normalizedHostname(new URL(value).hostname);
  } catch {
    return { allowed: false, reason: 'Invalid URL' };
  }

  if (!hostname) return { allowed: false, reason: 'Missing domain' };
  if (/^(localhost|127(?:\.\d{1,3}){3}|0\.0\.0\.0)$/.test(hostname) || /(^|[-.])(dev|development|staging|stage|qa|test)([-.]|$)/i.test(hostname)) {
    return { allowed: false, reason: `Disallowed environment domain: ${hostname}` };
  }

  const allowedDomains = config.allowedDomains.length
    ? config.allowedDomains
    : sourceUrls.map((sourceUrl) => {
      try {
        return normalizedHostname(new URL(sourceUrl).hostname);
      } catch {
        return '';
      }
    }).filter(Boolean);

  const allowed = allowedDomains.some((domain) => hostname === domain || hostname.endsWith(`.${domain}`));
  return allowed ? { allowed: true, hostname } : { allowed: false, reason: `Domain not allowed: ${hostname}` };
}

function validateUrlPolicy(value, sourceUrls) {
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    return { valid: false, status: 'Invalid URL', reason: 'Invalid URL' };
  }

  if (parsed.protocol !== 'https:') {
    return { valid: false, status: 'Non-HTTPS URL', reason: `Only HTTPS URLs are allowed (${parsed.protocol})` };
  }

  const domainCheck = isAllowedDomain(parsed.toString(), sourceUrls);
  if (!domainCheck.allowed) {
    return { valid: false, status: 'Domain Error', reason: domainCheck.reason };
  }

  return { valid: true, status: '', reason: '' };
}

function isIgnoredHref(href) {
  if (!href) return true;
  const cleaned = href.trim();
  if (!cleaned || cleaned === '#' || cleaned.startsWith('javascript:') || cleaned.startsWith('mailto:') || cleaned.startsWith('tel:') || cleaned.startsWith('sms:')) {
    return true;
  }

  try {
    const url = new URL(cleaned, 'https://placeholder.invalid');
    return !['http:', 'https:'].includes(url.protocol);
  } catch {
    return true;
  }
}

function normalizeStatus(status) {
  if (status === 'No direct URL') return 'No direct URL';
  if (typeof status !== 'number' || Number.isNaN(status)) return 'Broken URL';
  if (status >= 200 && status < 300) return '2xx OK';
  if (status >= 300 && status < 400) return '3xx Redirect';
  if (status >= 400 && status < 500) return '4xx Client Error';
  if (status >= 500 && status < 600) return '5xx Server Error';
  return 'Broken URL';
}

function classifyError(error) {
  if (!error) return 'Unknown error';
  const msg = String(error.message || error).toLowerCase();
  if (msg.includes('timeout')) return 'Timeout';
  if (msg.includes('dns')) return 'DNS Error';
  if (msg.includes('tls') || msg.includes('certificate')) return 'TLS Error';
  if (msg.includes('fetch failed')) return 'Connection Error';
  if (msg.includes('invalid url')) return 'Invalid URL';
  if (msg.includes('redirect')) return 'Redirect Error';
  return 'Request Failure';
}

function toSafeDomain(value) {
  try {
    return new URL(value).hostname;
  } catch {
    return '';
  }
}

function toSafeHttps(value) {
  try {
    return new URL(value).protocol === 'https:';
  } catch {
    return false;
  }
}

function buildRow({ sourceUrl, parentMenu, menuText, href, resolvedUrl, status, notes }) {
  const urlValue = resolvedUrl || href || '';
  const menuContext = parentMenu ? `${parentMenu} > ${menuText || 'Unnamed'}` : (menuText || 'Unnamed');
  return {
    'URL Under Which Header Menu': parentMenu || sourceUrl,
    URL: urlValue,
    'URL Status': status,
    'Is Https URL?': toSafeHttps(urlValue) ? 'Yes' : 'No',
    Domain: toSafeDomain(urlValue),
    Notes: notes || `Menu: ${menuContext}`
  };
}

async function collectChildLinksForTrigger(page, triggerLocator, parentLabel) {
  const handle = await triggerLocator.elementHandle();
  if (!handle) return [];

  try {
    await handle.hover();
  } catch {
    // some elements may not support hover reliably; continue with best effort
  }

  try {
    await handle.click({ trial: true });
  } catch {
    // non-clickable triggers are still valid as hover-driven menu controls
  }

  await page.waitForTimeout(450);

  return triggerLocator.evaluate((node, parentMenu) => {
    const root = node.closest('li, .dropdown, .nav-item, .menu-item, [role="menuitem"], [role="navigation"], nav, header');
    const scopes = [];

    if (root) {
      scopes.push(root);
      root.querySelectorAll('[role="menu"], .dropdown-menu, .mega-menu, .submenu, .menu-panel, ul').forEach((scope) => scopes.push(scope));
    }

    const results = [];
    const seen = new Set();

    const addLink = (anchor) => {
      if (!anchor || !(anchor instanceof Element)) return;
      const href = (anchor.getAttribute('href') || '').trim();
      const text = (anchor.textContent || anchor.getAttribute('aria-label') || anchor.getAttribute('title') || '').trim();
      if (!text || !href || href === '#' || href.startsWith('javascript:') || href.startsWith('mailto:') || href.startsWith('tel:')) {
        return;
      }
      const key = `${href.toLowerCase()}|${text.toLowerCase()}`;
      if (seen.has(key)) return;
      seen.add(key);
      results.push({ href, text, parentMenu });
    };

    if (scopes.length) {
      scopes.forEach((scope) => {
        scope.querySelectorAll('a[href]').forEach((link) => addLink(link));
      });
    }

    if (results.length === 0) {
      document.querySelectorAll('header a[href], nav a[href], [role="navigation"] a[href]').forEach((link) => {
        const text = (link.textContent || link.getAttribute('aria-label') || link.getAttribute('title') || '').trim();
        if (!text) return;
        if (link.closest('header, nav, [role="navigation"]')) {
          addLink(link);
        }
      });
    }

    return results;
  }, parentLabel);
}

async function extractHeaderLinks(page, selectors) {
  const found = [];
  const seen = new Set();

  const ignoredTriggerLabels = new Set(['Sign In', 'Get Free Trial', 'Search', 'Menu', 'Close', 'Profile', 'Login']);

  for (const selector of selectors) {
    const candidates = await page.locator(selector).all();

    for (const candidate of candidates) {
      const label = (await candidate.textContent())?.trim() || '';
      const ariaLabel = (await candidate.getAttribute('aria-label'))?.trim() || '';
      const title = (await candidate.getAttribute('title'))?.trim() || '';
      const href = (await candidate.getAttribute('href'))?.trim() || '';
      const menuText = (label || ariaLabel || title || 'Unnamed menu item').trim();

      if (!menuText || ignoredTriggerLabels.has(menuText)) continue;

      const directLinks = [];
      if (href) {
        directLinks.push({ href, text: menuText, parentMenu: menuText });
      }

      const discoveredLinks = href ? directLinks : await collectChildLinksForTrigger(page, candidate, menuText);

      if (!discoveredLinks.length && !href) {
        continue;
      }

      for (const item of discoveredLinks) {
        const linkText = (item.text || menuText || 'Unnamed menu item').trim();
        const rawHref = (item.href || '').trim();
        if (!rawHref) continue;
        if (isIgnoredHref(rawHref)) continue;

        const parentMenu = item.parentMenu || menuText;
        const dedupeKey = `${rawHref.toLowerCase()}|${parentMenu.toLowerCase()}|${linkText.toLowerCase()}`;
        if (seen.has(dedupeKey)) continue;
        seen.add(dedupeKey);

        try {
          const resolvedUrl = new URL(rawHref, page.url()).toString();
          found.push({
            href: rawHref,
            resolvedUrl,
            menuText: `${parentMenu} > ${linkText}`,
            parentMenu,
            selector
          });
        } catch (error) {
          found.push({
            href: rawHref,
            resolvedUrl: rawHref,
            menuText: `${parentMenu} > ${linkText}`,
            parentMenu,
            selector,
            invalid: true,
            error
          });
        }
      }
    }
  }

  return found;
}

async function validateUrl(context, targetUrl) {
  let finalUrl = targetUrl;
  let statusCode = 0;
  let notes = '';
  let error = null;

  async function request(method) {
    try {
      const response = await context.request.fetch(targetUrl, {
        method,
        failOnStatusCode: false,
        maxRedirects: 0,
        timeout: config.httpTimeout,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });

      finalUrl = response.url?.() || finalUrl;
      statusCode = response.status();

      if (statusCode >= 300 && statusCode < 400 && response.headers()['location']) {
        return { redirectLocation: response.headers()['location'], statusCode };
      }

      return { statusCode };
    } catch (err) {
      error = err;
      return { statusCode: 0, error: err };
    }
  }

  let result = await request('HEAD');
  if (result.error || result.statusCode === 0 || result.statusCode >= 400) {
    result = await request('GET');
  }

  if (result.error) {
    statusCode = 0;
    notes = classifyError(result.error);
    error = result.error;
  } else {
    statusCode = result.statusCode;
  }

  if (statusCode >= 300 && statusCode < 400) {
    notes = `Redirects to ${finalUrl || targetUrl}`;
  }

  if (statusCode === 0) {
    notes = notes || 'Broken URL';
  }

  return {
    finalUrl,
    statusCode,
    status: normalizeStatus(statusCode),
    notes,
    error
  };
}

async function runHeaderMenuLinkAudit() {
  const startedAt = Date.now();
  const stopSpinner = startSpinner('Initializing validator');
  logStep('Starting header menu validation...');
  fs.mkdirSync(config.outputDir, { recursive: true });

  const rows = [];
  const errors = [];
  const sourceSummaries = [];

  stopSpinner();
  logStep('Launching browser');
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();

  for (const sourceUrl of config.sourceUrls) {
    let page;
    try {
      const sourcePolicy = validateUrlPolicy(sourceUrl, config.sourceUrls);
      if (!sourcePolicy.valid) {
        const reason = `${sourcePolicy.status}: ${sourcePolicy.reason}`;
        errors.push({ sourceUrl, reason });
        sourceSummaries.push({ sourceUrl, linksFound: 0, pageStatus: 0, error: reason });
        continue;
      }

      stopSpinner();
      logStep(`Loading page: ${sourceUrl}`);
      page = await context.newPage();
      const response = await page.goto(sourceUrl, {
        waitUntil: 'domcontentloaded',
        timeout: config.pageTimeout
      });

      const spinner = startSpinner('Extracting header menu links');
      const links = await extractHeaderLinks(page, config.selectors);
      spinner();
      const pageStatus = response ? response.status() : 0;

      if (!response) {
        errors.push({ sourceUrl, reason: 'Page load failed: no response object' });
        sourceSummaries.push({ sourceUrl, linksFound: 0, pageStatus: 0, error: 'No response' });
        continue;
      }

      if (pageStatus >= 400) {
        errors.push({ sourceUrl, reason: `Page load returned ${pageStatus}` });
      }

      if (links.length === 0) {
        errors.push({ sourceUrl, reason: 'No header menu links matched the configured selectors' });
      }

      logStep(`Found ${links.length} menu candidates for ${sourceUrl}`);

      for (const item of links) {
        if (!item.href) {
          rows.push(buildRow({
            sourceUrl,
            parentMenu: item.parentMenu,
            menuText: item.menuText,
            href: '',
            resolvedUrl: '',
            status: 'No direct URL',
            notes: `Menu trigger: ${item.menuText || 'Unnamed'} (no direct href)`
          }));
          continue;
        }

        const targetUrl = item.resolvedUrl || item.href;
        const policy = validateUrlPolicy(targetUrl, config.sourceUrls);
        let validation;
        if (!policy.valid) {
          validation = { status: policy.status, notes: policy.reason, finalUrl: targetUrl };
          errors.push({ sourceUrl, reason: `${policy.status}: ${targetUrl} (${policy.reason})` });
        } else {
          const validationSpinner = startSpinner(`Validating ${item.menuText || 'menu item'}`);
          validation = await validateUrl(context, targetUrl);
          validationSpinner();
        }
        const noteParts = [`Menu: ${item.menuText || 'Unnamed'}`];

        if (validation.status === '3xx Redirect') {
          noteParts.push(`Redirects to ${validation.finalUrl || item.resolvedUrl || item.href}`);
        }

        if (validation.notes && validation.notes !== 'Broken URL') {
          noteParts.push(validation.notes);
        }

        rows.push(buildRow({
          sourceUrl,
          parentMenu: item.parentMenu,
          menuText: item.menuText,
          href: item.href,
          resolvedUrl: targetUrl,
          status: validation.status,
          notes: validation.notes ? `${noteParts.join(' | ')} | ${validation.notes}` : noteParts.join(' | ')
        }));
      }

      sourceSummaries.push({
        sourceUrl,
        finalUrl: page.url(),
        pageStatus,
        linksFound: links.length
      });
    } catch (error) {
      errors.push({ sourceUrl, reason: error.message || String(error) });
      sourceSummaries.push({ sourceUrl, linksFound: 0, error: error.message || String(error) });
    } finally {
      if (page) await page.close();
    }
  }

  await browser.close();

  const summarySheet = [{
    'Generated Timestamp': new Date().toISOString(),
    'Source Pages Processed': sourceSummaries.length,
    'Total Links Tested': rows.length,
    'Passed Links': rows.filter(r => r['URL Status'] === '2xx OK').length,
    'Redirects': rows.filter(r => r['URL Status'] === '3xx Redirect').length,
    'Client Errors': rows.filter(r => r['URL Status'] === '4xx Client Error').length,
    'Server Errors': rows.filter(r => r['URL Status'] === '5xx Server Error').length,
    'Broken URLs': rows.filter(r => r['URL Status'] === 'Broken URL').length,
    'Domain Errors': rows.filter(r => r['URL Status'] === 'Domain Error').length,
    'Non-HTTPS URLs': rows.filter(r => r['URL Status'] === 'Non-HTTPS URL').length,
    'Invalid URLs': rows.filter(r => r['URL Status'] === 'Invalid URL').length,
    'Menu Triggers Without Direct URL': rows.filter(r => r['URL Status'] === 'No direct URL').length,
    'Errors Encountered': errors.length
  }];

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows, {
    header: ['URL Under Which Header Menu', 'URL', 'URL Status', 'Is Https URL?', 'Domain', 'Notes']
  });

  sheet['!cols'] = [
    { wch: 40 },
    { wch: 50 },
    { wch: 18 },
    { wch: 14 },
    { wch: 25 },
    { wch: 80 }
  ];

  XLSX.utils.book_append_sheet(workbook, sheet, 'Header Links');

  const summaryWs = XLSX.utils.json_to_sheet(summarySheet);
  XLSX.utils.book_append_sheet(workbook, summaryWs, 'Summary');

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const excelPath = path.join(config.outputDir, `header-menu-links-report-${timestamp}.xlsx`);
  const htmlPath = path.join(config.outputDir, `header-menu-links-report-${timestamp}.html`);

  XLSX.writeFile(workbook, excelPath);

  const escapeHtml = (value) => String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
  const statusCounts = rows.reduce((counts, row) => {
    counts[row['URL Status']] = (counts[row['URL Status']] || 0) + 1;
    return counts;
  }, {});
  const passedCount = statusCounts['2xx OK'] || 0;
  const issueCount = rows.length - passedCount;
  const allowedDomains = config.allowedDomains.length ? config.allowedDomains.join(', ') : 'Source URL domains and their subdomains';
  const statusClass = (status) => status === '2xx OK' ? 'pass' : status === '3xx Redirect' ? 'redirect' : 'fail';
  const htmlRows = rows.map((row) => `
    <tr data-status="${escapeHtml(row['URL Status'])}">
      <td>${escapeHtml(row['URL Under Which Header Menu'])}</td>
      <td class="url-cell">${escapeHtml(row.URL)}</td>
      <td><span class="status ${statusClass(row['URL Status'])}">${escapeHtml(row['URL Status'])}</span></td>
      <td><span class="https ${row['Is Https URL?'] === 'Yes' ? 'yes' : 'no'}">${escapeHtml(row['Is Https URL?'])}</span></td>
      <td>${escapeHtml(row.Domain)}</td>
      <td>${escapeHtml(row.Notes || '')}</td>
    </tr>
  `).join('');

  const html = `
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Header Menu Link Validation Report</title>
        <style>
          :root { --ink:#172033; --muted:#667085; --line:#e5e7eb; --surface:#fff; --canvas:#f4f7fb; --blue:#2563eb; --green:#15803d; --amber:#b45309; --red:#b42318; }
          * { box-sizing: border-box; }
          body { margin: 0; padding: 28px; background: var(--canvas); color: var(--ink); font: 14px/1.5 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
          .shell { max-width: 1500px; margin: auto; }
          .hero { padding: 28px 32px; border-radius: 18px; color: #fff; background: linear-gradient(135deg, #102a43, #1769aa); box-shadow: 0 12px 30px rgba(16,42,67,.18); }
          .hero h1 { margin: 0; font-size: 28px; letter-spacing: -.02em; }
          .hero p { margin: 7px 0 0; color: #dbeafe; }
          .hero code { color: #fff; }
          .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin: 20px 0; }
          .kpi, .panel { background: var(--surface); border: 1px solid var(--line); border-radius: 14px; box-shadow: 0 5px 16px rgba(15,23,42,.04); }
          .kpi { padding: 18px; }
          .kpi .number { font-size: 28px; font-weight: 800; }
          .kpi .label { color: var(--muted); font-size: 11px; font-weight: 700; letter-spacing: .07em; text-transform: uppercase; }
          .kpi.pass .number { color: var(--green); } .kpi.issue .number { color: var(--red); } .kpi.redirect .number { color: var(--amber); }
          .panel { padding: 20px; overflow: hidden; }
          .toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; justify-content: space-between; margin-bottom: 14px; }
          .toolbar h2 { margin: 0; font-size: 18px; }
          .controls { display: flex; flex-wrap: wrap; gap: 8px; }
          input, select { border: 1px solid #cbd5e1; border-radius: 9px; padding: 9px 11px; background: #fff; color: var(--ink); }
          input { min-width: 260px; }
          .hint { color: var(--muted); font-size: 12px; margin: 0 0 12px; }
          .table-wrap { overflow-x: auto; }
          table { border-collapse: collapse; width: 100%; min-width: 980px; }
          th, td { border-bottom: 1px solid var(--line); padding: 11px 12px; text-align: left; vertical-align: top; }
          th { background: #f8fafc; color: #475467; cursor: pointer; font-size: 11px; letter-spacing: .05em; text-transform: uppercase; white-space: nowrap; }
          th:hover { color: var(--blue); }
          tbody tr:hover td { background: #f8fbff; }
          .url-cell { max-width: 430px; color: #475467; font-family: Consolas, "Courier New", monospace; font-size: 12px; overflow-wrap: anywhere; }
          .status, .https { display: inline-block; border-radius: 999px; padding: 3px 9px; font-size: 11px; font-weight: 800; white-space: nowrap; }
          .status.pass, .https.yes { background: #dcfce7; color: var(--green); }
          .status.redirect { background: #fef3c7; color: var(--amber); }
          .status.fail, .https.no { background: #fee4e2; color: var(--red); }
          .meta { color: var(--muted); font-size: 12px; }
          .empty { color: var(--muted); padding: 20px 0; }
          @media (max-width: 700px) { body { padding: 14px; } .hero { padding: 22px; } .hero h1 { font-size: 23px; } input { min-width: 100%; } }
        </style>
      </head>
      <body>
        <main class="shell">
          <header class="hero">
            <h1>Header Menu Link Validation</h1>
            <p>Generated ${escapeHtml(new Date().toISOString())} across ${sourceSummaries.length} source page(s). Allowed domains: <code>${escapeHtml(allowedDomains)}</code></p>
          </header>
          <section class="kpis">
            <div class="kpi"><div class="number">${rows.length}</div><div class="label">Total links tested</div></div>
            <div class="kpi pass"><div class="number">${passedCount}</div><div class="label">2xx passed</div></div>
            <div class="kpi redirect"><div class="number">${statusCounts['3xx Redirect'] || 0}</div><div class="label">Redirects</div></div>
            <div class="kpi issue"><div class="number">${issueCount}</div><div class="label">Failures</div></div>
          </section>
          <section class="panel">
            <div class="toolbar">
              <h2>Validated menu URLs</h2>
              <div class="controls">
                <input id="search" type="search" placeholder="Search URL, domain, menu..." aria-label="Search validated URLs" />
                <select id="statusFilter" aria-label="Filter by URL status">
                  <option value="all">All statuses</option>
                  ${Object.keys(statusCounts).sort().map((status) => `<option value="${escapeHtml(status)}">${escapeHtml(status)} (${statusCounts[status]})</option>`).join('')}
                </select>
              </div>
            </div>
            <p class="hint"><span id="visibleCount">${rows.length}</span> visible row(s). Click a column heading to sort.</p>
            <div class="table-wrap">
              <table id="resultsTable">
                <thead><tr><th>URL Under Which Header Menu</th><th>URL</th><th>URL Status</th><th>Is Https URL?</th><th>Domain</th><th>Notes</th></tr></thead>
                <tbody>${htmlRows || '<tr><td colspan="6" class="empty">No links found</td></tr>'}</tbody>
              </table>
            </div>
          </section>
          <section class="panel" style="margin-top:20px">
            <h2>Errors and policy failures</h2>
            ${errors.length ? `<ul>${errors.map((e) => `<li>${escapeHtml(e.sourceUrl)}: ${escapeHtml(e.reason)}</li>`).join('')}</ul>` : '<p class="meta">No errors</p>'}
          </section>
        </main>
        <script>
          (() => {
            const table = document.querySelector('#resultsTable');
            const search = document.querySelector('#search');
            const filter = document.querySelector('#statusFilter');
            const visibleCount = document.querySelector('#visibleCount');
            const update = () => {
              const query = search.value.toLowerCase().trim();
              let visible = 0;
              table.querySelectorAll('tbody tr').forEach((row) => {
                const matchesStatus = filter.value === 'all' || row.dataset.status === filter.value;
                const matchesSearch = !query || row.textContent.toLowerCase().includes(query);
                row.hidden = !(matchesStatus && matchesSearch);
                if (!row.hidden) visible += 1;
              });
              visibleCount.textContent = visible;
            };
            search.addEventListener('input', update);
            filter.addEventListener('change', update);
            table.querySelectorAll('th').forEach((heading, index) => heading.addEventListener('click', () => {
              const rows = Array.from(table.querySelectorAll('tbody tr'));
              const ascending = heading.dataset.order !== 'asc';
              heading.dataset.order = ascending ? 'asc' : 'desc';
              rows.sort((a, b) => {
                const left = a.cells[index].textContent.trim().toLowerCase();
                const right = b.cells[index].textContent.trim().toLowerCase();
                return left.localeCompare(right, undefined, { numeric: true }) * (ascending ? 1 : -1);
              });
              rows.forEach((row) => table.querySelector('tbody').appendChild(row));
            }));
          })();
        </script>
      </body>
    </html>
  `;

  fs.writeFileSync(htmlPath, html);

  const totalMs = Date.now() - startedAt;
  const totalSeconds = (totalMs / 1000).toFixed(2);

  logStep(`Completed validation in ${totalSeconds}s`);
  console.log(`[header-menu-link-validator] Excel report: ${excelPath}`);
  console.log(`[header-menu-link-validator] HTML report: ${htmlPath}`);
  console.log(`[header-menu-link-validator] Total links: ${rows.length}`);
  console.log(`[header-menu-link-validator] Errors: ${errors.length}`);
  console.log(`[header-menu-link-validator] Total time: ${totalSeconds}s`);

  return { rows, excelPath, htmlPath, errors, totalMs };
}

runHeaderMenuLinkAudit().catch((error) => {
  console.error('[header-menu-link-validator] Fatal error:', error);
  process.exit(1);
});
