#!/usr/bin/env node
'use strict';

const { chromium } = require('playwright');

function parseArgs(argv) {
  const args = {};
  for (let i = 2; i < argv.length; i++) {
    const current = argv[i];
    if (!current.startsWith('--')) {
      continue;
    }
    const key = current.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith('--')) {
      args[key] = 'true';
      continue;
    }
    args[key] = next;
    i++;
  }
  return args;
}

function buildUserAgent(uaType, overrideAgent) {
  if (overrideAgent) {
    return overrideAgent;
  }
  const map = {
    Windows: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    Linux: 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    MacOS: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 13_0) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Safari/605.1.15',
    Android: 'Mozilla/5.0 (Linux; Android 11; Pixel 5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36',
    iPhone: 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.1 Mobile/15E148 Safari/604.1',
    WindowsPhone: 'Mozilla/5.0 (Windows Phone 10.0; Android 6.0.1; Microsoft; Lumia 950) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0 Mobile Safari/537.36 Edge/15.14977'
  };
  return map[uaType] || map.Windows;
}

function resolveViewport(uaType) {
  switch (uaType) {
    case 'Android':
      return { width: 412, height: 915 };
    case 'iPhone':
      return { width: 390, height: 844 };
    case 'WindowsPhone':
      return { width: 360, height: 640 };
    case 'Linux':
    case 'MacOS':
    case 'Windows':
    default:
      return { width: 1280, height: 720 };
  }
}

async function findLocator(page, selectors) {
  for (const selector of selectors) {
    const locator = page.locator(selector).first();
    if (await locator.count() > 0) {
      return selector;
    }
  }
  return null;
}

async function ensureEmailSelector(page) {
  const selectors = ['input[name="loginfmt"]', 'input[type="email"]', '#i0116'];
  let selector = await findLocator(page, selectors);
  if (selector) {
    return selector;
  }

  const triggers = ['text=Sign in', 'text=Sign in to your account', 'text=Sign-in', '#mectrl_headerPicture'];
  for (const trigger of triggers) {
    const btn = page.locator(trigger).first();
    if (await btn.count() > 0) {
      await btn.click({ timeout: 5000 }).catch(() => {});
      await page.waitForTimeout(1500);
      selector = await findLocator(page, selectors);
      if (selector) {
        return selector;
      }
    }
  }

  await page.goto('https://login.microsoftonline.com/', { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {});
  await page.waitForTimeout(1500);
  selector = await findLocator(page, selectors);
  if (selector) {
    return selector;
  }
  throw new Error('Email input not found on the login page.');
}

async function ensurePasswordSelector(page) {
  const selectors = ['input[name="passwd"]', 'input[type="password"]', '#i0118'];
  await page.waitForTimeout(500);
  for (let attempt = 0; attempt < 2; attempt++) {
    const selector = await findLocator(page, selectors);
    if (selector) {
      return selector;
    }
    await page.waitForTimeout(1000);
  }
  await page.waitForSelector(selectors.join(', '), { timeout: 45000 });
  const fallback = await findLocator(page, selectors);
  if (fallback) {
    return fallback;
  }
  throw new Error('Password input not found after submitting username.');
}

async function automateLogin() {
  const args = parseArgs(process.argv);
  const username = process.env.MFASWEEP_USERNAME || '';
  const password = process.env.MFASWEEP_PASSWORD || '';
  const uaType = args.uaType || 'Windows';
  const userAgent = buildUserAgent(uaType, args.userAgent);
  const viewport = resolveViewport(uaType);
  const result = {
    success: false,
    mfaRequired: false,
    error: null,
    cookies: [],
    finalUrl: null
  };

  if (!username || !password) {
    result.error = 'Missing credentials in MFASWEEP_USERNAME/MFASWEEP_PASSWORD environment variables.';
    console.log(JSON.stringify(result));
    process.exit(2);
  }

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    userAgent,
    viewport,
    locale: 'en-US',
    deviceScaleFactor: viewport.width < 500 ? 2 : 1,
    permissions: []
  });
  const page = await context.newPage();

  try {
    await page.goto('https://outlook.office365.com/?authRedirect=true&state=0', { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});

    const emailSelector = await ensureEmailSelector(page);
    await page.fill(emailSelector, username);
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 45000 }).catch(() => {}),
      page.click('input[type="submit"]')
    ]);

    const passwordSelector = await ensurePasswordSelector(page);
    await page.fill(passwordSelector, password, { noWaitAfter: true });
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'load', timeout: 60000 }).catch(() => {}),
      page.click('input[type="submit"]')
    ]);

    await page.waitForTimeout(2000);

    const staySignedInPrompt = await page.locator('text=Stay signed in?').first().count();
    if (staySignedInPrompt > 0) {
      const noButton = page.locator('#idBtn_Back');
      if (await noButton.count() > 0) {
        await noButton.click().catch(() => {});
      }
      else {
        const yesButton = page.locator('#idSIButton9');
        await yesButton.click().catch(() => {});
      }
      await page.waitForTimeout(2000);
    }

    const mfaIndicators = [
      'text=Enter code',
      'text=Approve a request',
      '#idDiv_SAOTCAS_Proofs',
      '#idDiv_SAOTCS_Description',
      '#OtcEntry'
    ];
    for (const selector of mfaIndicators) {
      if (await page.locator(selector).first().count() > 0) {
        result.mfaRequired = true;
        break;
      }
    }

    await page.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => {});
    result.finalUrl = page.url();

    const cookies = await context.cookies();
    result.cookies = cookies;

    const hasEstsAuth = cookies.some((cookie) => cookie.name === 'ESTSAUTH');
    const reachedOutlook = result.finalUrl && result.finalUrl.includes('outlook.office.com');

    if (hasEstsAuth || reachedOutlook || result.mfaRequired) {
      result.success = true;
    }
    else {
      result.error = 'Did not observe successful Outlook session.';
    }
  }
  catch (err) {
    result.error = err.message;
  }
  finally {
    await context.close();
    await browser.close();
  }

  console.log(JSON.stringify(result));
  process.exit(result.success ? 0 : 1);
}

automateLogin();
