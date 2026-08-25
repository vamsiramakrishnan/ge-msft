import { createHash } from 'node:crypto';
import {
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { basename, dirname, join, relative } from 'node:path';
import { spawnSync } from 'node:child_process';

export const repoRoot = new URL('../..', import.meta.url).pathname.replace(/\/$/, '');
export const alphaProfile = 'internal-alpha-word-excel';
export const devProfile = 'development';
const DEFAULT_DEV_PORT = '13000';

export function parseArgs(argv = process.argv.slice(2)) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (!arg.startsWith('--')) continue;
    const key = arg.slice(2);
    const next = argv[i + 1];
    if (next && !next.startsWith('--')) {
      out[key] = next;
      i++;
    } else {
      out[key] = true;
    }
  }
  return out;
}

export function packageJson() {
  return JSON.parse(readFileSync(join(repoRoot, 'package.json'), 'utf8'));
}

export function rootVersion() {
  return packageJson().version;
}

export function officeXmlVersion() {
  const version = rootVersion();
  return version.startsWith('0.') ? '1.0.0' : version;
}

function parseDotEnv(path) {
  if (!existsSync(path)) return {};
  const values = {};
  for (const rawLine of readFileSync(path, 'utf8').split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) continue;
    const match = /^(?:export\s+)?([A-Z0-9_]+)\s*=\s*(.*)$/i.exec(line);
    if (!match) continue;
    const [, key, rawValue] = match;
    const value = rawValue.replace(/^['"]|['"]$/g, '');
    values[key] = value;
  }
  return values;
}

function devWebOrigin(env) {
  if (env.GE_DEV_WEB_ORIGIN) return env.GE_DEV_WEB_ORIGIN;
  if (env.GOOGLE_CLOUD_WORKSTATIONS === 'true' && env.WEB_HOST) {
    return `https://${env.GE_DEV_PORT ?? DEFAULT_DEV_PORT}-${env.WEB_HOST}`;
  }
  return `https://localhost:${env.GE_DEV_PORT ?? DEFAULT_DEV_PORT}`;
}

export function profileFromArgs(args) {
  const profile = String(args.profile ?? alphaProfile);
  if (profile !== alphaProfile && profile !== devProfile) {
    throw new Error(`Unsupported profile "${profile}".`);
  }
  return profile;
}

export function command(cmd, args, options = {}) {
  const start = Date.now();
  const res = spawnSync(cmd, args, {
    cwd: repoRoot,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    ...options,
  });
  return {
    command: [cmd, ...args].join(' '),
    status: res.status ?? 1,
    ok: res.status === 0,
    durationMs: Date.now() - start,
    stdout: res.stdout ?? '',
    stderr: res.stderr ?? '',
  };
}

export function gitSha() {
  const res = command('git', ['rev-parse', 'HEAD']);
  return res.ok ? res.stdout.trim() : 'unknown';
}

export function gitDirty() {
  const res = command('git', ['status', '--short']);
  return !res.ok || res.stdout.trim().length > 0;
}

export function nodeVersion() {
  return process.version;
}

export function bunVersion() {
  const res = command('bun', ['--version']);
  return res.ok ? res.stdout.trim() : 'unknown';
}

export function ensureDir(path) {
  mkdirSync(path, { recursive: true });
}

export function writeJson(path, value) {
  ensureDir(dirname(path));
  writeFileSync(path, `${JSON.stringify(value, null, 2)}\n`);
}

export function walk(dir, files = []) {
  if (!existsSync(dir)) return files;
  if (statSync(dir).isFile()) {
    files.push(dir);
    return files.sort();
  }
  for (const entry of readdirSync(dir)) {
    const path = join(dir, entry);
    const st = statSync(path);
    if (st.isDirectory()) walk(path, files);
    else files.push(path);
  }
  return files.sort();
}

export function sha256File(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

export function sha256Text(text) {
  return createHash('sha256').update(text).digest('hex');
}

export function copyFile(src, dest) {
  ensureDir(dirname(dest));
  writeFileSync(dest, readFileSync(src));
}

export function copyDir(src, dest) {
  if (!existsSync(src)) return;
  for (const file of walk(src)) {
    copyFile(file, join(dest, relative(src, file)));
  }
}

export function cleanDir(dir) {
  rmSync(dir, { recursive: true, force: true });
  ensureDir(dir);
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const HOST = /^[a-z0-9][a-z0-9.-]*[a-z0-9]$/i;
const DEV_GUIDS = new Set([
  '00000000-0000-0000-0000-000000000000',
  '11111111-1111-1111-1111-111111111111',
  '22222222-2222-2222-2222-222222222222',
]);

export function releaseConfig(profile, env = process.env) {
  if (profile === devProfile) {
    const devEnv = {
      ...parseDotEnv(join(repoRoot, 'packages', 'web-shell', '.env')),
      ...env,
    };
    const webOrigin = devWebOrigin(devEnv);
    const webUrl = new URL(webOrigin);
    return {
      profile,
      appId: devEnv.GE_DEV_APP_ID ?? '11111111-1111-4111-8111-111111111111',
      officeXmlAppId: devEnv.GE_DEV_OFFICE_XML_APP_ID ?? '44444444-4444-4444-8444-444444444444',
      outlookAppId: devEnv.GE_DEV_OUTLOOK_APP_ID ?? '88888888-8888-4888-8888-888888888888',
      oneNoteAppId: devEnv.GE_DEV_ONENOTE_APP_ID ?? '33333333-3333-4333-8333-333333333333',
      entraClientId:
        devEnv.GE_DEV_ENTRA_CLIENT_ID ??
        devEnv.VITE_ENTRA_CLIENT_ID ??
        '22222222-2222-4222-8222-222222222222',
      webDomain: devEnv.GE_DEV_WEB_DOMAIN ?? webUrl.hostname,
      webOrigin,
      developerName: devEnv.GE_DEV_DEVELOPER_NAME ?? 'Gemini Enterprise Dev',
      websiteUrl: devEnv.GE_DEV_WEBSITE_URL ?? `${webOrigin}/`,
      privacyUrl: devEnv.GE_DEV_PRIVACY_URL ?? `${webOrigin}/privacy`,
      termsUrl: devEnv.GE_DEV_TERMS_URL ?? `${webOrigin}/terms`,
      supportUrl: devEnv.GE_DEV_SUPPORT_URL ?? `${webOrigin}/support`,
    };
  }

  const required = [
    'GE_ALPHA_APP_ID',
    'GE_ALPHA_ENTRA_CLIENT_ID',
    'GE_ALPHA_WEB_DOMAIN',
    'GE_ALPHA_DEVELOPER_NAME',
    'GE_ALPHA_WEBSITE_URL',
    'GE_ALPHA_PRIVACY_URL',
    'GE_ALPHA_TERMS_URL',
    'GE_ALPHA_SUPPORT_URL',
  ];
  const missing = required.filter((key) => !env[key]);
  if (missing.length) {
    const err = new Error(`Missing release manifest configuration: ${missing.join(', ')}`);
    err.code = 'BLOCKED_EXTERNAL';
    throw err;
  }
  const cfg = {
    profile,
    appId: env.GE_ALPHA_APP_ID,
    entraClientId: env.GE_ALPHA_ENTRA_CLIENT_ID,
    webOrigin: `https://${env.GE_ALPHA_WEB_DOMAIN}`,
    webDomain: env.GE_ALPHA_WEB_DOMAIN,
    developerName: env.GE_ALPHA_DEVELOPER_NAME,
    websiteUrl: env.GE_ALPHA_WEBSITE_URL,
    privacyUrl: env.GE_ALPHA_PRIVACY_URL,
    termsUrl: env.GE_ALPHA_TERMS_URL,
    supportUrl: env.GE_ALPHA_SUPPORT_URL,
  };
  validateReleaseConfig(cfg);
  return cfg;
}

export function validateReleaseConfig(cfg) {
  if (!GUID.test(cfg.appId) || DEV_GUIDS.has(cfg.appId.toLowerCase())) {
    throw new Error('GE_ALPHA_APP_ID must be a non-development GUID.');
  }
  if (!GUID.test(cfg.entraClientId) || DEV_GUIDS.has(cfg.entraClientId.toLowerCase())) {
    throw new Error('GE_ALPHA_ENTRA_CLIENT_ID must be a non-development GUID.');
  }
  if (!HOST.test(cfg.webDomain)) throw new Error('GE_ALPHA_WEB_DOMAIN must be a bare DNS host.');
  if (/localhost|example\.com/i.test(cfg.webDomain)) {
    throw new Error('GE_ALPHA_WEB_DOMAIN must not be localhost or example.com.');
  }
  for (const key of ['websiteUrl', 'privacyUrl', 'termsUrl', 'supportUrl']) {
    const url = new URL(cfg[key]);
    if (url.protocol !== 'https:') throw new Error(`${key} must be HTTPS.`);
    if (url.username || url.password) throw new Error(`${key} must not contain credentials.`);
    if (/localhost|example\.com/i.test(url.hostname)) {
      throw new Error(`${key} must not use localhost or example.com.`);
    }
  }
}

export function generatedManifestPath(profile) {
  return join(repoRoot, 'dist', 'manifests', `${profile}.manifest.json`);
}

export function generatedOneNoteManifestPath(profile) {
  return join(repoRoot, 'dist', 'manifests', `${profile}.onenote.manifest.xml`);
}

export function generatedOfficeXmlManifestPath(profile, surface) {
  return join(repoRoot, 'dist', 'manifests', `${profile}.${surface}.manifest.xml`);
}

export function packageDir(profile) {
  return join(repoRoot, 'dist', 'package', profile);
}

export function packageZip(profile) {
  if (profile === devProfile) {
    return join(repoRoot, 'dist', 'release', `development-m365-v${rootVersion()}.zip`);
  }
  return join(repoRoot, 'dist', 'release', `${profile}-v${rootVersion()}.zip`);
}

function ribbonIcons(origin, extension = 'png') {
  return [16, 32, 80].map((size) => ({
    size,
    url: `${origin}/icon-${size}.${extension}`,
  }));
}

function supertip(title, description) {
  return { title, description };
}

function askSelectionMenu(origin, entryPoint, suffix, description) {
  const target = entryPoint === 'cell' ? 'range' : 'selection';
  return {
    entryPoint,
    controls: [
      {
        id: `geminiAsk${suffix}Menu`,
        type: 'menu',
        label: 'Gemini Enterprise',
        icons: ribbonIcons(origin),
        supertip: supertip('Gemini Enterprise', description),
        items: [
          {
            id: `geminiSummarize${suffix}`,
            type: 'menuItem',
            label: `Summarize ${target}`,
            supertip: supertip(
              `Summarize ${target}`,
              `Summarize the current ${target} in the Gemini pane.`,
            ),
            actionId: 'summarizeSelection',
          },
          {
            id: `geminiExplain${suffix}`,
            type: 'menuItem',
            label: `Explain ${target}`,
            supertip: supertip(
              `Explain ${target}`,
              `Explain the current ${target} in the Gemini pane.`,
            ),
            actionId: 'explainSelection',
          },
        ],
      },
    ],
  };
}

export function alphaManifest(cfg) {
  const domain = cfg.webDomain;
  const origin = cfg.webOrigin ?? `https://${domain}`;
  return {
    $schema: 'https://developer.microsoft.com/json-schemas/teams/v1.23/MicrosoftTeams.schema.json',
    manifestVersion: '1.23',
    id: cfg.appId,
    version: rootVersion(),
    name: { short: 'Gemini Enterprise', full: 'Gemini Enterprise for Microsoft 365 Alpha' },
    description: {
      short: 'Internal Word and Excel alpha for grounded Gemini Enterprise assistance.',
      full: 'Internal tenant alpha for Word and Excel only. Changes are reviewable, gated, and require durable provenance.',
    },
    developer: {
      name: cfg.developerName,
      websiteUrl: cfg.websiteUrl,
      privacyUrl: cfg.privacyUrl,
      termsOfUseUrl: cfg.termsUrl,
    },
    icons: { color: 'icon-color.png', outline: 'icon-outline.png' },
    accentColor: '#1E6B52',
    webApplicationInfo: {
      id: cfg.entraClientId,
      resource: `api://${domain}/${cfg.entraClientId}`,
    },
    validDomains: [domain, 'login.microsoftonline.com'],
    extensions: [
      {
        requirements: {
          scopes: ['document', 'workbook'],
        },
        runtimes: [
          {
            id: 'taskpane_runtime',
            type: 'general',
            code: { page: `${origin}/taskpane.html` },
            lifetime: 'short',
            actions: [{ id: 'openPanel', type: 'openPage', view: 'Gemini' }],
          },
          {
            id: 'commands_runtime',
            type: 'general',
            code: {
              page: `${origin}/commands.html`,
              script: `${origin}/assets/commands.js?v=${rootVersion()}`,
            },
            lifetime: 'short',
            actions: [
              { id: 'openGemini', type: 'executeFunction' },
              { id: 'askSelection', type: 'executeFunction' },
              { id: 'summarizeSelection', type: 'executeFunction' },
              { id: 'explainSelection', type: 'executeFunction' },
            ],
          },
        ],
        ribbons: [
          {
            contexts: ['default'],
            tabs: [
              {
                builtInTabId: 'TabHome',
                groups: [
                  {
                    id: 'geminiGroup',
                    label: 'Gemini Enterprise',
                    icons: ribbonIcons(origin),
                    controls: [
                      {
                        id: 'openGeminiBtn',
                        type: 'button',
                        label: 'Open Gemini',
                        actionId: 'openPanel',
                        icons: ribbonIcons(origin),
                        supertip: supertip('Open Gemini', 'Open the Gemini Enterprise task pane.'),
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
        contextMenus: [
          {
            requirements: { scopes: ['document', 'workbook'] },
            menus: [
              askSelectionMenu(origin, 'text', 'Text', 'Ask about the current Word selection.'),
              askSelectionMenu(origin, 'cell', 'Cell', 'Ask about the current Excel range.'),
            ],
          },
        ],
      },
    ],
  };
}

export function developmentManifest(cfg) {
  const origin = cfg.webOrigin;
  const domain = cfg.webDomain;
  return {
    $schema: 'https://developer.microsoft.com/json-schemas/teams/v1.23/MicrosoftTeams.schema.json',
    manifestVersion: '1.23',
    id: cfg.appId,
    version: rootVersion(),
    name: { short: 'Gemini Enterprise Dev', full: 'Gemini Enterprise for Microsoft 365 Dev' },
    description: {
      short: 'Development package for Gemini Enterprise across Microsoft Office.',
      full: 'Development sideload package for Word, Excel, PowerPoint, and Outlook. Not a production release artifact.',
    },
    developer: {
      name: cfg.developerName,
      websiteUrl: cfg.websiteUrl,
      privacyUrl: cfg.privacyUrl,
      termsOfUseUrl: cfg.termsUrl,
    },
    icons: { color: 'icon-color.png', outline: 'icon-outline.png' },
    accentColor: '#5B5FC7',
    webApplicationInfo: {
      id: cfg.entraClientId,
      resource: `api://${domain}/${cfg.entraClientId}`,
    },
    validDomains: [domain, 'login.microsoftonline.com'],
    extensions: [
      {
        requirements: {
          scopes: ['mail', 'workbook', 'document', 'presentation'],
        },
        runtimes: [
          {
            id: 'taskpane_runtime',
            type: 'general',
            code: { page: `${origin}/taskpane.html` },
            lifetime: 'short',
            actions: [{ id: 'openPanel', type: 'openPage', view: 'Gemini' }],
          },
          {
            id: 'commands_runtime',
            type: 'general',
            code: {
              page: `${origin}/commands.html`,
              script: `${origin}/assets/commands.js?v=${rootVersion()}`,
            },
            lifetime: 'short',
            actions: [
              { id: 'openGemini', type: 'executeFunction' },
              { id: 'onMessageSend', type: 'executeFunction' },
              { id: 'askSelection', type: 'executeFunction' },
              { id: 'summarizeSelection', type: 'executeFunction' },
              { id: 'explainSelection', type: 'executeFunction' },
            ],
          },
        ],
        ribbons: [
          {
            contexts: ['default'],
            tabs: [
              {
                builtInTabId: 'TabHome',
                groups: [
                  {
                    id: 'geminiGroup',
                    label: 'Gemini Enterprise',
                    icons: ribbonIcons(origin),
                    controls: [
                      {
                        id: 'openGeminiBtn',
                        type: 'button',
                        label: 'Open Gemini',
                        actionId: 'openPanel',
                        icons: ribbonIcons(origin),
                        supertip: supertip('Open Gemini', 'Open the Gemini Enterprise task pane.'),
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
        contextMenus: [
          {
            requirements: { scopes: ['document', 'workbook', 'presentation'] },
            menus: [
              askSelectionMenu(
                origin,
                'text',
                'Text',
                'Summarize or review the current selection in the Gemini pane.',
              ),
              askSelectionMenu(
                origin,
                'cell',
                'Cell',
                'Summarize or review the current spreadsheet range in the Gemini pane.',
              ),
            ],
          },
        ],
        autoRunEvents: [
          {
            requirements: {
              capabilities: [{ name: 'Mailbox', minVersion: '1.12' }],
              scopes: ['mail'],
            },
            events: [
              {
                type: 'messageSending',
                actionId: 'onMessageSend',
                options: { sendMode: 'softBlock' },
              },
            ],
          },
        ],
      },
    ],
  };
}

export function oneNoteManifest(cfg) {
  const origin = cfg.webOrigin;
  const esc = xmlEscape;
  return `<?xml version="1.0" encoding="UTF-8"?>
<!-- Development OneNote manifest. OneNote ships separately from the unified M365 package. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>${esc(cfg.oneNoteAppId)}</Id>
  <Version>${esc(officeXmlVersion())}</Version>
  <ProviderName>${esc(cfg.developerName)}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Gemini Enterprise Dev (OneNote)" />
  <Description DefaultValue="Development OneNote add-in for Gemini Enterprise." />
  <IconUrl DefaultValue="${esc(origin)}/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${esc(origin)}/icon-64.png" />
  <SupportUrl DefaultValue="${esc(cfg.supportUrl)}" />
  <AppDomains>
    <AppDomain>${esc(origin)}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Notebook" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${esc(origin)}/taskpane.html?host=onenote" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Notebook">
        <DesktopFormFactor>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="geminiGroup">
                <Label resid="Gemini.Group" />
                <Control xsi:type="Button" id="openGeminiBtn">
                  <Label resid="Gemini.Open" />
                  <Supertip>
                    <Title resid="Gemini.Open" />
                    <Description resid="Gemini.Desc" />
                  </Supertip>
                  <Icon><bt:Image size="16" resid="Icon.16" /><bt:Image size="32" resid="Icon.32" /></Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>GeminiPane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          <ExtensionPoint xsi:type="ContextMenu">
            <OfficeMenu id="ContextMenuText">
              <Control xsi:type="Button" id="GeminiAskCtx">
                <Label resid="Gemini.Ask" />
                <Supertip>
                  <Title resid="Gemini.Ask" />
                  <Description resid="Gemini.Desc" />
                </Supertip>
                <Icon><bt:Image size="16" resid="Icon.16" /><bt:Image size="32" resid="Icon.32" /></Icon>
                <Action xsi:type="ShowTaskpane">
                  <TaskpaneId>GeminiPane</TaskpaneId>
                  <SourceLocation resid="Taskpane.Url" />
                </Action>
              </Control>
            </OfficeMenu>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${esc(origin)}/icon-16.png" />
        <bt:Image id="Icon.32" DefaultValue="${esc(origin)}/icon-32.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${esc(origin)}/taskpane.html?host=onenote" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Gemini.Group" DefaultValue="Gemini Enterprise" />
        <bt:String id="Gemini.Open" DefaultValue="Open Gemini" />
        <bt:String id="Gemini.Ask" DefaultValue="Ask Gemini about this" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Gemini.Desc" DefaultValue="Capture and synthesize research grounded on your unit." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

export function taskPaneXmlManifest(cfg, surface) {
  const origin = cfg.webOrigin;
  const esc = xmlEscape;
  const hostBySurface = {
    word: 'Document',
    excel: 'Workbook',
    powerpoint: 'Presentation',
  };
  const titleBySurface = {
    word: 'Word',
    excel: 'Excel',
    powerpoint: 'PowerPoint',
  };
  const idBySurface = {
    word: cfg.wordAppId ?? '55555555-5555-4555-8555-555555555555',
    excel: cfg.excelAppId ?? '66666666-6666-4666-8666-666666666666',
    powerpoint: cfg.powerPointAppId ?? '77777777-7777-4777-8777-777777777777',
  };
  const host = hostBySurface[surface];
  const title = titleBySurface[surface];
  const id = idBySurface[surface];
  if (!host || !title || !id) throw new Error(`Unsupported task pane XML surface: ${surface}`);

  // Excel only: point Office at the =GE.ASK streaming custom-function metadata. The
  // FunctionFile in VersionOverrides names the page that associates the implementation.
  const customFunctionsExtendedOverrides =
    surface === 'excel'
      ? `
  <ExtendedOverrides Url="${esc(origin)}/functions.json" />`
      : '';
  const functionFile =
    surface === 'excel'
      ? `
          <FunctionFile resid="Functions.Url" />`
      : '';
  const functionUrl =
    surface === 'excel'
      ? `
        <bt:Url id="Functions.Url" DefaultValue="${esc(origin)}/functions.html" />`
      : '';

  return `<?xml version="1.0" encoding="UTF-8"?>
<!-- Development ${esc(title)} XML manifest for Office web Upload Add-in testing. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xsi:type="TaskPaneApp">
  <Id>${esc(id)}</Id>
  <Version>${esc(officeXmlVersion())}</Version>
  <ProviderName>${esc(cfg.developerName)}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Gemini Enterprise Dev (${esc(title)})" />
  <Description DefaultValue="Development ${esc(title)} add-in for Gemini Enterprise." />
  <IconUrl DefaultValue="${esc(origin)}/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${esc(origin)}/icon-64.png" />
  <SupportUrl DefaultValue="${esc(cfg.supportUrl)}" />
  <AppDomains>
    <AppDomain>${esc(origin)}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="${esc(host)}" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${esc(origin)}/taskpane.html?host=${esc(surface)}" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
                    xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="${esc(host)}">
        <DesktopFormFactor>${functionFile}
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="geminiGroup">
                <Label resid="Gemini.Group" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16" />
                  <bt:Image size="32" resid="Icon.32" />
                  <bt:Image size="80" resid="Icon.80" />
                </Icon>
                <Control xsi:type="Button" id="openGeminiBtn">
                  <Label resid="Gemini.Open" />
                  <Supertip>
                    <Title resid="Gemini.Open" />
                    <Description resid="Gemini.Desc" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16" />
                    <bt:Image size="32" resid="Icon.32" />
                    <bt:Image size="80" resid="Icon.80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>GeminiPane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${esc(origin)}/icon-16.png" />
        <bt:Image id="Icon.32" DefaultValue="${esc(origin)}/icon-32.png" />
        <bt:Image id="Icon.80" DefaultValue="${esc(origin)}/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${esc(origin)}/taskpane.html?host=${esc(surface)}" />${functionUrl}
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Gemini.Group" DefaultValue="Gemini Enterprise" />
        <bt:String id="Gemini.Open" DefaultValue="Open Gemini" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Gemini.Desc" DefaultValue="Open the Gemini Enterprise task pane." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>${customFunctionsExtendedOverrides}
</OfficeApp>
`;
}

export function multiHostOfficeXmlManifest(cfg) {
  const origin = cfg.webOrigin;
  const esc = xmlEscape;
  const id = cfg.officeXmlAppId ?? '44444444-4444-4444-8444-444444444444';

  const hostBlock = (host, buttonId, functionUrlId = 'Commands.Url') => `
      <Host xsi:type="${esc(host)}">
        <DesktopFormFactor>
          <FunctionFile resid="${esc(functionUrlId)}" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="geminiGroup">
                <Label resid="Gemini.Group" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16" />
                  <bt:Image size="32" resid="Icon.32" />
                  <bt:Image size="80" resid="Icon.80" />
                </Icon>
                <Control xsi:type="Button" id="${esc(buttonId)}">
                  <Label resid="Gemini.Open" />
                  <Supertip>
                    <Title resid="Gemini.Open" />
                    <Description resid="Gemini.Desc" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16" />
                    <bt:Image size="32" resid="Icon.32" />
                    <bt:Image size="80" resid="Icon.80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>GeminiPane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>`;

  return `<?xml version="1.0" encoding="UTF-8"?>
<!-- Centralized deployment manifest for Word, Excel, and PowerPoint. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xsi:type="TaskPaneApp">
  <Id>${esc(id)}</Id>
  <Version>${esc(officeXmlVersion())}</Version>
  <ProviderName>${esc(cfg.developerName)}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Gemini Enterprise Dev" />
  <Description DefaultValue="Gemini Enterprise for Word, Excel, and PowerPoint." />
  <IconUrl DefaultValue="${esc(origin)}/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${esc(origin)}/icon-64.png" />
  <SupportUrl DefaultValue="${esc(cfg.supportUrl)}" />
  <AppDomains>
    <AppDomain>${esc(origin)}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document" />
    <Host Name="Workbook" />
    <Host Name="Presentation" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${esc(origin)}/taskpane.html" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
                    xsi:type="VersionOverridesV1_0">
    <Hosts>${hostBlock('Document', 'openGeminiWordBtn')}${hostBlock(
      'Workbook',
      'openGeminiExcelBtn',
      'Functions.Url',
    )}${hostBlock('Presentation', 'openGeminiPowerPointBtn')}
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${esc(origin)}/icon-16.png" />
        <bt:Image id="Icon.32" DefaultValue="${esc(origin)}/icon-32.png" />
        <bt:Image id="Icon.80" DefaultValue="${esc(origin)}/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${esc(origin)}/taskpane.html" />
        <bt:Url id="Commands.Url" DefaultValue="${esc(origin)}/commands.html" />
        <bt:Url id="Functions.Url" DefaultValue="${esc(origin)}/functions.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Gemini.Group" DefaultValue="Gemini Enterprise" />
        <bt:String id="Gemini.Open" DefaultValue="Open Gemini" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Gemini.Desc" DefaultValue="Open the Gemini Enterprise task pane." />
      </bt:LongStrings>
    </Resources>
    <WebApplicationInfo>
      <Id>${esc(cfg.entraClientId)}</Id>
      <Resource>api://${esc(cfg.webDomain)}/${esc(cfg.entraClientId)}</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
  <ExtendedOverrides Url="${esc(origin)}/functions.json" />
</OfficeApp>
`;
}

export function outlookXmlManifest(cfg) {
  const origin = cfg.webOrigin;
  const esc = xmlEscape;
  const id = cfg.outlookAppId ?? '88888888-8888-4888-8888-888888888888';
  return `<?xml version="1.0" encoding="UTF-8"?>
<!-- Development Outlook XML manifest for Outlook web Upload Add-in testing. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
           xsi:type="MailApp">
  <Id>${esc(id)}</Id>
  <Version>${esc(officeXmlVersion())}</Version>
  <ProviderName>${esc(cfg.developerName)}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Gemini Enterprise Dev (Outlook)" />
  <Description DefaultValue="Development Outlook add-in for Gemini Enterprise." />
  <IconUrl DefaultValue="${esc(origin)}/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${esc(origin)}/icon-64.png" />
  <SupportUrl DefaultValue="${esc(cfg.supportUrl)}" />
  <AppDomains>
    <AppDomain>${esc(origin)}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.12" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="${esc(origin)}/taskpane.html?host=outlook" />
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="${esc(origin)}/taskpane.html?host=outlook" />
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteMailbox</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
                    xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.12">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="geminiGroup">
                <Label resid="Gemini.Group" />
                <Control xsi:type="Button" id="openGeminiRead">
                  <Label resid="Gemini.Open" />
                  <Supertip>
                    <Title resid="Gemini.Open" />
                    <Description resid="Gemini.Desc" />
                  </Supertip>
                  <Icon><bt:Image size="16" resid="Icon.16" /><bt:Image size="32" resid="Icon.32" /><bt:Image size="80" resid="Icon.80" /></Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          <ExtensionPoint xsi:type="MessageComposeCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="geminiGroupCompose">
                <Label resid="Gemini.Group" />
                <Control xsi:type="Button" id="openGeminiCompose">
                  <Label resid="Gemini.Open" />
                  <Supertip>
                    <Title resid="Gemini.Open" />
                    <Description resid="Gemini.Desc" />
                  </Supertip>
                  <Icon><bt:Image size="16" resid="Icon.16" /><bt:Image size="32" resid="Icon.32" /><bt:Image size="80" resid="Icon.80" /></Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${esc(origin)}/icon-16.png" />
        <bt:Image id="Icon.32" DefaultValue="${esc(origin)}/icon-32.png" />
        <bt:Image id="Icon.80" DefaultValue="${esc(origin)}/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${esc(origin)}/taskpane.html?host=outlook" />
        <bt:Url id="Commands.Url" DefaultValue="${esc(origin)}/commands.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Gemini.Group" DefaultValue="Gemini Enterprise" />
        <bt:String id="Gemini.Open" DefaultValue="Open Gemini" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Gemini.Desc" DefaultValue="Open the Gemini Enterprise task pane." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

function xmlEscape(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('"', '&quot;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

export function validateGeneratedManifest(manifest, profile) {
  const errors = [];
  const text = JSON.stringify(manifest);
  const forbiddenTokens =
    profile === alphaProfile
      ? ['REPLACE_', 'example.com', 'localhost']
      : ['REPLACE_', 'example.com'];
  for (const token of forbiddenTokens) {
    if (text.includes(token)) errors.push(`manifest contains forbidden token ${token}`);
  }
  if (/\{\{[^}]+\}\}/.test(text)) errors.push('manifest contains unresolved template syntax');
  if (profile === alphaProfile) {
    const ext = manifest.extensions?.[0];
    const scopes = ext?.requirements?.scopes ?? [];
    for (const allowed of scopes) {
      if (!['document', 'workbook'].includes(allowed)) {
        errors.push(`alpha manifest advertises disabled scope ${allowed}`);
      }
    }
    for (const disabled of ['mail', 'presentation', 'Notebook', 'team', 'groupChat']) {
      if (text.includes(disabled))
        errors.push(`alpha manifest advertises disabled surface ${disabled}`);
    }
    if (
      manifest.authorization ||
      manifest.bots ||
      manifest.composeExtensions ||
      manifest.staticTabs
    ) {
      errors.push('alpha manifest includes Teams/Outlook-only top-level blocks');
    }
    if (manifest.version !== rootVersion()) {
      errors.push(
        `manifest version ${manifest.version} does not match package version ${rootVersion()}`,
      );
    }
  } else if (
    manifest.authorization ||
    manifest.bots ||
    manifest.composeExtensions ||
    manifest.staticTabs
  ) {
    errors.push('development Office manifest includes Teams-only top-level blocks');
  }
  return errors;
}

const SECRET_PATTERNS = [
  /ENTRA_CLIENT_SECRET/i,
  /GOOGLE_APPLICATION_CREDENTIALS/i,
  /SERVICE_ACCOUNT_JSON/i,
  /"private_key"\s*:/i,
  /-----BEGIN (?:RSA |EC |OPENSSH |)PRIVATE KEY-----/i,
  /\b(?:refresh|access)_token["']?\s*[:=]\s*["'][A-Za-z0-9._-]{20,}/i,
  /ya29\.[A-Za-z0-9_-]+/,
  /AIza[0-9A-Za-z_-]{20,}/,
];

export function scanForbiddenSecrets(paths) {
  const findings = [];
  for (const root of paths) {
    for (const file of walk(root)) {
      if (!/\.(js|css|html|json|xml|txt|map)$/i.test(file)) continue;
      const text = readFileSync(file, 'utf8');
      for (const pattern of SECRET_PATTERNS) {
        if (pattern.test(text)) {
          findings.push({
            file: relative(repoRoot, file),
            pattern: pattern.source,
          });
        }
      }
    }
  }
  return findings;
}

export function artifactPlaceholderFindings(paths, options = {}) {
  const findings = [];
  const artifactPattern = options.allowLocalhost
    ? /REPLACE_|example\.com|00000000-0000-0000-0000-000000000000/i
    : /REPLACE_|example\.com|localhost|00000000-0000-0000-0000-000000000000/i;
  const templatePattern = /\{\{[^}]+\}\}/i;
  for (const root of paths) {
    for (const file of walk(root)) {
      if (!/\.(js|css|html|json|xml|txt|md)$/i.test(file)) continue;
      const text = readFileSync(file, 'utf8');
      const canContainRuntimeTemplates = /\.(js|css|html)$/i.test(file);
      if (
        artifactPattern.test(text) ||
        (!canContainRuntimeTemplates && templatePattern.test(text))
      ) {
        findings.push(relative(repoRoot, file));
      }
    }
  }
  return findings;
}

export function writeChecksums(files, outPath) {
  const lines = files.map((file) => `${sha256File(file)}  ${relative(repoRoot, file)}`);
  ensureDir(dirname(outPath));
  writeFileSync(outPath, `${lines.join('\n')}\n`);
}

export function verifyChecksums(path) {
  const failures = [];
  if (!existsSync(path)) return [{ file: path, error: 'missing checksum file' }];
  const lines = readFileSync(path, 'utf8').split(/\r?\n/).filter(Boolean);
  for (const line of lines) {
    const match = /^([a-f0-9]{64})  (.+)$/.exec(line);
    if (!match) {
      failures.push({ file: path, error: `malformed checksum line: ${line}` });
      continue;
    }
    const file = join(repoRoot, match[2]);
    if (!existsSync(file)) failures.push({ file: match[2], error: 'missing file' });
    else if (sha256File(file) !== match[1])
      failures.push({ file: match[2], error: 'hash mismatch' });
  }
  return failures;
}

export function createZip(files, root, outPath) {
  ensureDir(dirname(outPath));
  const local = [];
  const central = [];
  let offset = 0;
  const dosTime = 0;
  const dosDate = 33; // 1980-01-01, the earliest valid DOS ZIP date.
  for (const file of files) {
    const name = relative(root, file).replace(/\\/g, '/');
    const data = readFileSync(file);
    const nameBuf = Buffer.from(name);
    const crc = crc32(data);
    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(dosTime, 10);
    localHeader.writeUInt16LE(dosDate, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(data.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(nameBuf.length, 26);
    localHeader.writeUInt16LE(0, 28);
    local.push(localHeader, nameBuf, data);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(dosTime, 12);
    centralHeader.writeUInt16LE(dosDate, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(data.length, 20);
    centralHeader.writeUInt32LE(data.length, 24);
    centralHeader.writeUInt16LE(nameBuf.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    central.push(centralHeader, nameBuf);
    offset += localHeader.length + nameBuf.length + data.length;
  }
  const centralSize = central.reduce((sum, b) => sum + b.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);
  writeFileSync(outPath, Buffer.concat([...local, ...central, end]));
}

let crcTable;
function crc32(buffer) {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[i] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (const byte of buffer) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

export function latestFileBySuffix(dir, suffix) {
  return walk(dir)
    .filter((file) => basename(file).endsWith(suffix))
    .sort()[0];
}
