#!/usr/bin/env bun
import { Command } from "commander";
import { confirm, input, select } from "@inquirer/prompts";
import pc from "picocolors";
import { spawn, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { releaseConfig as resolveReleaseConfig } from "../tools/release/common.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const webEnvPath = path.join(rootDir, "packages/web-shell/.env");
const sideloadStateDir = path.join(rootDir, ".ge-sideload");
const sideloadStatePath = path.join(sideloadStateDir, "unified-sideload.json");

// bootstrap deploy constants (declared here so they exist before the command action runs)
const DEPLOY_STATE = path.join(rootDir, ".ge-deploy.json");
const GCLOUD_ENV = () => ({
  CLOUDSDK_CONFIG: process.env.CLOUDSDK_CONFIG || path.join(rootDir, ".gcloud"),
});
const BOOTSTRAP_SKILLS = [
  {
    name: "m365-command-planner",
    agentId: "m365-command-planner",
    zip: "skill/m365-command-planner.zip",
    idEnv: "VITE_GE_COMMAND_PLANNER_SKILL",
  },
  {
    name: "m365-surface-commander",
    agentId: "m365-surface-commander",
    zip: "skill/m365-surface-commander.zip",
    idEnv: "VITE_GE_SURFACE_COMMANDER_SKILL",
  },
];
const repoGcloudConfigDir = path.join(rootDir, ".gcloud");

const program = new Command();

program
  .name("ge-setup")
  .description("Guided setup and readiness CLI for the Gemini Enterprise Microsoft 365 add-in")
  .option("--yes", "prefer non-interactive defaults where safe");

program
  .command("guide")
  .description("interactive operator guide")
  .action(async () => runGuide(program.opts()));

program
  .command("doctor")
  .description("check local readiness for development and deployment")
  .option("--json", "print machine-readable JSON")
  .option("--fix", "offer guided setup fixes")
  .action(async (opts) => runDoctor(opts));

program
  .command("prereqs")
  .description("install or repair local prerequisites through the Bun setup CLI")
  .option(
    "--target <target>",
    "all | deps | azure | gcloud | atk | m365cli | powershell | exchange | cloudflared | package, comma-separated",
    "all",
  )
  .option("--dry-run", "print commands without running them")
  .option("--yes", "prefer non-interactive defaults where safe")
  .action(async (opts) => runPrereqs(opts, program.opts()));

program
  .command("login")
  .description("guided login helpers for Azure, ATK/M365, Exchange Online, gcloud WIF, and widget skills")
  .option("--target <target>", "azure | atk | exchange | gcloud | wif-extend | widget")
  .option("--force", "force a fresh interactive login even when cached credentials are still valid")
  .option("--tenant <tenant>", "Microsoft Entra tenant id/domain for ATK M365 login")
  .option("--duration <seconds>", "wif-extend: workforce-pool session duration (max 43200s = 12h)", "43200s")
  .action(async (opts) => runLogin(opts));

program
  .command("package")
  .description("generate, validate, and package manifests")
  .option("--profile <profile>", "development | internal-alpha-word-excel", "development")
  .action(async (opts) => runPackage(opts));

program
  .command("dev-tunnel")
  .description("guided Vite + Cloudflare + Entra redirect flow")
  .option("--port <port>", "dev server port")
  .option("--skip-entra", "do not patch the Entra SPA redirect")
  .option("--keep-stale-redirects", "keep old trycloudflare auth redirects")
  .action(async (opts) => runDevTunnel(opts, program.opts()));

program
  .command("addins [action]")
  .alias("m365-addins")
  .description("guided Microsoft 365 tenant add-in deployment")
  .option("--upn <upn>", "admin UPN for Connect-ExchangeOnline")
  .option("--device", "use device-code auth")
  .option("--backend <backend>", "Auto | ExchangeOnlineManagement | O365CentralizedAddInDeployment")
  .option("--members <members>", "comma-separated users/groups")
  .option("--assignment <assignment>", "members | everyone | upload-only")
  .option("--skip-package", "do not regenerate/package manifests first")
  .action(async (action, opts) => runAddins(action, opts, program.opts()));

program
  .command("catalog [action]")
  .alias("m365-catalog")
  .description("upload/list the unified Microsoft 365 app package with CLI for Microsoft 365")
  .option("--package <path>", "unified M365 package zip")
  .option("--profile <profile>", "development | internal-alpha-word-excel")
  .option("--install-cli", "install @pnp/cli-microsoft365 with Bun when m365 is missing")
  .option("--login", "run m365 login before the action")
  .option("--dry-run", "print commands without running them")
  .action(async (action, opts) => runM365Catalog(action, opts, program.opts()));

program
  .command("sideload [action]")
  .description("install/uninstall the unified Microsoft 365 package for local developer testing with Agents Toolkit")
  .option("--package <path>", "unified M365 package zip")
  .option("--skip-package", "use the existing package without rebuilding")
  .option("--skip-tunnel", "do not refresh the Vite + Cloudflare dev tunnel before install")
  .option("--port <port>", "dev server port when refreshing the tunnel")
  .option("--skip-entra", "when refreshing the tunnel, do not patch the Entra SPA redirect")
  .option("--keep-stale-redirects", "when refreshing the tunnel, keep old trycloudflare auth redirects")
  .option("--login", "force `atk auth login m365` before install")
  .option("--skip-atk-login", "skip ATK auth preflight/login before install")
  .option("--tenant <tenant>", "Microsoft Entra tenant id/domain for ATK M365 login")
  .option("--title-id <id>", "title ID to uninstall, e.g. U_90d141c6-cf4f-40ee-b714-9df9ea593f39")
  .option("--dry-run", "print commands without running them")
  .action(async (action, opts) => runSideload(action, opts, program.opts()));

program
  .command("skills")
  .description("guided Gemini Enterprise skill list/update flow")
  .option("--mode <mode>", "list | dry-run | update | paste-update")
  .action(async (opts) => runSkills(opts, program.opts()));

program
  .command("bootstrap")
  .description(
    "one-shot deploy: prereqs → login → package → skills (public AgentService) → tenant deploy, " +
      "with content-hash versioning (.ge-deploy.json). Idempotent: skips unchanged steps.",
  )
  .option("--profile <profile>", "development | internal-alpha-word-excel", "development")
  .option("--deploy-only", "skip prereqs + login (assume the machine is already set up)")
  .option("--bump", "bump the manifest patch version when the manifest changed")
  .option("--assignment <mode>", "members | everyone | upload-only", "members")
  .option("--deployment-lane <lane>", "auto | xml | catalog | none", "auto")
  .option("--tenant-backend <backend>", "Auto | ExchangeOnlineManagement | O365CentralizedAddInDeployment", "Auto")
  .option("--dev-tunnel", "for development profile, refresh the Vite + Cloudflare tunnel before packaging")
  .option("--port <port>", "dev server port when --dev-tunnel is enabled")
  .option("--skip-entra", "do not sync the Entra SPA auth redirect before tenant/catalog deploy")
  .option("--keep-stale-redirects", "keep old trycloudflare auth redirects when syncing Entra")
  .option("--skip-tenant-deploy", "package and update skills, but do not run Microsoft 365 tenant deployment")
  .option("--force-tenant-deploy", "try live Microsoft 365 tenant deployment even when the local preflight says it is unsupported")
  .option("--dev", "print dev-server + tunnel commands after deploy")
  .option("--force", "redeploy every step regardless of the state diff")
  .option("--dry-run", "print the plan and run nothing destructive")
  .action(async (opts) => runBootstrap(opts, program.opts()));

function readEnvFile(filePath) {
  const result = {};
  if (!fs.existsSync(filePath)) return result;
  for (const line of fs.readFileSync(filePath, "utf8").split(/\r?\n/)) {
    const match = line.match(/^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)\s*$/);
    if (!match) continue;
    result[match[1]] = match[2].trim().replace(/^['"]|['"]$/g, "");
  }
  return result;
}

function envValue(name, fallback = "") {
  const envFile = readEnvFile(webEnvPath);
  return process.env[name] || envFile[name] || fallback;
}

function exists(relPath) {
  return fs.existsSync(path.join(rootDir, relPath));
}

function rootPackageVersion() {
  const pkg = safeJson(fs.readFileSync(path.join(rootDir, "package.json"), "utf8"));
  return pkg?.version || "0.0.0";
}

function safeJson(text) {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

function runSync(cmd, args = [], options = {}) {
  return spawnSync(cmd, args, {
    cwd: rootDir,
    encoding: "utf8",
    timeout: options.timeout ?? 8000,
    env: { ...process.env, ...(options.env ?? {}) },
  });
}

function repoWifTokenStatus() {
  const gcloud = firstExecutable(["gcloud"]);
  if (!gcloud) return { ok: false, detail: "gcloud not found" };
  const result = runSync(gcloud, ["auth", "print-access-token"], {
    env: { CLOUDSDK_CONFIG: repoGcloudConfigDir },
    timeout: 20000,
  });
  if (result.status === 0 && String(result.stdout || "").trim()) {
    return { ok: true, detail: "access token minted from cached repo-local WIF credentials" };
  }
  const detail = String(result.stderr || result.stdout || "cached WIF token is unavailable")
    .trim()
    .split(/\r?\n/)
    .slice(-2)
    .join(" ");
  return { ok: false, detail };
}

function atkAuthStatus() {
  const atk = firstExecutable(["atk"]);
  if (!atk) return { ok: false, detail: "Agents Toolkit CLI not found" };
  const result = runSync(atk, ["auth", "list", "--telemetry", "false"], { timeout: 20000 });
  const output = String(result.stdout || result.stderr || "").trim();
  if (result.status === 0 && /Microsoft 365 account is:/i.test(output)) {
    return {
      ok: true,
      detail: output.replace(/\s+/g, " "),
    };
  }
  return {
    ok: false,
    detail: output.replace(/\s+/g, " ") || "no connected Microsoft 365 account",
  };
}

/** Parse the workforce pool {location, pool} from the WIF login config's audience. */
function wifPoolRef() {
  const cfgPath = path.join(repoGcloudConfigDir, "saib-wif-login-config.json");
  const cfg = safeJson(fs.existsSync(cfgPath) ? fs.readFileSync(cfgPath, "utf8") : "") || {};
  const m = String(cfg.audience || "").match(/locations\/([^/]+)\/workforcePools\/([^/]+)/);
  if (!m) throw new Error(`Could not parse a workforce pool from ${cfgPath} (audience missing?).`);
  return { location: m[1], pool: m[2] };
}

function commandExists(cmd) {
  const result = runSync("bash", ["-lc", `command -v ${shellQuote(cmd)}`], { timeout: 3000 });
  return result.status === 0 ? result.stdout.trim() : "";
}

function m365TenantDeploySupport(requestedBackend = "Auto") {
  const pwsh = firstExecutable(["pwsh", "powershell"]);
  if (!pwsh) {
    return {
      ok: false,
      reason: "PowerShell is not installed",
    };
  }

  const exchangeHasCmdlets = runSync(pwsh, [
    "-NoProfile",
    "-Command",
    [
      "Import-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue",
      "if (Get-Command New-OrganizationAddIn -ErrorAction SilentlyContinue) { 'yes' }",
    ].join("; "),
  ], { timeout: 8000 });
  if (exchangeHasCmdlets.stdout.includes("yes")) {
    return {
      ok: true,
      backend: "ExchangeOnlineManagement",
      reason: "ExchangeOnlineManagement exposes New-OrganizationAddIn",
    };
  }

  if (requestedBackend === "ExchangeOnlineManagement") {
    return {
      ok: false,
      backend: "ExchangeOnlineManagement",
      reason: "ExchangeOnlineManagement is installed/importable but does not expose New-OrganizationAddIn in this session",
    };
  }

  if (process.platform !== "win32") {
    const o365HasCmdlets = runSync(pwsh, [
      "-NoProfile",
      "-Command",
      [
        "Import-Module O365CentralizedAddInDeployment -Force -ErrorAction SilentlyContinue",
        "if (Get-Command New-OrganizationAddIn -ErrorAction SilentlyContinue) { 'yes' }",
      ].join("; "),
    ], { timeout: 8000 });
    if (o365HasCmdlets.stdout.includes("yes")) {
      return {
        ok: false,
        backend: "O365CentralizedAddInDeployment",
        reason:
          "Linux has New-OrganizationAddIn through O365CentralizedAddInDeployment, but live auth loads Windows-native kernel32.dll",
      };
    }
  }

  return {
    ok: process.platform === "win32",
    backend: "Auto",
    reason:
      process.platform === "win32"
        ? "Windows PowerShell can use the centralized deployment backend"
        : "no Linux-compatible live Centralized Deployment backend exposes New-OrganizationAddIn",
  };
}

function tenantDeployHandoff(action, assignment, backend = "Auto") {
  const admin = process.env.M365_ADDIN_UPN || "admin@tenant.com";
  const member = process.env.M365_ADDIN_MEMBERS || process.env.M365_ADDIN_UPN || "user@tenant.com";
  const parts = [
    ".\\scripts\\m365-tenant-addin.ps1",
    "-Action",
    action === "update" ? "Update" : "Deploy",
    "-Backend",
    backend,
    "-InstallModule",
    "-UserPrincipalName",
    shellQuote(admin),
  ];
  if (assignment === "everyone") {
    parts.push("-AssignToEveryone");
  } else if (assignment === "upload-only") {
    parts.push("-UploadOnly");
  } else {
    parts.push("-Members", shellQuote(member));
  }
  return parts.join(" ");
}

function firstExecutable(candidates) {
  for (const candidate of candidates) {
    if (!candidate) continue;
    if (candidate.includes("/")) {
      const abs = path.isAbsolute(candidate) ? candidate : path.join(rootDir, candidate);
      if (fs.existsSync(abs)) return abs;
      continue;
    }
    const found = commandExists(candidate);
    if (found) return found;
  }
  return "";
}

function shellQuote(value) {
  return `'${String(value).replaceAll("'", "'\\''")}'`;
}

// Ramsian CLI house style: quiet structure, sparse accent color, aligned facts.
const STATUS = {
  ok: { symbol: "✓", paint: pc.green },
  warn: { symbol: "⚠", paint: pc.yellow },
  missing: { symbol: "✗", paint: pc.red },
};
const PROMPT = "$";

function accent(value) {
  return pc.cyan(value);
}

function muted(value) {
  return pc.dim(value);
}

function section(title) {
  console.log("");
  console.log(`${accent("›")} ${pc.bold(title)}`);
}

function keyval(key, value, pad = 18) {
  console.log(`  ${muted((key + ":").padEnd(pad))} ${value}`);
}

function note(text) {
  console.log(`  ${muted(text)}`);
}

function statusLine(level, text, detail = "") {
  const status = STATUS[level] ?? STATUS.warn;
  console.log(`  ${status.paint(status.symbol)} ${text}${detail ? `  ${muted(detail)}` : ""}`);
}

function commandLabel(cmd, args = [], options = {}) {
  if (options.label) return options.label;
  const rendered = [cmd, ...args].map(shellQuote).join(" ");
  if (process.env.GE_SETUP_VERBOSE === "1" || rendered.length <= 140) return rendered;
  const shownArgs = args.slice(0, 4).map(shellQuote).join(" ");
  return `${shellQuote(cmd)}${shownArgs ? ` ${shownArgs}` : ""} ${muted(`… ${args.length} args`)}`;
}

function printCommand(cmd, args = [], options = {}) {
  console.log(muted(`${PROMPT} ${commandLabel(cmd, args, options)}`));
  if (options.label && process.env.GE_SETUP_VERBOSE === "1") {
    note([cmd, ...args].map(shellQuote).join(" "));
  }
}

function printCheckGroup(area, areaChecks) {
  const width = Math.max(...areaChecks.map((check) => check.name.length), 0);
  console.log(`\n${accent(area.toLowerCase())}`);
  for (const check of areaChecks) {
    const status = STATUS[check.level] ?? STATUS.warn;
    const name = check.name.padEnd(width);
    const detail = check.detail ? muted(check.detail) : "";
    console.log(`  ${status.paint(status.symbol)} ${name}  ${detail}`);
  }
}

async function run(cmd, args = [], options = {}) {
  printCommand(cmd, args, options);
  await new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      cwd: rootDir,
      stdio: "inherit",
      env: { ...process.env, ...(options.env ?? {}) },
    });
    child.on("exit", (code) => (code === 0 ? resolve() : reject(new Error(`${cmd} exited with ${code}`))));
    child.on("error", reject);
  });
}

async function maybeInput(message, defaultValue, yes) {
  if (yes || !process.stdin.isTTY) return defaultValue;
  return input({ message, default: defaultValue });
}

async function maybeConfirm(message, defaultValue, yes) {
  if (yes || !process.stdin.isTTY) return defaultValue;
  return confirm({ message, default: defaultValue });
}

async function maybeSelect(message, choices, defaultValue, yes) {
  if (yes || !process.stdin.isTTY) return defaultValue;
  return select({ message, choices, default: defaultValue });
}

function platform() {
  if (process.platform === "linux") {
    const osRelease = fs.existsSync("/etc/os-release")
      ? fs.readFileSync("/etc/os-release", "utf8")
      : "";
    const id = /^ID=(.+)$/m.exec(osRelease)?.[1]?.replaceAll('"', "") ?? "";
    const versionId = /^VERSION_ID=(.+)$/m.exec(osRelease)?.[1]?.replaceAll('"', "") ?? "";
    return { os: process.platform, id, versionId };
  }
  return { os: process.platform, id: "", versionId: "" };
}

async function runMaybe(cmd, args = [], options = {}) {
  if (options.dryRun) {
    printCommand(cmd, args, options);
    console.log(`  ${pc.yellow("dry-run")} ${muted("not executed")}`);
    return;
  }
  await run(cmd, args, options);
}

async function runCapture(cmd, args = [], options = {}) {
  if (options.dryRun) {
    printCommand(cmd, args, options);
    console.log(`  ${pc.yellow("dry-run")} ${muted("not executed")}`);
    return "";
  }
  printCommand(cmd, args, options);
  return await new Promise((resolve, reject) => {
    let output = "";
    const child = spawn(cmd, args, {
      cwd: rootDir,
      stdio: ["ignore", "pipe", "pipe"],
      env: { ...process.env, ...(options.env ?? {}) },
    });
    child.stdout.on("data", (chunk) => {
      const text = String(chunk);
      output += text;
      process.stdout.write(text);
    });
    child.stderr.on("data", (chunk) => {
      const text = String(chunk);
      output += text;
      process.stderr.write(text);
    });
    child.on("exit", (code) => (code === 0 ? resolve(output) : reject(new Error(`${cmd} exited with ${code}`))));
    child.on("error", reject);
  });
}

function targetSet(raw = "all") {
  const parts = String(raw)
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
  return new Set(parts.length ? parts : ["all"]);
}

function includesTarget(targets, name) {
  return targets.has("all") || targets.has(name);
}

async function runGuide(globalOpts) {
  for (;;) {
    section("GE M365 setup guide");
    const choice = await select({
      message: "What do you want to do?",
      default: "doctor",
      choices: [
        { name: "Run readiness doctor", value: "doctor" },
        { name: "Install/repair prerequisites", value: "prereqs" },
        { name: "Run guided login", value: "login" },
        { name: "Generate/validate/package manifests", value: "package" },
        { name: "Restart dev server + Cloudflare + Entra redirect", value: "tunnel" },
        { name: "Deploy/list Microsoft 365 tenant add-ins", value: "addins" },
        { name: "Upload/list unified M365 app package", value: "catalog" },
        { name: "Sideload/uninstall unified package locally", value: "sideload" },
        { name: "List/update Gemini Enterprise skills", value: "skills" },
        { name: "Quit", value: "quit" },
      ],
    });
    if (choice === "quit") return;
    if (choice === "doctor") await runDoctor({});
    if (choice === "prereqs") await runPrereqs({}, globalOpts);
    if (choice === "login") await runLogin({});
    if (choice === "package") await runPackage({ profile: "development" });
    if (choice === "tunnel") await runDevTunnel({}, globalOpts);
    if (choice === "addins") await runAddins(undefined, {}, globalOpts);
    if (choice === "catalog") await runM365Catalog(undefined, {}, globalOpts);
    if (choice === "sideload") await runSideload(undefined, {}, globalOpts);
    if (choice === "skills") await runSkills({}, globalOpts);
  }
}

async function runDoctor(opts = {}) {
  const checks = collectChecks();
  if (opts.json) {
    console.log(JSON.stringify(checks, null, 2));
    return;
  }
  section("Readiness doctor");
  for (const [area, areaChecks] of Object.entries(groupBy(checks, (check) => check.area))) {
    printCheckGroup(area, areaChecks);
  }
  const missing = checks.filter((check) => check.level === "missing");
  const warn = checks.filter((check) => check.level === "warn");
  const okCount = checks.length - missing.length - warn.length;
  console.log("");
  console.log(
    `${pc.bold("Summary")}  ` +
      `${STATUS.ok.paint(`${okCount} ok`)} · ` +
      `${warn.length ? STATUS.warn.paint(`${warn.length} warn`) : pc.dim("0 warn")} · ` +
      `${missing.length ? STATUS.missing.paint(`${missing.length} missing`) : pc.dim("0 missing")}`,
  );

  if (opts.fix) {
    if (checks.some((check) => check.name === "PowerShell" && check.level === "missing") || checks.some((check) => check.name === "Centralized add-in cmdlets" && check.level !== "ok")) {
      if (await maybeConfirm("Install/repair PowerShell + Microsoft 365 add-in modules now?", true, Boolean(program.opts().yes))) {
        await runPrereqs({ target: "powershell,exchange" }, program.opts());
      }
    }
    if (checks.some((check) => check.area === "Artifacts" && check.level !== "ok")) {
      if (await maybeConfirm("Generate and package development manifests now?", true, Boolean(program.opts().yes))) {
        await runPackage({ profile: "development" });
      }
    }
  }
}

function collectChecks() {
  const envFile = readEnvFile(webEnvPath);
  const nodeMajor = Number(process.versions.node.split(".")[0]);
  const bun = runSync("bun", ["--version"], { timeout: 3000 });
  const az = firstExecutable(["./bin/az", ".venv-az/bin/az", "az"]);
  const cloudflared = firstExecutable([process.env.CLOUDFLARED_BIN, "/tmp/cloudflared", "cloudflared"]);
  const pwsh = firstExecutable(["pwsh", "powershell"]);
  const gcloud = firstExecutable(["gcloud"]);
  const atk = firstExecutable(["atk"]);
  const m365 = firstExecutable(["m365"]);
  const requiredEnv = [
    "VITE_GCP_PROJECT",
    "VITE_GCP_LOCATION",
    "VITE_GE_ENGINE",
    "VITE_GE_COLLECTION",
    "VITE_GE_ASSISTANT",
    "VITE_ENTRA_CLIENT_ID",
    "VITE_ENTRA_AUTHORITY",
  ];
  const missingEnv = requiredEnv.filter((key) => !envFile[key] && !process.env[key]);
  const checks = [];

  checks.push(check("Core", "Bun", bun.status === 0 ? "ok" : "missing", bun.status === 0 ? bun.stdout.trim() : "install Bun"));
  checks.push(check("Core", "Node >= 20", nodeMajor >= 20 ? "ok" : "missing", process.version));
  checks.push(check("Core", "node_modules", exists("node_modules") ? "ok" : "warn", exists("node_modules") ? "installed" : "run bun install"));
  checks.push(check("Config", "packages/web-shell/.env", fs.existsSync(webEnvPath) ? "ok" : "missing", fs.existsSync(webEnvPath) ? "present" : "copy packages/web-shell/.env.example"));
  checks.push(check("Config", "required web-shell env", missingEnv.length === 0 ? "ok" : "warn", missingEnv.length === 0 ? "present" : missingEnv.join(", ")));
  checks.push(check("Microsoft", "Azure CLI", az ? "ok" : "warn", az || "run bun run setup:azure before syncing Entra redirects"));
  if (az) {
    const account = runSync(az, ["account", "show"], { timeout: 8000 });
    checks.push(check("Microsoft", "Azure CLI signed in", account.status === 0 ? "ok" : "warn", account.status === 0 ? "signed in" : "run setup:login -> Azure"));
  }
  checks.push(check("Microsoft", "PowerShell", pwsh ? "ok" : "missing", pwsh || "run bun run setup:powershell"));
  if (pwsh) {
    const addInCmdlets = runSync(pwsh, [
      "-NoProfile",
      "-Command",
      [
        "$mods = @(Get-Module -ListAvailable O365CentralizedAddInDeployment,ExchangeOnlineManagement | Select-Object -ExpandProperty Name -Unique)",
        "if ($mods -contains 'O365CentralizedAddInDeployment') { Import-Module O365CentralizedAddInDeployment -Force -ErrorAction SilentlyContinue }",
        "if (Get-Command New-OrganizationAddIn -ErrorAction SilentlyContinue) { 'yes' } elseif ($mods.Count -gt 0) { 'module-without-cmdlets' }",
      ].join("; "),
    ], { timeout: 8000 });
    checks.push(check(
      "Microsoft",
      "Centralized add-in cmdlets",
      addInCmdlets.stdout.includes("yes") ? "ok" : "warn",
      addInCmdlets.stdout.includes("yes") ? "available" : "run bun run setup:powershell",
    ));
  }
  checks.push(check("Microsoft", "Agents Toolkit CLI", atk ? "ok" : "warn", atk || "run bun run setup:atk for unified package dev install"));
  checks.push(check("Microsoft", "CLI for Microsoft 365", m365 ? "ok" : "warn", m365 || "run bun run setup:prereqs -- --target m365cli"));
  checks.push(check("Google/Gemini", "gcloud", gcloud ? "ok" : "warn", gcloud || "run bun run setup:gcloud for WIF CLI checks"));
  checks.push(check("Google/Gemini", "WIF login config", exists(".gcloud/saib-wif-login-config.json") ? "ok" : "warn", exists(".gcloud/saib-wif-login-config.json") ? ".gcloud/saib-wif-login-config.json" : "ask admin for login config"));
  if (gcloud) {
    const localConfig = runSync(gcloud, ["config", "list", "--format=json"], {
      timeout: 8000,
      env: { CLOUDSDK_CONFIG: repoGcloudConfigDir },
    });
    const localConfigJson = localConfig.status === 0 ? safeJson(localConfig.stdout) : undefined;
    const localAccount = localConfigJson?.core?.account;
    const localProject = localConfigJson?.core?.project;
    const localQuota = localConfigJson?.billing?.quota_project;
    checks.push(check(
      "Google/Gemini",
      "repo-local WIF gcloud config",
      localAccount && localProject ? "ok" : "warn",
      localAccount && localProject
        ? `${localAccount} / ${localProject}${localQuota ? ` quota=${localQuota}` : ""}`
        : "run bun run setup:gcloud:wif; this uses .gcloud and does not update ADC",
    ));
  }
  const adcPath = path.join(os.homedir(), ".config/gcloud/application_default_credentials.json");
  if (fs.existsSync(adcPath)) {
    const adc = safeJson(fs.readFileSync(adcPath, "utf8"));
    const adcType = adc?.type ?? "unknown";
    const adcQuota = adc?.quota_project_id ?? "no quota project";
    const looksLikeRepoWif = adcType.includes("external_account") && adcQuota === envValue("VITE_GCP_PROJECT", "saib-ai-playground");
    checks.push(check(
      "Google/Gemini",
      "global ADC isolation",
      looksLikeRepoWif ? "warn" : "ok",
      looksLikeRepoWif
        ? `global ADC is ${adcType} for ${adcQuota}; WIF should stay repo-local, move this ADC file aside`
        : `global ADC present: ${adcType} / ${adcQuota}`,
    ));
  } else {
    checks.push(check("Google/Gemini", "global ADC isolation", "ok", "no local ADC file; metadata/default service account can be used"));
  }
  const widgetConfigured = Boolean(envValue("VITE_GE_WIDGET_CONFIG_ID") || process.env.GE_WIDGET_CONFIG_ID);
  checks.push(check("Google/Gemini", "widget config", widgetConfigured ? "ok" : "warn", widgetConfigured ? "configured" : "needed for widget skill flows"));
  checks.push(check("Dev server", "cloudflared", cloudflared ? "ok" : "warn", cloudflared || "needed for remote Office web dev"));
  checks.push(check("Artifacts", "development package", exists("dist/release/development-m365-v0.1.0.zip") ? "ok" : "warn", exists("dist/release/development-m365-v0.1.0.zip") ? "present" : "run bun run setup:package"));
  checks.push(check("Artifacts", "XML manifests", exists("dist/package/development/xml/excel.manifest.xml") ? "ok" : "warn", exists("dist/package/development/xml/excel.manifest.xml") ? "present" : "run bun run setup:package"));
  checks.push(
    check(
      "Skills",
      "skill zips",
      exists("skill/m365-command-planner.zip") && exists("skill/m365-surface-commander.zip")
        ? "ok"
        : "warn",
      "run bun run ge:skills to rebuild/upload",
    ),
  );
  return checks;
}

function check(area, name, level, detail) {
  return { area, name, level, detail };
}

function groupBy(items, keyFn) {
  return items.reduce((acc, item) => {
    const key = keyFn(item);
    acc[key] ??= [];
    acc[key].push(item);
    return acc;
  }, {});
}

async function runPrereqs(opts = {}, globalOpts = {}) {
  const yes = Boolean(globalOpts.yes || opts.yes);
  const dryRun = Boolean(opts.dryRun);
  const targets = targetSet(opts.target);

  section("Prerequisite setup");
  note("Bun orchestrates the flow; system tools still use their native installers.");

  if (includesTarget(targets, "deps")) {
    if (await maybeConfirm("Install/update Bun workspace dependencies?", true, yes)) {
      await runMaybe("bun", ["install"], { dryRun });
    }
  }

  if (includesTarget(targets, "azure")) {
    await ensureAzureCli({ yes, dryRun });
  }

  if (includesTarget(targets, "gcloud")) {
    await ensureGcloud({ yes, dryRun });
  }

  if (includesTarget(targets, "atk")) {
    await ensureAgentsToolkit({ yes, dryRun });
  }

  if (includesTarget(targets, "m365cli")) {
    await ensureM365Cli({ yes, dryRun });
  }

  if (includesTarget(targets, "powershell")) {
    await ensurePowerShell({ yes, dryRun });
  }

  if (includesTarget(targets, "exchange")) {
    await ensureExchangeOnlineManagement({ yes, dryRun });
  }

  if (includesTarget(targets, "cloudflared")) {
    await ensureCloudflared({ yes, dryRun });
  }

  if (includesTarget(targets, "package")) {
    if (await maybeConfirm("Generate, validate, and package development manifests?", true, yes)) {
      await runPackage({ profile: "development", dryRun });
    }
  }

  section("Post-setup doctor");
  if (!dryRun) await runDoctor({});
}

async function ensurePowerShell({ yes, dryRun }) {
  const pwsh = firstExecutable(["pwsh", "powershell"]);
  if (pwsh) {
    statusLine("ok", "PowerShell", pwsh);
    return;
  }

  const p = platform();
  if (p.os !== "linux" || p.id !== "ubuntu") {
    statusLine("warn", "PowerShell auto-install skipped", "Ubuntu only; install manually and rerun setup:doctor");
    return;
  }

  if (!(await maybeConfirm(`Install PowerShell Core for Ubuntu ${p.versionId} using sudo apt-get?`, true, yes))) {
    return;
  }
  await runMaybe("scripts/setup-pwsh-ubuntu.sh", [], {
    dryRun,
    env: { M365_ADDIN_INSTALL_MODULE: "0" },
  });
}

async function ensureAzureCli({ yes, dryRun }) {
  const current = firstExecutable(["./bin/az", ".venv-az/bin/az", "az"]);
  if (current) {
    statusLine("ok", "Azure CLI", current);
    return;
  }
  const uv = firstExecutable(["uv"]);
  if (!uv) {
    statusLine("warn", "Azure CLI install skipped", "uv is required for the repo-local install path");
    return;
  }
  if (!(await maybeConfirm("Install Azure CLI into repo-local .venv-az and create bin/az?", true, yes))) {
    return;
  }
  await runMaybe(uv, ["venv", ".venv-az"], { dryRun });
  await runMaybe(uv, ["pip", "install", "--python", ".venv-az/bin/python", "azure-cli"], { dryRun });
  const wrapper = `#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "\${BASH_SOURCE[0]}")/.." && pwd)"
export AZURE_CONFIG_DIR="\${AZURE_CONFIG_DIR:-$ROOT/.azure}"

exec "$ROOT/.venv-az/bin/python" -m azure.cli "$@"
`;
  if (dryRun) {
    console.log(muted(`${PROMPT} write bin/az wrapper`) + pc.yellow("  dry-run"));
    console.log(muted(`${PROMPT} chmod +x bin/az`) + pc.yellow("  dry-run"));
    return;
  }
  fs.mkdirSync(path.join(rootDir, "bin"), { recursive: true });
  fs.writeFileSync(path.join(rootDir, "bin/az"), wrapper);
  fs.chmodSync(path.join(rootDir, "bin/az"), 0o755);
}

async function ensureGcloud({ yes, dryRun }) {
  const current = firstExecutable(["gcloud"]);
  if (current) {
    statusLine("ok", "gcloud", current);
    return;
  }
  const p = platform();
  if (p.os !== "linux" || p.id !== "ubuntu") {
    statusLine("warn", "gcloud auto-install skipped", "Ubuntu only; install manually and rerun setup:doctor");
    return;
  }
  if (!(await maybeConfirm("Install Google Cloud CLI using the official Ubuntu apt repository?", true, yes))) {
    return;
  }
  await runMaybe("sudo", ["apt-get", "update"], { dryRun });
  await runMaybe("sudo", ["apt-get", "install", "-y", "apt-transport-https", "ca-certificates", "gnupg", "curl"], { dryRun });
  await runMaybe("bash", [
    "-lc",
    "curl -fsSL https://packages.cloud.google.com/apt/doc/apt-key.gpg | sudo gpg --dearmor -o /usr/share/keyrings/cloud.google.gpg",
  ], { dryRun });
  await runMaybe("bash", [
    "-lc",
    "echo 'deb [signed-by=/usr/share/keyrings/cloud.google.gpg] https://packages.cloud.google.com/apt cloud-sdk main' | sudo tee /etc/apt/sources.list.d/google-cloud-sdk.list >/dev/null",
  ], { dryRun });
  await runMaybe("sudo", ["apt-get", "update"], { dryRun });
  await runMaybe("sudo", ["apt-get", "install", "-y", "google-cloud-cli"], { dryRun });
}

async function ensureAgentsToolkit({ yes, dryRun }) {
  const current = firstExecutable(["atk"]);
  if (current) {
    statusLine("ok", "Agents Toolkit CLI", current);
    return;
  }
  if (!(await maybeConfirm("Install Microsoft 365 Agents Toolkit CLI globally with Bun?", true, yes))) {
    return;
  }
  await runMaybe("bun", ["add", "-g", "@microsoft/m365agentstoolkit-cli"], { dryRun });
}

async function ensureM365Cli({ yes, dryRun }) {
  const current = firstExecutable(["m365"]);
  if (current) {
    statusLine("ok", "CLI for Microsoft 365", current);
    return;
  }
  if (!(await maybeConfirm("Install CLI for Microsoft 365 globally with Bun?", true, yes))) {
    return;
  }
  await runMaybe("bun", ["add", "-g", "@pnp/cli-microsoft365"], { dryRun });
}

async function loginAgentsToolkitM365(opts = {}, globalOpts = {}) {
  const yes = Boolean(globalOpts.yes);
  if (!firstExecutable(["atk"]) && !opts.dryRun) {
    await ensureAgentsToolkit({ yes, dryRun: false });
  }
  const atk = firstExecutable(["atk"]) || "atk";
  const force = Boolean(opts.force);
  if (!force && !opts.dryRun) {
    const current = atkAuthStatus();
    if (current.ok) {
      statusLine("ok", "ATK Microsoft 365 auth", current.detail);
      return;
    }
    statusLine("warn", "ATK Microsoft 365 auth", current.detail);
  }
  const tenant = opts.tenant ?? envValue("VITE_ENTRA_TENANT_ID") ?? "";
  const args = ["auth", "login", "m365"];
  if (tenant) args.push("--tenant", tenant);
  args.push("--telemetry", "false");
  await runMaybe(atk, args, {
    dryRun: Boolean(opts.dryRun),
    label: "atk auth login m365",
  });
}

async function ensureExchangeOnlineManagement({ yes, dryRun }) {
  const pwsh = firstExecutable(["pwsh", "powershell"]);
  if (!pwsh && !dryRun) {
    await ensurePowerShell({ yes, dryRun });
  }
  const resolvedPwsh = firstExecutable(["pwsh", "powershell"]);
  if (!resolvedPwsh && !dryRun) {
    statusLine("warn", "Microsoft 365 add-in modules skipped", "PowerShell is unavailable");
    return;
  }
  const checkModule = resolvedPwsh
    ? runSync(resolvedPwsh, [
        "-NoProfile",
        "-Command",
        [
          "$mods = @(Get-Module -ListAvailable O365CentralizedAddInDeployment,ExchangeOnlineManagement | Select-Object -ExpandProperty Name -Unique)",
          "if ($mods -contains 'O365CentralizedAddInDeployment') { Import-Module O365CentralizedAddInDeployment -Force -ErrorAction SilentlyContinue }",
          "if (Get-Command New-OrganizationAddIn -ErrorAction SilentlyContinue) { 'yes' }",
        ].join("; "),
      ])
    : { stdout: "" };
  if (checkModule.stdout.includes("yes")) {
    statusLine("ok", "Microsoft 365 add-in cmdlets", "available");
    return;
  }
  if (!(await maybeConfirm("Install Microsoft 365 add-in PowerShell modules for the current user?", true, yes))) {
    return;
  }
  await runMaybe(resolvedPwsh || "pwsh", [
    "-NoProfile",
    "-Command",
    "'ExchangeOnlineManagement','O365CentralizedAddInDeployment' | ForEach-Object { Install-Module -Name $_ -Scope CurrentUser -Force }",
  ], { dryRun });
}

async function ensureCloudflared({ yes, dryRun }) {
  const current = firstExecutable([process.env.CLOUDFLARED_BIN, "/tmp/cloudflared", "cloudflared"]);
  if (current) {
    statusLine("ok", "cloudflared", current);
    return;
  }
  if (!(await maybeConfirm("Download cloudflared Linux AMD64 binary to /tmp/cloudflared?", true, yes))) {
    return;
  }
  await runMaybe("curl", [
    "-L",
    "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64",
    "-o",
    "/tmp/cloudflared",
  ], { dryRun });
  await runMaybe("chmod", ["+x", "/tmp/cloudflared"], { dryRun });
}

async function runLogin(opts = {}) {
  const target = opts.target ?? await maybeSelect("Which login/check?", [
    { name: "Azure CLI / Entra", value: "azure" },
    { name: "Agents Toolkit / Microsoft 365", value: "atk" },
    { name: "Exchange Online add-in cmdlet connection", value: "exchange" },
    { name: "Google WIF CLI", value: "gcloud" },
    { name: "Extend WIF session duration (fewer sign-ins)", value: "wif-extend" },
    { name: "Gemini widget skill credential flow", value: "widget" },
  ], "azure", false);

  if (target === "azure") {
    const az = firstExecutable(["./bin/az", ".venv-az/bin/az", "az"]);
    if (!az) throw new Error("Azure CLI not found. Install az or set up ./bin/az.");
    const tenant = await input({ message: "Tenant ID/domain", default: envValue("VITE_ENTRA_TENANT_ID") || envValue("VITE_ENTRA_AUTHORITY").split("/").pop() || "" });
    await run(az, ["login", "--use-device-code", ...(tenant ? ["--tenant", tenant] : [])]);
  }

  if (target === "atk") {
    section("Agents Toolkit Microsoft 365 login");
    await loginAgentsToolkitM365(opts);
  }

  if (target === "exchange") {
    const upn = await input({ message: "Admin UPN for Connect-ExchangeOnline", default: process.env.M365_ADDIN_UPN || "vamramak@psott.onmicrosoft.com" });
    await run("scripts/m365-tenant-addin.sh", ["list"], {
      env: {
        M365_ADDIN_UPN: upn,
        M365_ADDIN_DEVICE: "1",
        M365_ADDIN_INSTALL_MODULE: "1",
        M365_ADDIN_SKIP_PACKAGE: "1",
      },
    });
  }

  if (target === "gcloud") {
    const gcloud = firstExecutable(["gcloud"]);
    if (!gcloud) throw new Error("gcloud not found.");
    fs.mkdirSync(repoGcloudConfigDir, { recursive: true });
    if (!opts.force) {
      const existing = repoWifTokenStatus();
      if (existing.ok) {
        section("Google WIF CLI");
        keyval("CLOUDSDK_CONFIG", repoGcloudConfigDir);
        statusLine("ok", "cached WIF token", existing.detail);
        note("login skipped; use --force when switching accounts or repairing the local WIF session");
        await run(gcloud, ["config", "list"], { env: { CLOUDSDK_CONFIG: repoGcloudConfigDir } });
        return;
      }
      statusLine("warn", "cached WIF token", existing.detail || "not usable; interactive login required");
    }
    const loginConfig = await input({ message: "WIF login config", default: path.join(rootDir, ".gcloud/saib-wif-login-config.json") });
    const project = await input({ message: "WIF project/quota project", default: envValue("VITE_GCP_PROJECT", "saib-ai-playground") });
    console.log("");
    keyval("CLOUDSDK_CONFIG", repoGcloudConfigDir);
    statusLine("ok", "ADC isolation", "no application-default login; global ADC is left untouched");
    await run(gcloud, ["auth", "login", "--login-config", loginConfig, "--no-launch-browser"], {
      env: { CLOUDSDK_CONFIG: repoGcloudConfigDir },
    });
    if (project) {
      await run(gcloud, ["config", "set", "project", project], {
        env: { CLOUDSDK_CONFIG: repoGcloudConfigDir },
      });
      await run(gcloud, ["config", "set", "billing/quota_project", project], {
        env: { CLOUDSDK_CONFIG: repoGcloudConfigDir },
      });
    }
    await run(gcloud, ["config", "list"], { env: { CLOUDSDK_CONFIG: repoGcloudConfigDir } });
    note(
      "this sign-in lasts as long as the workforce-pool session duration (default ~1h). To sign in " +
        "far less often, run: bun run setup:gcloud:wif:extend (raises it to the 12h max; needs IAM admin on the pool).",
    );
  }

  if (target === "wif-extend") {
    const gcloud = firstExecutable(["gcloud"]);
    if (!gcloud) throw new Error("gcloud not found.");
    const ref = wifPoolRef();
    const duration = opts.duration || "43200s";
    section("Extend WIF session duration");
    keyval("pool", `${ref.pool}`);
    keyval("location", ref.location);
    keyval("session-duration", duration);
    note(
      "the Entra token itself has a fixed ~1h lifetime (Microsoft no longer lets you extend it), but " +
        "WIF only consumes it once per session — gcloud then refreshes the Google access token silently " +
        "until this pool session duration elapses. Raising it = fewer interactive sign-ins. Needs IAM " +
        "admin (roles/iam.workforcePoolAdmin) on the pool; it changes the pool for everyone who uses it.",
    );
    await run(
      gcloud,
      ["iam", "workforce-pools", "update", ref.pool, "--location", ref.location, "--session-duration", duration],
      { env: { CLOUDSDK_CONFIG: repoGcloudConfigDir } },
    );
    statusLine("ok", "session duration", `set to ${duration} — re-run WIF login once to mint a token with the new lifetime`);
  }

  if (target === "widget") {
    await run("scripts/update-ge-widget-skills.sh", ["--list-only"]);
  }
}

async function runPackage(opts = {}) {
  const profile = opts.profile ?? "development";
  await runMaybe("bun", ["run", "--filter", "@ge/web-shell", "build"], {
    dryRun: Boolean(opts.dryRun),
    label: "build web shell",
  });
  await runMaybe("bun", ["run", "manifests:generate", "--", "--profile", profile], {
    dryRun: Boolean(opts.dryRun),
    label: `generate ${profile} manifests`,
  });
  await runMaybe("bun", ["run", "manifests:validate", "--", "--profile", profile], {
    dryRun: Boolean(opts.dryRun),
    label: `validate ${profile} manifests`,
  });
  if (profile === "alpha" || profile === "internal-alpha-word-excel") {
    await runMaybe("bun", ["run", "package:alpha"], {
      dryRun: Boolean(opts.dryRun),
      label: "package internal alpha release",
    });
  } else {
    await runMaybe("bun", ["run", "package:dev"], {
      dryRun: Boolean(opts.dryRun),
      label: "package development release",
    });
  }
}

async function runDevTunnel(opts = {}, globalOpts = {}) {
  const yes = Boolean(globalOpts.yes);
  await ensureCloudflared({ yes, dryRun: false });
  const scriptArgs = [];
  const port = opts.port ?? await maybeInput("Dev port", envValue("GE_DEV_PORT", "13000"), yes);
  if (port) scriptArgs.push("--port", port);
  if (opts.skipEntra || !(await maybeConfirm("Patch Entra SPA redirect?", true, yes))) scriptArgs.push("--skip-entra");
  if (opts.keepStaleRedirects || await maybeConfirm("Keep old trycloudflare auth redirects?", false, yes)) scriptArgs.push("--keep-stale-redirects");
  await run("scripts/dev-tunnel-entra.sh", scriptArgs);
}

async function runAddins(action, opts = {}, globalOpts = {}) {
  const yes = Boolean(globalOpts.yes);
  const selectedAction = action ?? await maybeSelect("Add-in action", [
    { name: "List deployed add-ins", value: "list" },
    { name: "Deploy generated XML manifests", value: "deploy" },
    { name: "Update deployed manifests", value: "update" },
    { name: "Assign existing add-ins", value: "assign" },
    { name: "Delete generated add-ins", value: "delete" },
  ], "list", yes);
  const env = {};
  const backend = opts.backend ?? process.env.M365_ADDIN_BACKEND ?? "Auto";
  env.M365_ADDIN_BACKEND = backend;
  const upnDefault = process.env.M365_ADDIN_UPN || "vamramak@psott.onmicrosoft.com";
  env.M365_ADDIN_UPN = opts.upn ?? await maybeInput("Admin UPN", upnDefault, yes);
  if (backend === "O365CentralizedAddInDeployment") {
    env.M365_ADDIN_DEVICE = "0";
    note("O365CentralizedAddInDeployment does not support device-code login; it uses Connect-OrganizationAddInService.");
  } else {
    env.M365_ADDIN_DEVICE = opts.device || await maybeConfirm("Use device-code login?", true, yes) ? "1" : "0";
  }
  if (process.platform !== "win32" && backend !== "ExchangeOnlineManagement") {
    note("on Linux, Auto may select O365CentralizedAddInDeployment for the cmdlets, but live auth can fail if that module loads kernel32.dll; dry-run still works.");
  }
  env.M365_ADDIN_INSTALL_MODULE = "1";
  if (opts.skipPackage) env.M365_ADDIN_SKIP_PACKAGE = "1";

  if (["deploy", "update", "assign"].includes(selectedAction)) {
    const assignment = opts.assignment ?? await maybeSelect("Assignment", [
      { name: "Specific users/groups", value: "members" },
      { name: "Entire tenant", value: "everyone" },
      { name: "Upload only", value: "upload-only" },
    ], "members", yes);
    if (assignment === "everyone") env.M365_ADDIN_ASSIGN_EVERYONE = "1";
    if (assignment === "upload-only") env.M365_ADDIN_UPLOAD_ONLY = "1";
    if (assignment === "members") {
      env.M365_ADDIN_MEMBERS = opts.members ?? await maybeInput("Members, comma-separated", process.env.M365_ADDIN_MEMBERS || env.M365_ADDIN_UPN, yes);
    }
  }
  await run("scripts/m365-tenant-addin.sh", [selectedAction], { env });
}

async function runM365Catalog(action, opts = {}, globalOpts = {}) {
  const yes = Boolean(globalOpts.yes);
  const selectedAction = action ?? await maybeSelect("M365 app catalog action", [
    { name: "Show CLI login status", value: "status" },
    { name: "Login with device code", value: "login" },
    { name: "List tenant app catalog apps", value: "list" },
    { name: "Add unified package", value: "add" },
    { name: "Update existing app by manifest id", value: "update" },
    { name: "Upsert unified package", value: "upsert" },
    { name: "Remove app by manifest id", value: "remove" },
  ], "status", yes);
  const env = {};
  if (opts.package) env.M365_APP_PACKAGE = opts.package;
  else if (opts.profile) env.M365_APP_PACKAGE = m365PackageForProfile(opts.profile);
  if (opts.profile) env.M365_APP_PROFILE = opts.profile;
  if (opts.installCli || await maybeConfirm("Install CLI for Microsoft 365 if missing?", true, yes)) {
    env.M365_CLI_INSTALL = "1";
  }
  if (opts.login || (["add", "update", "upsert", "remove", "list"].includes(selectedAction) && await maybeConfirm("Run m365 login first?", false, yes))) {
    env.M365_CLI_LOGIN = "1";
    env.M365_CLI_AUTH_TYPE = "deviceCode";
  }
  const args = [selectedAction];
  if (opts.dryRun) args.push("--dry-run");
  await run("scripts/m365-app-catalog.sh", args, { env });
}

async function runEntraSync(profile, opts = {}) {
  const args = ["tools/release/sync-entra-spa.mjs", "--profile", profile];
  if (opts.keepStaleRedirects) args.push("--keep-stale-redirects");
  if (opts.dryRun) args.push("--dry-run");
  await runMaybe("bun", args, {
    dryRun: false,
    label: `sync ${profile} Entra SPA redirect`,
  });
}

function m365PackageForProfile(profile = "development") {
  if (profile === "development") {
    const artifactPath = path.join(rootDir, "dist/release/development-artifact.json");
    const artifact = safeJson(fs.existsSync(artifactPath) ? fs.readFileSync(artifactPath, "utf8") : "");
    if (artifact?.m365Package) return artifact.m365Package;
    return path.join(rootDir, "dist", "release", `development-m365-v${rootPackageVersion()}.zip`);
  }
  const artifactPath = path.join(rootDir, "dist/release/artifact.json");
  const artifact = safeJson(fs.existsSync(artifactPath) ? fs.readFileSync(artifactPath, "utf8") : "");
  if (artifact?.package && artifact?.profile === profile) return artifact.package;
  return path.join(rootDir, "dist", "release", `${profile}-v${rootPackageVersion()}.zip`);
}

function readSideloadState() {
  return safeJson(fs.existsSync(sideloadStatePath) ? fs.readFileSync(sideloadStatePath, "utf8") : "") ?? {};
}

function writeSideloadState(state) {
  fs.mkdirSync(sideloadStateDir, { recursive: true });
  fs.writeFileSync(sideloadStatePath, JSON.stringify(state, null, 2) + "\n");
}

function extractTitleId(output) {
  const direct = String(output).match(/\bU_[0-9a-fA-F-]{36}\b/);
  if (direct) return direct[0];
  const labeled = String(output).match(/title\s*id[^A-Za-z0-9_]*(U_)?([0-9a-fA-F-]{36})/i);
  if (labeled) return `${labeled[1] || "U_"}${labeled[2]}`;
  return "";
}

async function runSideload(action, opts = {}, globalOpts = {}) {
  const yes = Boolean(globalOpts.yes);
  const selectedAction = action ?? await maybeSelect("Unified package sideload action", [
    { name: "Install / update local developer sideload", value: "install" },
    { name: "Uninstall last local developer sideload", value: "uninstall" },
    { name: "Show remembered sideload state", value: "status" },
  ], "install", yes);

  const state = readSideloadState();
  if (selectedAction === "status") {
    section("Unified package sideload");
    keyval("state file", path.relative(rootDir, sideloadStatePath));
    keyval("title id", state.titleId || pc.dim("unknown"));
    keyval("package", state.packagePath || pc.dim("unknown"));
    keyval("installed at", state.installedAt || pc.dim("unknown"));
    return;
  }

  const atk = firstExecutable(["atk"]);
  if (!atk && !opts.dryRun) {
    await ensureAgentsToolkit({ yes, dryRun: false });
  }
  const resolvedAtk = firstExecutable(["atk"]) || "atk";

  if (selectedAction === "uninstall") {
    const titleId = opts.titleId || state.titleId || await maybeInput("Title ID to uninstall", state.titleId || "", yes);
    if (!titleId) {
      throw new Error(
        `No title ID found. Pass --title-id U_<guid>, or reinstall once so ${path.relative(rootDir, sideloadStatePath)} can be written.`,
      );
    }
    section("Uninstall unified package sideload");
    keyval("title id", titleId);
    await runMaybe(
      resolvedAtk,
      ["uninstall", "--mode", "title-id", "--title-id", titleId, "--interactive", "false", "--telemetry", "false"],
      {
        dryRun: Boolean(opts.dryRun),
        label: "atk uninstall unified sideload",
      },
    );
    if (!opts.dryRun) {
      writeSideloadState({ ...state, uninstalledAt: new Date().toISOString() });
      statusLine("ok", "uninstall requested", path.relative(rootDir, sideloadStatePath));
    }
    return;
  }

  if (selectedAction !== "install") {
    throw new Error(`Unknown sideload action: ${selectedAction}`);
  }

  if (!opts.skipTunnel) {
    section("Dev tunnel");
    if (opts.dryRun) {
      note("would ensure cloudflared, start/restart Vite + Cloudflare, sync Entra redirect, and regenerate dev manifests");
    } else {
      await runDevTunnel(
        {
          port: opts.port,
          skipEntra: opts.skipEntra,
          keepStaleRedirects: opts.keepStaleRedirects,
        },
        globalOpts,
      );
    }
  }

  if (!opts.skipPackage) {
    section("Package");
    await runPackage({ profile: "development", dryRun: Boolean(opts.dryRun) });
  }

  const packagePath = path.resolve(rootDir, opts.package || m365PackageForProfile("development"));
  if (!opts.dryRun && !fs.existsSync(packagePath)) {
    throw new Error(`Unified package not found: ${packagePath}. Run bun run setup:package first.`);
  }

  section("Install unified package sideload");
  keyval("package", path.relative(rootDir, packagePath));
  if (opts.skipAtkLogin) {
    statusLine("warn", "ATK Microsoft 365 auth", "skipped by --skip-atk-login");
  } else {
    const current = opts.dryRun ? { ok: false, detail: "dry-run" } : atkAuthStatus();
    const shouldLogin =
      Boolean(opts.login) ||
      !current.ok ||
      (await maybeConfirm("Refresh Agents Toolkit Microsoft 365 login before install?", false, yes));
    if (shouldLogin) {
      await loginAgentsToolkitM365({
        force: Boolean(opts.login),
        tenant: opts.tenant,
        dryRun: Boolean(opts.dryRun),
      }, globalOpts);
    } else {
      statusLine("ok", "ATK Microsoft 365 auth", current.detail);
    }
  }
  const output = await runCapture(resolvedAtk, ["install", "--file-path", packagePath, "--telemetry", "false"], {
    dryRun: Boolean(opts.dryRun),
    label: "atk install unified package",
  });
  const titleId = extractTitleId(output);
  if (!opts.dryRun) {
    writeSideloadState({
      titleId,
      packagePath,
      packageSha256: sha256File(packagePath),
      installedAt: new Date().toISOString(),
    });
    if (titleId) {
      statusLine("ok", "remembered title id", titleId);
      note(`uninstall with: bun sideload -- uninstall`);
    } else {
      statusLine("warn", "title id not found in atk output", `state written to ${path.relative(rootDir, sideloadStatePath)}`);
      note("if uninstall needs it, pass --title-id U_<guid> from the atk install output");
    }
  }
}

async function runSkills(opts = {}, globalOpts = {}) {
  const yes = Boolean(globalOpts.yes);
  const mode = opts.mode ?? await maybeSelect("Skill flow", [
    { name: "List private widget skills", value: "list" },
    { name: "Build and show replacement plan", value: "dry-run" },
    { name: "Replace skills with latest zips", value: "update" },
    { name: "Paste cURL/HAR then replace skills", value: "paste-update" },
  ], "list", yes);
  const scriptArgs = [];
  if (mode === "list") scriptArgs.push("--list-only");
  if (mode === "dry-run") scriptArgs.push("--dry-run");
  if (mode === "paste-update") scriptArgs.push("--paste-curl");
  if (mode === "update" && yes) scriptArgs.push("--yes");
  await run("scripts/update-ge-widget-skills.sh", scriptArgs);
}

// ---------------------------------------------------------------------------
// bootstrap: one-shot, idempotent deploy with content-hash versioning.
// (DEPLOY_STATE / GCLOUD_ENV / BOOTSTRAP_SKILLS are declared near the top of the
// file so they are initialized before the command action can run.)
// ---------------------------------------------------------------------------

function sha256File(relPath) {
  const abs = path.isAbsolute(relPath) ? relPath : path.join(rootDir, relPath);
  if (!fs.existsSync(abs)) return null;
  return createHash("sha256").update(fs.readFileSync(abs)).digest("hex");
}

function readDeployState() {
  return safeJson(fs.existsSync(DEPLOY_STATE) ? fs.readFileSync(DEPLOY_STATE, "utf8") : "") ?? {};
}

function manifestVersion() {
  const j = safeJson(
    exists("manifests/m365-unified.manifest.json")
      ? fs.readFileSync(path.join(rootDir, "manifests/m365-unified.manifest.json"), "utf8")
      : "",
  );
  return j?.version ?? null;
}

/** Bump the patch segment of x.y.z in both source manifests. Returns the new version or null. */
function bumpManifestVersion() {
  const uni = path.join(rootDir, "manifests/m365-unified.manifest.json");
  if (!fs.existsSync(uni)) return null;
  const uniTxt = fs.readFileSync(uni, "utf8");
  const m = uniTxt.match(/"version"\s*:\s*"(\d+)\.(\d+)\.(\d+)"/);
  if (!m) return null;
  const next = `${m[1]}.${m[2]}.${Number(m[3]) + 1}`;
  fs.writeFileSync(uni, uniTxt.replace(m[0], `"version": "${next}"`));
  const one = path.join(rootDir, "manifests/onenote.manifest.xml");
  if (fs.existsSync(one)) {
    fs.writeFileSync(
      one,
      fs.readFileSync(one, "utf8").replace(/<Version>[\d.]+<\/Version>/, `<Version>${next}</Version>`),
    );
  }
  return next;
}

/** Extract the agent id (last path segment) from a `label=projects/.../agents/<id>` env value. */
function agentIdFromEnvSkill(envName) {
  const raw = envValue(envName);
  if (!raw) return null;
  const resource = raw.includes("=") ? raw.slice(raw.indexOf("=") + 1) : raw;
  const at = resource.lastIndexOf("/agents/");
  return at >= 0 ? resource.slice(at + "/agents/".length) : null;
}

function skillResourceEnvValue(label, geEnv, agentId) {
  const location = geEnv.GE_LOCATION || "global";
  return `${label}=projects/${geEnv.GE_PROJECT_NUMBER}/locations/${location}/collections/default_collection/engines/${geEnv.GE_ENGINE}/assistants/default_assistant/agents/${agentId}`;
}

function writeEnvAssignments(filePath, assignments) {
  const existing = fs.existsSync(filePath) ? fs.readFileSync(filePath, "utf8").split(/\r?\n/) : [];
  const seen = new Set();
  const keys = new Set(Object.keys(assignments));
  const next = existing.map((line) => {
    const match = line.match(/^(\s*(?:export\s+)?)([A-Za-z_][A-Za-z0-9_]*)(\s*=).*/);
    if (!match || !keys.has(match[2])) return line;
    seen.add(match[2]);
    return `${match[1]}${match[2]}=${assignments[match[2]]}`;
  });
  for (const [key, value] of Object.entries(assignments)) {
    if (!seen.has(key)) next.push(`${key}=${value}`);
  }
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, next.join("\n").replace(/\n*$/, "\n"));
}

/** Read `name` + `description` from a skill's SKILL.md YAML frontmatter (handles folded `>-`/`>` scalars). */
function skillFrontmatter(name) {
  const p = path.join(rootDir, `skill/${name}/SKILL.md`);
  if (!fs.existsSync(p)) return { name };
  const fmMatch = fs.readFileSync(p, "utf8").match(/^---\r?\n([\s\S]*?)\r?\n---/);
  if (!fmMatch) return { name };
  const lines = fmMatch[1].split(/\r?\n/);
  const nameLine = lines.find((l) => /^name:\s*/.test(l));
  const displayName = nameLine ? nameLine.replace(/^name:\s*/, "").trim() : name;
  let description = "";
  const dIdx = lines.findIndex((l) => /^description:\s*/.test(l));
  if (dIdx >= 0) {
    const inline = lines[dIdx].replace(/^description:\s*/, "").trim();
    if (inline && !inline.startsWith(">") && !inline.startsWith("|")) {
      description = inline;
    } else {
      const buf = [];
      for (let i = dIdx + 1; i < lines.length && /^\s/.test(lines[i]); i++) buf.push(lines[i].trim());
      description = buf.join(" ").trim();
    }
  }
  const versionLine = lines.find((l) => /^\s*version:\s*/.test(l));
  const version = versionLine ? versionLine.replace(/^\s*version:\s*/, "").trim().replace(/^['"]|['"]$/g, "") : "";
  return { name: displayName, description: description.slice(0, 1024), version };
}

function versionedAgentId(baseAgentId, version) {
  const suffix = String(version || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return suffix ? `${baseAgentId}-v${suffix}` : `${baseAgentId}-live`;
}

function bootstrapAgentCandidates(skill, version) {
  return [
    ...new Set(
      [agentIdFromEnvSkill(skill.idEnv), skill.agentId, versionedAgentId(skill.agentId, version)].filter(Boolean),
    ),
  ];
}

/** The numeric GCP project number embedded in a skill resource name (`projects/<number>/...`). */
function projectNumberFromEnv() {
  for (const name of ["VITE_GE_SURFACE_COMMANDER_SKILL", "VITE_GE_COMMAND_PLANNER_SKILL"]) {
    const m = envValue(name).match(/projects\/(\d+)\//);
    if (m) return m[1];
  }
  return "";
}

/** Env for `create_skill.py --live`: its GE_* vars mapped from the web-shell .env + the WIF config dir. */
function geProvisioningEnv() {
  return {
    ...GCLOUD_ENV(),
    // Use the core `gcloud auth login` token (refreshed by step 2) via print-access-token, NOT ADC.
    // Our login flow deliberately refreshes core creds and leaves ADC untouched, so this is the
    // token that is actually fresh; GE_AUTH_MODE=gcloud makes create_skill.py use it.
    GE_AUTH_MODE: "gcloud",
    GE_PROJECT: envValue("VITE_GCP_PROJECT"),
    GE_PROJECT_NUMBER: projectNumberFromEnv(),
    GE_ENGINE: envValue("VITE_GE_ENGINE"),
    GE_LOCATION: envValue("VITE_GCP_LOCATION") || "global",
  };
}

function haveGoogleToken() {
  return repoWifTokenStatus().ok;
}

async function runBootstrap(opts = {}, globalOpts = {}) {
  const dryRun = Boolean(opts.dryRun);
  const force = Boolean(opts.force);
  const profile = opts.profile ?? "development";
  const prev = readDeployState();
  let profileConfig = null;
  if (profile !== "development") {
    profileConfig = resolveReleaseConfig(profile);
  }

  if (profile === "development" && opts.devTunnel) {
    section("dev tunnel");
    if (dryRun) {
      note("would ensure cloudflared, start/restart Vite + Cloudflare, sync Entra redirect, and regenerate dev manifests");
    } else {
      await runDevTunnel(
        {
          port: opts.port,
          skipEntra: opts.skipEntra,
          keepStaleRedirects: opts.keepStaleRedirects,
        },
        globalOpts,
      );
    }
  }

  // --- fingerprint the deployable artifacts ---
  const manifestSha = createHash("sha256")
    .update((sha256File("manifests/m365-unified.manifest.json") ?? "") + "|" + (sha256File("manifests/onenote.manifest.xml") ?? ""))
    .digest("hex");
  const skillShas = {};
  for (const s of BOOTSTRAP_SKILLS) skillShas[s.name] = sha256File(s.zip);

  const manifestChanged = force || prev.manifestSha !== manifestSha;
  const changedSkills = BOOTSTRAP_SKILLS.filter(
    (s) => force || (prev.skills?.[s.name] ?? null) !== skillShas[s.name],
  );

  const yn = (b) => (b ? pc.yellow("yes") : pc.dim("no"));
  section("Bootstrap plan");
  keyval("profile", profile);
  keyval("version", `${pc.dim(prev.manifestVersion ?? "none")} ${pc.dim("→")} ${manifestVersion() ?? "?"}`);
  keyval("manifest changed", yn(manifestChanged));
  keyval("skills to deploy", changedSkills.map((s) => s.name).join(", ") || pc.dim("none"));
  if (dryRun) note("dry-run: destructive steps are printed, not executed");

  // 1 · prereqs (read-only status)
  if (!opts.deployOnly) {
    section("1 · prereqs");
    await runDoctor({});
  }

  // 2 · login — ensure a fresh Google token for the public AgentService (create_skill.py uses ADC)
  if (!opts.deployOnly) {
    section("2 · login (Google token for skill provisioning)");
    if (haveGoogleToken()) statusLine("ok", "repo-local WIF token", "fresh enough to continue");
    else if (dryRun) note("would run setup:gcloud:wif; ADC remains untouched");
    else await runLogin({ target: "gcloud" });
  }

  // 3 · package (manifests + web build) when the manifest changed or artifacts are missing
  let version = manifestVersion();
  const artifactsMissing =
    profile === "development"
      ? !exists(`dist/package/${profile}/xml/word.manifest.xml`)
      : !exists(`dist/release/${profile}-v${rootPackageVersion()}.zip`);
  if (manifestChanged || artifactsMissing) {
    if (manifestChanged && opts.bump && !dryRun) {
      const bumped = bumpManifestVersion();
      if (bumped) {
        version = bumped;
        console.log(`  bumped manifest version → ${bumped}`);
      }
    }
    section("3 · package");
    await runPackage({ profile, dryRun });
  } else {
    section("3 · package");
    statusLine("ok", "package skipped", "manifest unchanged and artifacts present");
  }

  // 4 · skills via the public AgentService (only the changed bundles)
  if (changedSkills.length) {
    section("4 · skills (public AgentService)");
    const geEnv = geProvisioningEnv();
    const missing = ["GE_PROJECT", "GE_PROJECT_NUMBER", "GE_ENGINE"].filter((k) => !geEnv[k]);
    if (missing.length) {
      statusLine(
        "warn",
        "skills skipped",
        `missing ${missing.join(", ")}; fill packages/web-shell/.env before provisioning`,
      );
    } else {
      keyval("target", `${geEnv.GE_PROJECT} (${geEnv.GE_PROJECT_NUMBER})`);
      keyval("engine", geEnv.GE_ENGINE);
      const python = firstExecutable(["python3", "python"]) || "python3";
      const envAssignments = {};
      for (const s of changedSkills) {
        const fm = skillFrontmatter(s.name);
        const candidateAgentIds = bootstrapAgentCandidates(s, fm.version);
        const envAgentId = agentIdFromEnvSkill(s.idEnv);
        if (envAgentId) {
          note(`${s.idEnv} starts with configured id ${envAgentId}; fallback is ${s.agentId}`);
        }
        let provisionedAgentId = "";
        for (const [index, agentId] of candidateAgentIds.entries()) {
          const args = [
            "skill/create_skill.py",
            "--live",
            "--api-mode",
            "public",
            // NOT --replace: delete+create soft-tombstones the stable agent id (create→409 / get→404
            // afterward). create_shell upserts (create, else PATCH in place), keeping the id.
            "--yes",
            "--zip",
            s.zip,
            "--display-name",
            fm.name || s.name,
            ...(fm.description ? ["--description", fm.description] : []),
            "--agent-id",
            agentId,
          ];
          try {
            await runMaybe(python, args, {
              dryRun,
              env: geEnv,
              label: `upsert ${s.name} skill via public AgentService (${agentId})`,
            });
            provisionedAgentId = agentId;
            break;
          } catch (error) {
            if (index === candidateAgentIds.length - 1) throw error;
            statusLine("warn", `${s.name} id unavailable`, `retrying with ${candidateAgentIds[index + 1]}`);
          }
        }
        if (provisionedAgentId) {
          envAssignments[s.idEnv] = skillResourceEnvValue(s.name, geEnv, provisionedAgentId);
        }
      }
      if (!dryRun && Object.keys(envAssignments).length) {
        writeEnvAssignments(webEnvPath, envAssignments);
        statusLine("ok", "updated skill refs", path.relative(rootDir, webEnvPath));
      }
    }
  } else {
    section("4 · skills");
    statusLine("ok", "skills skipped", "bundles unchanged");
  }

  // 5 · tenant deploy/upload (only when the manifest changed)
  let tenantDeployCompleted = false;
  let completedDeploymentLane = "";
  if (manifestChanged) {
    section("5 · deploy");
    const action = prev.manifestSha ? "update" : "deploy";
    const assignment = opts.assignment ?? "members";
    const requestedLane = opts.skipTenantDeploy ? "none" : opts.deploymentLane ?? "auto";
    const tenantBackend = opts.tenantBackend ?? "Auto";
    keyval("action", action);
    keyval("assignment", assignment);
    keyval("lane", requestedLane);
    if (dryRun) {
      if (requestedLane !== "none" && !opts.skipEntra) {
        const cfg = profileConfig ?? resolveReleaseConfig(profile);
        note(`would sync Entra SPA redirect: ${cfg.webOrigin}/auth-redirect.html`);
      }
      if (requestedLane === "catalog") {
        note("would install CLI for Microsoft 365 if needed, login with device code, and upsert unified package");
      } else if (requestedLane === "none") {
        note("would skip tenant deployment/upload");
      } else {
        note("would choose XML Centralized Deployment if supported, otherwise unified app catalog upload");
      }
    } else if (requestedLane === "none") {
      statusLine("warn", "deploy skipped", "requested with --skip-tenant-deploy or --deployment-lane none");
      note(`Windows handoff: ${tenantDeployHandoff(action, assignment, tenantBackend)}`);
    } else {
      const yes = Boolean(globalOpts.yes);
      const support = requestedLane === "catalog" ? { ok: false } : m365TenantDeploySupport(tenantBackend);
      const resolvedLane = requestedLane === "auto" ? (support.ok || opts.forceTenantDeploy ? "xml" : "catalog") : requestedLane;
      keyval("resolved lane", resolvedLane);

      if (!opts.skipEntra) {
        section("5a · Entra SPA redirect");
        await runEntraSync(profile, {
          dryRun: false,
          keepStaleRedirects: opts.keepStaleRedirects,
        });
        section("5b · publish");
      }

      if (resolvedLane === "catalog") {
        note("using unified M365 package app catalog upload; this is not the same as XML Centralized Deployment assignment");
        await runM365Catalog("upsert", { installCli: true, login: true, package: m365PackageForProfile(profile) }, globalOpts);
        tenantDeployCompleted = true;
        completedDeploymentLane = "catalog";
      } else if (resolvedLane === "xml") {
        keyval("backend", tenantBackend);
        if (!opts.deployOnly) {
          await ensurePowerShell({ yes, dryRun: false });
          await ensureExchangeOnlineManagement({ yes, dryRun: false });
        }
        if (!support.ok && !opts.forceTenantDeploy) {
          statusLine("warn", "XML tenant deploy skipped", support.reason);
          note("the package and manifests were generated successfully; live XML tenant deployment must run elsewhere");
          note(`Windows handoff: ${tenantDeployHandoff(action, assignment, tenantBackend)}`);
        } else {
          const resolvedBackend = tenantBackend === "Auto" ? support.backend ?? "Auto" : tenantBackend;
          if (support.backend) keyval("resolved backend", support.backend);
          await runAddins(action, { assignment, backend: resolvedBackend }, globalOpts);
          tenantDeployCompleted = true;
          completedDeploymentLane = "xml";
        }
      } else {
        throw new Error(`Unknown deployment lane: ${resolvedLane}`);
      }
    }
  } else {
    section("5 · deploy");
    statusLine("ok", "deploy skipped", "manifest unchanged; users reload the web app from the manifest URL");
    tenantDeployCompleted = true;
    completedDeploymentLane = prev.deploymentLane ?? "unchanged";
  }

  // 6 · persist state
  if (!dryRun) {
    const next = {
      manifestVersion: tenantDeployCompleted ? version : prev.manifestVersion ?? version,
      profile,
      manifestSha: tenantDeployCompleted ? manifestSha : prev.manifestSha,
      deploymentLane: tenantDeployCompleted ? completedDeploymentLane : prev.deploymentLane,
      skills: skillShas,
      deployedAt: new Date().toISOString(),
      ...(tenantDeployCompleted
        ? {}
        : {
            tenantDeploySkippedAt: new Date().toISOString(),
            tenantDeploySkippedReason: "local environment cannot complete live Microsoft 365 Centralized Deployment",
          }),
    };
    fs.writeFileSync(DEPLOY_STATE, JSON.stringify(next, null, 2) + "\n");
    section("state");
    statusLine("ok", `wrote ${path.relative(rootDir, DEPLOY_STATE)}`, `version ${version ?? "?"}`);
  }

  if (opts.dev) {
    section("dev");
    note("start a live origin to debug the deployed manifest:");
    console.log(`  ${muted("1.")} bun run ge:dev:tunnel`);
    console.log(`  ${muted("2.")} bun run --filter @ge/web-shell dev`);
  }
}

program.parseAsync(process.argv).catch((error) => {
  console.error(STATUS.missing.paint(STATUS.missing.symbol), error.message || error);
  process.exit(1);
});
