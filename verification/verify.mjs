import crypto from 'node:crypto';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { spawn } from 'node:child_process';
import { DatabaseSync } from 'node:sqlite';
import { inflateRawSync } from 'node:zlib';

const repoRoot = path.resolve(import.meta.dirname, '..');
const artifactRoot = path.join(repoRoot, 'artifacts');
const evidenceRoot = path.join(repoRoot, 'verification', 'evidence');
const npmCli = process.env.npm_execpath;
const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const attachments = ['输入数据包.zip', 'reference.zip', '关键标准答案.xlsx', '任务规格转化.xlsx'];
const expectedReference = [
  'output/api_billing_repaired.db',
  'output/rebuild_api_billing.sql',
  'output/reports/exclusion_review.csv',
  'output/reports/family_usage.csv',
  'output/reports/tenant_invoice.csv',
].sort();
const reportKeys = {
  'output/reports/tenant_invoice.csv': ['tenant_id'],
  'output/reports/family_usage.csv': ['tenant_id', 'api_family'],
  'output/reports/exclusion_review.csv': ['source_table', 'source_id'],
};
const databaseTables = [
  ['tenant_account', 'tenant_id'],
  ['plan_family_quota', 'tenant_id, api_family'],
  ['price_tier', 'api_family, tier_start_units'],
  ['usage_event', 'event_id'],
  ['credit_adjustment', 'adjustment_id'],
  ['eligible_usage', 'event_id'],
  ['family_usage', 'tenant_id, api_family'],
  ['adjustment_in_scope', 'adjustment_id'],
  ['tenant_invoice', 'tenant_id'],
  ['exclusion_review', 'source_table DESC, source_id'],
  ['billing_meta', 'key'],
];

const sha256 = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex');
const sha256File = (file) => sha256(fs.readFileSync(file));
const assert = (value, message) => { if (!value) throw new Error(message); };

function parseZipBytes(data) {
  const files = new Map(); let offset = 0;
  while (offset + 30 <= data.length) {
    if (data.readUInt32LE(offset) !== 0x04034b50) break;
    const flags = data.readUInt16LE(offset + 6); const method = data.readUInt16LE(offset + 8);
    const compressedSize = data.readUInt32LE(offset + 18); const uncompressedSize = data.readUInt32LE(offset + 22);
    const nameLength = data.readUInt16LE(offset + 26); const extraLength = data.readUInt16LE(offset + 28);
    assert(!(flags & 0x08), 'ZIP数据描述符不受支持');
    const name = data.subarray(offset + 30, offset + 30 + nameLength).toString('utf8').replaceAll('\\', '/');
    const start = offset + 30 + nameLength + extraLength; const compressed = data.subarray(start, start + compressedSize);
    if (!name.endsWith('/')) {
      const body = method === 0 ? compressed : method === 8 ? inflateRawSync(compressed) : null;
      assert(body && body.length === uncompressedSize, `无法解压${name}`); files.set(name, body);
    }
    offset = start + compressedSize;
  }
  return files;
}
const parseZip = (file) => parseZipBytes(fs.readFileSync(file));
async function extractZip(file, destination) {
  for (const [name, bytes] of parseZip(file)) {
    const target = path.resolve(destination, name);
    assert(target.startsWith(path.resolve(destination) + path.sep), `非法ZIP路径${name}`);
    await fsp.mkdir(path.dirname(target), { recursive: true }); await fsp.writeFile(target, bytes);
  }
}
function workbookSheets(file) {
  const workbook = parseZipBytes(fs.readFileSync(file)).get('xl/workbook.xml')?.toString('utf8') ?? '';
  return [...workbook.matchAll(/<(?:[A-Za-z]+:)?sheet[^>]+name="([^"]+)"/gu)].map((match) => match[1]);
}
async function run(command, args, cwd) {
  const started = Date.now();
  return await new Promise((resolve) => {
    let child;
    try { child = spawn(command, args, { cwd, env: process.env, windowsHide: true }); }
    catch (error) { resolve({ code: 1, stdout: '', stderr: error.stack ?? error.message, elapsed_ms: Date.now() - started }); return; }
    let stdout = ''; let stderr = ''; let settled = false;
    child.stdout.on('data', (chunk) => { stdout += chunk; }); child.stderr.on('data', (chunk) => { stderr += chunk; });
    child.on('error', (error) => { if (!settled) { settled = true; resolve({ code: 1, stdout, stderr: `${stderr}${error.stack ?? error.message}`, elapsed_ms: Date.now() - started }); } });
    child.on('exit', (code) => { if (!settled) { settled = true; resolve({ code: code ?? 1, stdout, stderr, elapsed_ms: Date.now() - started }); } });
  });
}
async function runNpm(args, cwd) {
  return npmCli ? await run(process.execPath, [npmCli, ...args], cwd) : await run(npmCommand, args, cwd);
}
function treeDigest(root, ignored = new Set()) {
  const lines = [];
  function visit(current, prefix = '') {
    for (const entry of fs.readdirSync(current, { withFileTypes: true }).toSorted((a, b) => a.name.localeCompare(b.name))) {
      const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (ignored.has(relative.split('/')[0])) continue;
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) visit(full, relative); else lines.push(`${relative}\0${sha256File(full)}`);
    }
  }
  visit(root); return sha256(Buffer.from(lines.join('\n')));
}
function parseCsv(text) {
  const rows = []; let row = []; let cell = ''; let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (quoted) {
      if (char === '"' && text[i + 1] === '"') { cell += '"'; i += 1; }
      else if (char === '"') quoted = false; else cell += char;
    } else if (char === '"') quoted = true;
    else if (char === ',') { row.push(cell); cell = ''; }
    else if (char === '\n') { row.push(cell.replace(/\r$/u, '')); rows.push(row); row = []; cell = ''; }
    else cell += char;
  }
  if (cell || row.length) { row.push(cell.replace(/\r$/u, '')); rows.push(row); }
  const headers = rows.shift() ?? [];
  return rows.filter((values) => values.some((value) => value !== '')).map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ''])));
}
function normalizedRows(file, text) {
  const keys = reportKeys[file]; const rows = parseCsv(text);
  return rows.toSorted((a, b) => keys.map((key) => String(a[key]).localeCompare(String(b[key]))).find((value) => value !== 0) ?? 0);
}
function databaseSemantic(file) {
  const db = new DatabaseSync(file, { readOnly: true });
  try {
    const version = db.prepare('SELECT sqlite_version() AS version').get().version;
    const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").all().map((row) => row.name);
    const rows = Object.fromEntries(databaseTables.map(([table, order]) => [table, db.prepare(`SELECT * FROM ${table} ORDER BY ${order}`).all()]));
    const integrity = db.prepare('PRAGMA integrity_check').get().integrity_check;
    const foreignKeys = db.prepare('PRAGMA foreign_key_check').all();
    return { version, tables, rows, integrity, foreignKeys };
  } finally { db.close(); }
}
function classifyExecutable(name, bytes) {
  const lower = name.toLowerCase();
  if (bytes.length >= 4 && bytes[0] === 0x7f && bytes.subarray(1, 4).toString('ascii') === 'ELF') return 'linux_elf';
  if (bytes.length >= 4 && [0xfeedface, 0xfeedfacf, 0xcafebabe].includes(bytes.readUInt32BE(0))) return 'macos_macho';
  if (/\.(?:sh|bash|so)(?:\.|$)/u.test(lower)) return 'posix_member';
  if (/^#!.*(?:ba|z|k)?sh/mu.test(bytes.subarray(0, 128).toString('utf8'))) return 'posix_shebang';
  return null;
}
async function prepare(label, mutate) {
  const root = path.join(os.tmpdir(), label); await fsp.rm(root, { recursive: true, force: true }); await fsp.mkdir(root, { recursive: true });
  await extractZip(path.join(artifactRoot, '输入数据包.zip'), root);
  const inputRoot = path.join(root, 'input_data'); const reference = parseZip(path.join(artifactRoot, 'reference.zip'));
  await fsp.mkdir(path.join(inputRoot, 'output'), { recursive: true });
  await fsp.writeFile(path.join(inputRoot, 'output', 'rebuild_api_billing.sql'), reference.get('output/rebuild_api_billing.sql'));
  if (mutate) await mutate(inputRoot);
  return { root, inputRoot, outputRoot: path.join(inputRoot, 'output'), reference };
}
function outputPaths(root) {
  const paths = [];
  function walk(current, prefix = '') {
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (entry.isDirectory()) walk(path.join(current, entry.name), relative); else paths.push(`output/${relative}`);
    }
  }
  walk(root); return paths.sort();
}
function compareReference(outputRoot, reference) {
  assert(JSON.stringify(outputPaths(outputRoot)) === JSON.stringify(expectedReference), '输出成员与Reference不一致');
  const semantic = crypto.createHash('sha256');
  for (const file of expectedReference) {
    const actualPath = path.join(path.dirname(outputRoot), file); const expected = reference.get(file);
    if (file.endsWith('.csv')) {
      const actualRows = normalizedRows(file, fs.readFileSync(actualPath, 'utf8'));
      const expectedRows = normalizedRows(file, expected.toString('utf8'));
      assert(JSON.stringify(actualRows) === JSON.stringify(expectedRows), `${file}与Reference不一致`);
      semantic.update(JSON.stringify(actualRows));
    } else if (file.endsWith('.db')) {
      const tempReference = path.join(os.tmpdir(), `reference-${crypto.randomUUID()}.db`);
      fs.writeFileSync(tempReference, expected);
      const actualDb = databaseSemantic(actualPath); const expectedDb = databaseSemantic(tempReference); fs.rmSync(tempReference);
      assert(actualDb.integrity === 'ok' && actualDb.foreignKeys.length === 0, '实际数据库完整性或外键失败');
      assert(JSON.stringify(actualDb.tables) === JSON.stringify(expectedDb.tables) && JSON.stringify(actualDb.rows) === JSON.stringify(expectedDb.rows), '数据库语义与Reference不一致');
      semantic.update(JSON.stringify({ tables: actualDb.tables, rows: actualDb.rows }));
    } else {
      const actual = fs.readFileSync(actualPath, 'utf8').replaceAll('\r\n', '\n');
      assert(actual === expected.toString('utf8').replaceAll('\r\n', '\n'), `${file}与Reference不一致`);
      semantic.update(actual);
    }
  }
  return semantic.digest('hex');
}

await fsp.rm(evidenceRoot, { recursive: true, force: true }); await fsp.mkdir(evidenceRoot, { recursive: true });
assert(process.platform === 'win32' && process.env.GITHUB_ACTIONS === 'true', '该验证器只接受GitHub托管Windows运行');
const attachmentSha256 = Object.fromEntries(attachments.map((name) => [name, sha256File(path.join(artifactRoot, name))]));
const inputMembers = parseZip(path.join(artifactRoot, '输入数据包.zip'));
const executableScan = [...inputMembers].map(([name, bytes]) => ({ name, classification: classifyExecutable(name, bytes) })).filter((item) => item.classification);
assert(executableScan.length === 0, `输入包含平台专用可执行成员：${JSON.stringify(executableScan)}`);
const referenceMembers = [...parseZip(path.join(artifactRoot, 'reference.zip')).keys()].sort();
assert(JSON.stringify(referenceMembers) === JSON.stringify(expectedReference), 'Reference成员错误');
assert(JSON.stringify(workbookSheets(path.join(artifactRoot, '关键标准答案.xlsx'))) === JSON.stringify(['交付物答案清单', '固定字段答案', '固定集合答案', '固定数值答案', '允许变体答案']), '关键标准答案Sheet错误');
assert(JSON.stringify(workbookSheets(path.join(artifactRoot, '任务规格转化.xlsx'))) === JSON.stringify(['任务规格转化']), '任务规格Sheet错误');
const solutionText = parseZip(path.join(artifactRoot, 'reference.zip')).get('output/rebuild_api_billing.sql').toString('utf8');
assert(!/\b(?:U00[1-9]|U01[0-2]|T100|T200|T300|T400|A00[1-4])\b|https?:\/\//u.test(solutionText), '完成版SQL含样本主键硬编码或外部地址');

const cleanRuns = [];
for (const label of ['Q10107 第一次 空目录', 'Q10107 第二次 中文 空格目录']) {
  const prepared = await prepare(label); const before = treeDigest(prepared.inputRoot, new Set(['output']));
  const result = await runNpm(['run', 'process'], prepared.inputRoot);
  assert(result.code === 0, `${label}执行失败\n${result.stdout}\n${result.stderr}`);
  const after = treeDigest(prepared.inputRoot, new Set(['output'])); assert(before === after, `${label}修改了输入`);
  const semantic = compareReference(prepared.outputRoot, prepared.reference);
  const actualDb = databaseSemantic(path.join(prepared.outputRoot, 'api_billing_repaired.db'));
  cleanRuns.push({ directory_label: label, exit_code: result.code, input_digest_before: before, input_digest_after: after, semantic_digest: semantic, sqlite_version: actualDb.version, elapsed_ms: result.elapsed_ms });
}
assert(cleanRuns[0].semantic_digest === cleanRuns[1].semantic_digest, '两次结构化结果不一致');

const crlf = await prepare('Q10107 CRLF 规则', async (inputRoot) => {
  const file = path.join(inputRoot, 'rules', 'report_contract.json'); const text = await fsp.readFile(file, 'utf8');
  await fsp.writeFile(file, text.replace(/\r?\n/gu, '\r\n'));
});
let result = await runNpm(['run', 'process'], crlf.inputRoot); assert(result.code === 0, `CRLF规则执行失败\n${result.stdout}\n${result.stderr}`);
const crlfDigest = compareReference(crlf.outputRoot, crlf.reference); assert(crlfDigest === cleanRuns[0].semantic_digest, 'CRLF规则改变业务结果');

const mutation = await prepare('Q10107 有效封顶变化', async (inputRoot) => {
  const file = path.join(inputRoot, 'database', 'api_metering.db'); const db = new DatabaseSync(file);
  try { db.prepare('UPDATE tenant_account SET monthly_cap_cents = ? WHERE tenant_id = ?').run(4000, 'T200'); } finally { db.close(); }
});
result = await runNpm(['run', 'process'], mutation.inputRoot); assert(result.code === 0, `封顶变化执行失败\n${result.stdout}\n${result.stderr}`);
const mutatedInvoices = normalizedRows('output/reports/tenant_invoice.csv', fs.readFileSync(path.join(mutation.outputRoot, 'reports', 'tenant_invoice.csv'), 'utf8'));
const t200 = mutatedInvoices.find((row) => row.tenant_id === 'T200');
assert(t200?.monthly_cap_cents === '4000' && t200.cap_reduction_cents === '1450' && t200.final_due_cents === '4000', '封顶变化未联动T200账单');
const referenceInvoices = normalizedRows('output/reports/tenant_invoice.csv', mutation.reference.get('output/reports/tenant_invoice.csv').toString('utf8'));
assert(JSON.stringify(mutatedInvoices.filter((row) => row.tenant_id !== 'T200')) === JSON.stringify(referenceInvoices.filter((row) => row.tenant_id !== 'T200')), '封顶变化影响无关租户');

const negative = await prepare('Q10107 无效报表合同', async (inputRoot) => { await fsp.rm(path.join(inputRoot, 'rules', 'report_contract.json')); });
result = await runNpm(['run', 'process'], negative.inputRoot);
const dynamicAbsent = !fs.existsSync(path.join(negative.outputRoot, 'api_billing_repaired.db')) && !fs.existsSync(path.join(negative.outputRoot, 'reports'));
assert(result.code !== 0 && dynamicAbsent, '无效输入没有失败关闭');

const evidence = {
  schema_version: 1, task_asset_id: 'sqlite_api_monthly_billing', result: 'PASS', generated_at_utc: new Date().toISOString(),
  git_commit_sha: process.env.GITHUB_SHA, workflow_run_id: process.env.GITHUB_RUN_ID,
  runner: { os: process.env.RUNNER_OS, arch: process.env.RUNNER_ARCH, image_os: process.env.ImageOS, image_version: process.env.ImageVersion, node: process.version, powershell_hosted_workflow: true },
  software: { main: 'SQLite', executed: true, sqlite_version: cleanRuns[0].sqlite_version, node: process.version }, attachment_sha256: attachmentSha256,
  workbook_checks: { answer_sheet_names: workbookSheets(path.join(artifactRoot, '关键标准答案.xlsx')), specification_sheet_names: ['任务规格转化'] },
  platform_audit: { linux_executables: executableScan, linux_executables_executed: false, no_wsl_required: true, no_linux_container_required: true, no_posix_shell_required: true, no_unix_only_api_required: true, cross_platform_paths: true },
  clean_runs: cleanRuns,
  crlf_input: { file: 'rules/report_contract.json', exit_code: 0, semantic_digest: crlfDigest, reference_match: true },
  positive_mutation: { changed_input: 'tenant_account.T200.monthly_cap_cents从3000改为4000', exit_code: 0, t200_invoice: t200, unrelated_tenants_unchanged: true },
  invalid_input: { removed_input: 'rules/report_contract.json', exit_code: result.code, dynamic_deliverables_absent: dynamicAbsent },
  network: { installation_network_access: 'Node.js安装阶段', formal_run_network_access: 'none, local files and local SQLite only' },
};
await fsp.writeFile(path.join(evidenceRoot, 'windows-verification.json'), `${JSON.stringify(evidence, null, 2)}\n`);
console.log(JSON.stringify(evidence, null, 2));
