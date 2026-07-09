#!/usr/bin/env node
'use strict';

const fs = require('fs');
const https = require('https');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const SERVERS = new Set(['excel', 'md', 'pdf', 'docx', 'pptx', 'csv', 'html', 'text', 'json']);
const OWNER = 'ndhkaeru';
const REPO = 'docloupe-mcp';

function platformKey() {
  const platform = process.platform;
  const arch = process.arch;
  if (platform === 'win32' && arch === 'x64') return 'win32-x64';
  if (platform === 'linux' && arch === 'x64') return 'linux-x64';
  if (platform === 'darwin' && arch === 'x64') return 'darwin-x64';
  if (platform === 'darwin' && arch === 'arm64') return 'darwin-arm64';
  throw new Error(`Unsupported platform: ${platform}-${arch}. Supported: win32-x64, linux-x64, darwin-x64, darwin-arm64.`);
}

function releasePlatform() {
  return {
    'win32-x64': 'windows-x64',
    'linux-x64': 'linux-x64',
    'darwin-x64': 'macos-x64',
    'darwin-arm64': 'macos-arm64',
  }[platformKey()];
}

function executableName(server) {
  return `${server}-tools${process.platform === 'win32' ? '.exe' : ''}`;
}

function envName(server) {
  return `DOCLOUPE_${server.toUpperCase().replace(/-/g, '_')}_TOOLS_BINARY`;
}

function packageVersion() {
  return require('../package.json').version;
}

function releaseTag() {
  return process.env.DOCLOUPE_MCP_RELEASE_TAG || `v${packageVersion()}`;
}

function cacheRoot() {
  if (process.env.DOCLOUPE_MCP_CACHE_DIR) return process.env.DOCLOUPE_MCP_CACHE_DIR;
  if (process.platform === 'win32' && process.env.LOCALAPPDATA) {
    return path.join(process.env.LOCALAPPDATA, 'docloupe-mcp');
  }
  return path.join(os.homedir(), '.cache', 'docloupe-mcp');
}

function cachedBinary(server) {
  return path.join(cacheRoot(), releaseTag(), platformKey(), executableName(server));
}

function assetUrl(server) {
  const suffix = process.platform === 'win32' ? '.exe' : '';
  const asset = `docloupe-mcp-${server}-tools-${releasePlatform()}${suffix}`;
  return `https://github.com/${OWNER}/${REPO}/releases/download/${releaseTag()}/${asset}`;
}

function download(url, outputPath, redirects = 0) {
  return new Promise((resolve, reject) => {
    const request = https.get(url, { headers: { 'User-Agent': 'docloupe-mcp-npm' } }, (response) => {
      if ([301, 302, 303, 307, 308].includes(response.statusCode)) {
        response.resume();
        if (!response.headers.location || redirects >= 5) {
          reject(new Error(`Too many redirects while downloading ${url}`));
          return;
        }
        download(response.headers.location, outputPath, redirects + 1).then(resolve, reject);
        return;
      }
      if (response.statusCode < 200 || response.statusCode >= 300) {
        reject(new Error(`Download failed (${response.statusCode}): ${url}`));
        response.resume();
        return;
      }

      const tmpPath = `${outputPath}.tmp`;
      const file = fs.createWriteStream(tmpPath);
      response.pipe(file);
      file.on('finish', () => {
        file.close(() => {
          fs.renameSync(tmpPath, outputPath);
          if (process.platform !== 'win32') fs.chmodSync(outputPath, 0o755);
          resolve();
        });
      });
      file.on('error', (error) => {
        fs.rmSync(tmpPath, { force: true });
        reject(error);
      });
    });
    request.on('error', reject);
  });
}

async function findBinary(server) {
  const override = process.env[envName(server)] || process.env.DOCLOUPE_MCP_BINARY;
  if (override) return override;

  const bundled = path.join(__dirname, '..', 'native', platformKey(), executableName(server));
  if (fs.existsSync(bundled)) return bundled;

  const cached = cachedBinary(server);
  if (fs.existsSync(cached)) return cached;

  fs.mkdirSync(path.dirname(cached), { recursive: true });
  console.error(`Downloading docloupe ${server}-tools ${releaseTag()} for ${platformKey()}...`);
  await download(assetUrl(server), cached);
  return cached;
}

function usage() {
  console.error([
    'Usage:',
    '  docloupe-mcp <excel|md|pdf|docx|pptx|csv|html|text|json> [server args...]',
    '  docloupe-excel-tools [server args...]',
    '',
    'Environment overrides:',
    '  DOCLOUPE_EXCEL_TOOLS_BINARY=/path/to/excel-tools',
    '  DOCLOUPE_MCP_BINARY=/path/to/server-binary',
    '  DOCLOUPE_MCP_CACHE_DIR=/path/to/cache',
    '  DOCLOUPE_MCP_RELEASE_TAG=v1.2.3',
  ].join('\n'));
}

async function runAsync(server, args) {
  if (!SERVERS.has(server)) {
    usage();
    process.exit(2);
  }

  const binary = await findBinary(server);
  const child = spawn(binary, args, { stdio: 'inherit', windowsHide: true });
  child.on('error', (error) => {
    console.error(error.message);
    process.exit(1);
  });
  child.on('exit', (code, signal) => {
    if (signal) {
      process.kill(process.pid, signal);
      return;
    }
    process.exit(code ?? 0);
  });
}

function run(server, args) {
  runAsync(server, args).catch((error) => {
    console.error(error.stack || error.message);
    process.exit(1);
  });
}

function main() {
  const invoked = path.basename(process.argv[1] || '').replace(/\.js$/, '');
  const direct = /^docloupe-(.+)-tools$/.exec(invoked);
  if (direct) {
    run(direct[1], process.argv.slice(2));
    return;
  }
  const [server, ...args] = process.argv.slice(2);
  run(server, args);
}

module.exports = { run, platformKey, executableName };

if (require.main === module) main();
