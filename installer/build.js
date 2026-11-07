'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { preparePackage, validatePackage } = require('./package');

function run(command, args, options = {}) {
  const result = spawnSync(command, args, { stdio: 'inherit', ...options });
  if (result.error || result.status !== 0) throw new Error(`${path.basename(command)} failed: ${result.error?.message || `exit ${result.status}`}`);
}

function main() {
  if (process.platform !== 'win32' || process.arch !== 'x64') throw new Error('Build the Windows x64 installer on a Windows x64 machine or GitHub Actions.');
  const root = path.resolve(__dirname, '..');
  const metadata = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  if (!/^\d+\.\d+\.\d+$/.test(metadata.version)) throw new Error('Installer version must have three numeric parts.');
  const stage = preparePackage(root);
  const csc = path.join(process.env.WINDIR || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');
  if (!fs.existsSync(csc)) throw new Error('The Windows .NET Framework C# compiler was not found.');
  const assembly = path.join(root, 'dist', 'launcher-version.cs');
  fs.writeFileSync(assembly, `using System.Reflection;\n[assembly: AssemblyTitle("PHDTech Commission Tool")]\n[assembly: AssemblyVersion("${metadata.version}.0")]\n`);
  run(csc, ['/nologo', '/target:winexe', '/platform:x64', '/reference:System.Windows.Forms.dll',
    `/out:${path.join(stage, 'PHDTech Commission Tool.exe')}`, path.join(__dirname, 'launcher.cs'), assembly]);
  const node = path.join(stage, 'runtime', 'node.exe');
  const npm = process.env.npm_execpath || path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js');
  if (!fs.existsSync(npm)) throw new Error('Run the build through npm run build:installer so the npm CLI can be located.');
  run(node, [npm, 'ci', '--prefix', stage, '--omit=dev', '--ignore-scripts', '--no-audit', '--no-fund'], { cwd: stage });
  const env = { ...process.env, PLAYWRIGHT_BROWSERS_PATH: path.join(stage, 'browsers'), PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD: '0' };
  run(node, [path.join(stage, 'node_modules', 'playwright', 'cli.js'), 'install', 'chromium', '--no-shell'], { cwd: stage, env });
  validatePackage(stage);
  run(node, [path.join(__dirname, 'smoke-package.js'), stage], { cwd: stage, env });
  validatePackage(stage);
  const compiler = [process.env['ProgramFiles(x86)'], process.env.ProgramFiles]
    .filter(Boolean).map(base => path.join(base, 'Inno Setup 6', 'ISCC.exe')).find(file => fs.existsSync(file));
  if (!compiler) throw new Error('Inno Setup 6 was not found on the build machine.');
  run(compiler, [`/DAppVersion=${metadata.version}`, `/DPackageDir=${stage}`, `/O${path.join(root, 'dist')}`, path.join(__dirname, 'setup.iss')]);
}

if (require.main === module) {
  try { main(); } catch (error) { console.error(error.stack || error.message); process.exitCode = 1; }
}
