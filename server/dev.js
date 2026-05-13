import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const rootDir = resolve(__dirname, '..');
const viteBin = resolve(rootDir, 'node_modules', 'vite', 'bin', 'vite.js');

const processes = [
  {
    name: 'proxy',
    command: process.execPath,
    args: [resolve(rootDir, 'server', 'productoHoraProxy.js')],
  },
  {
    name: 'vite',
    command: process.execPath,
    args: [viteBin, '--host', '0.0.0.0'],
  },
];

const children = processes.map(({ name, command, args }) => {
  const child = spawn(command, args, {
    cwd: rootDir,
    env: process.env,
    stdio: 'inherit',
  });

  child.on('exit', (code, signal) => {
    if (code !== 0 && signal !== 'SIGTERM') {
      console.error(`${name} finalizo con codigo ${code ?? signal}`);
      shutdown(code ?? 1);
    }
  });

  return child;
});

function shutdown(code = 0) {
  children.forEach((child) => {
    if (!child.killed) {
      child.kill('SIGTERM');
    }
  });
  process.exit(typeof code === 'number' ? code : 0);
}

process.on('SIGINT', () => shutdown(0));
process.on('SIGTERM', () => shutdown(0));
