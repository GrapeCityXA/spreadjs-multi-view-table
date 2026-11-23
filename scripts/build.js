import fs from 'node:fs/promises';
import path from 'node:path';
import { projectRoot, publicRoot, distRoot, vendorAssets } from './paths.js';

// Check dependencies before replacing a previous build.
for (const source of vendorAssets.values()) await fs.access(source);

// Only the project's fixed dist directory may be removed.
if (path.dirname(distRoot) !== path.resolve(projectRoot) || path.basename(distRoot) !== 'dist') {
  throw new Error('构建目录必须是当前项目的 dist 子目录');
}
await fs.rm(distRoot, { recursive: true, force: true });
await fs.cp(publicRoot, distRoot, { recursive: true });
for (const [url, source] of vendorAssets) {
  const destination = path.join(distRoot, url.slice(1));
  await fs.mkdir(path.dirname(destination), { recursive: true });
  await fs.copyFile(source, destination);
}
console.log('构建完成：dist/（包含页面、应用代码及 SpreadJS 静态资源）');
console.log('运行 npm run preview 预览；也可将 dist/ 交给静态 Web 服务器托管。');
