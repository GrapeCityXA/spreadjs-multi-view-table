import fs from 'node:fs/promises';
import path from 'node:path';
import { projectRoot, publicRoot, distRoot } from './paths.js';

// Only the project's fixed dist directory may be removed.
if (path.dirname(distRoot) !== path.resolve(projectRoot) || path.basename(distRoot) !== 'dist') {
  throw new Error('构建目录必须是当前项目的 dist 子目录');
}
await fs.rm(distRoot, { recursive: true, force: true });
await fs.cp(publicRoot, distRoot, { recursive: true });
console.log('构建完成：dist/（包含页面与应用代码）');
console.log('SpreadJS 和 Inter 字体均通过 CDN 加载，不再打包进 dist/。');
console.log('运行 npm run preview 预览；也可将 dist/ 交给静态 Web 服务器托管。');
