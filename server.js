import fs from 'node:fs/promises';
import path from 'node:path';
import { createStaticServer } from './scripts/static-server.js';
import { publicRoot, distRoot, vendorAssets } from './scripts/paths.js';

const preview = process.argv.includes('--preview');
const port = Number(process.env.PORT || 5188);
if (!Number.isInteger(port) || port < 1 || port > 65535) {
  console.error('PORT 必须是 1–65535 之间的整数。');
  process.exit(1);
}
try {
  if (preview) await fs.access(path.join(distRoot, 'index.html'));
  else for (const asset of vendorAssets.values()) await fs.access(asset);
} catch {
  console.error(preview ? '未找到构建产物，请先运行 npm run build。' : '依赖不完整，请先运行 npm ci。');
  process.exit(1);
}

const server = createStaticServer({ root: preview ? distRoot : publicRoot, assets: preview ? undefined : vendorAssets });
server.on('error', error => {
  console.error(error.code === 'EADDRINUSE' ? `端口 ${port} 已被占用，请设置 PORT 后重试。` : error.message);
  process.exitCode = 1;
});
server.listen(port, '127.0.0.1', () => {
  console.log(`任务管理${preview ? '（构建预览）' : ''}：http://localhost:${port}`);
  console.log('按 Ctrl+C 停止服务。');
});
for (const signal of ['SIGINT', 'SIGTERM']) process.on(signal, () => server.close());
