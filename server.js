import fs from 'node:fs/promises';
import path from 'node:path';
import { createStaticServer } from './scripts/static-server.js';
import { publicRoot, distRoot } from './scripts/paths.js';

const preview = process.argv.includes('--preview');
const port = Number(process.env.PORT || 5188);
if (!Number.isInteger(port) || port < 1 || port > 65535) {
  console.error('PORT 必须是 1–65535 之间的整数。');
  process.exit(1);
}
try {
  await fs.access(path.join(preview ? distRoot : publicRoot, 'index.html'));
} catch {
  console.error(preview ? '未找到构建产物，请先运行 npm run build。' : '未找到 public/index.html。');
  process.exit(1);
}

const server = createStaticServer({ root: preview ? distRoot : publicRoot });
server.on('error', error => {
  console.error(error.code === 'EADDRINUSE' ? `端口 ${port} 已被占用，请设置 PORT 后重试。` : error.message);
  process.exitCode = 1;
});
server.listen(port, '127.0.0.1', () => {
  console.log(`任务管理${preview ? '（构建预览）' : ''}：http://localhost:${port}`);
  console.log('按 Ctrl+C 停止服务。');
});
for (const signal of ['SIGINT', 'SIGTERM']) process.on(signal, () => server.close());
