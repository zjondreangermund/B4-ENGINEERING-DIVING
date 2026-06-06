import Fastify from 'fastify';
import fastifyStatic from '@fastify/static';
import { fileURLToPath } from 'url';
import path from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const app = Fastify({ logger: true });
const port = Number(process.env.PORT || 3000);
const host = '0.0.0.0';

await app.register(fastifyStatic, {
  root: path.join(__dirname, 'public'),
  prefix: '/',
});

app.get('/health', async () => ({ status: 'ok', app: 'B4 Nautilus Operations' }));

app.setNotFoundHandler((request, reply) => {
  reply.sendFile('index.html');
});

try {
  await app.listen({ port, host });
} catch (err) {
  app.log.error(err);
  process.exit(1);
}
