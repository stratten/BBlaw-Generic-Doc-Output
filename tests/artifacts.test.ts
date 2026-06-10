import assert from 'node:assert/strict';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { rm } from 'node:fs/promises';
import http from 'node:http';
import test from 'node:test';

const storageDir = path.join(tmpdir(), `bblaw-doc-artifacts-test-${process.pid}`);
process.env.ARTIFACT_STORAGE_DIR = storageDir;
process.env.ARTIFACT_TTL_DAYS = '1';

test.after(async () => {
  await rm(storageDir, { recursive: true, force: true });
});

test('normalizes markdown into document blocks', async () => {
  const { normalizeDocument } = await import('../src/services/renderDocument');
  const document = normalizeDocument({
    title: 'Status Memo',
    markdown: '# Summary\n\n- First point\n- Second point\n\n| Field | Value |\n| --- | --- |\n| Status | Ready |',
    formats: ['docx'],
    style: 'memo',
  });

  assert.equal(document.title, 'Status Memo');
  assert.equal(document.style, 'memo');
  assert.deepEqual(document.blocks.map((block) => block.type), ['heading', 'list', 'table']);
});

test('creates stored artifact metadata and retrievable DOCX file', async () => {
  const { renderDocument } = await import('../src/services/renderDocument');
  const { createStoredArtifact, getStoredArtifact, getStoredArtifactFile, publicArtifactMetadata } = await import('../src/services/artifactStore');
  const rendered = await renderDocument({
    title: 'Download Test',
    markdown: 'This is a short generated artifact.',
    formats: ['docx'],
    style: 'plain',
  });

  const metadata = await createStoredArtifact({
    title: rendered.document.title,
    style: rendered.document.style,
    files: rendered.files,
    creator: { id: 'user-1', email: 'user@example.com' },
  });
  const loaded = await getStoredArtifact(metadata.artifactId);
  const file = await getStoredArtifactFile(metadata.artifactId, metadata.files[0].fileId);
  const publicMetadata = publicArtifactMetadata(metadata);

  assert.equal(loaded?.artifactId, metadata.artifactId);
  assert.equal(file?.metadata.contentType, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  assert.ok(file && file.buffer.byteLength > 1000);
  assert.equal('storagePath' in publicMetadata.files[0], false);
});

function request(server: http.Server, input: { method: string; path: string; body?: unknown; headers?: Record<string, string> }) {
  const address = server.address();
  assert.ok(address && typeof address === 'object');
  return new Promise<{ status: number; body: string; headers: http.IncomingHttpHeaders }>((resolve, reject) => {
    const body = input.body ? JSON.stringify(input.body) : undefined;
    const req = http.request({
      method: input.method,
      hostname: '127.0.0.1',
      port: address.port,
      path: input.path,
      headers: {
        ...(body ? { 'content-type': 'application/json', 'content-length': Buffer.byteLength(body).toString() } : {}),
        ...input.headers,
      },
    }, (res) => {
      const chunks: Buffer[] = [];
      res.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
      res.on('end', () => resolve({
        status: res.statusCode || 0,
        body: Buffer.concat(chunks).toString('utf8'),
        headers: res.headers,
      }));
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

test('artifact route is capability gated and returns download metadata', async () => {
  const { createApp } = await import('../src/app');
  const server = createApp().listen(0);
  try {
    const body = {
      title: 'Route Test',
      markdown: 'Generated from the route test.',
      formats: ['docx'],
      style: 'report',
    };
    const forbidden = await request(server, { method: 'POST', path: '/artifacts', body });
    assert.equal(forbidden.status, 403);

    const created = await request(server, {
      method: 'POST',
      path: '/artifacts',
      body,
      headers: { 'x-gateway-broker': '1', 'x-portal-user-id': 'user-1', 'x-portal-user-email': 'user@example.com' },
    });
    assert.equal(created.status, 201);
    const metadata = JSON.parse(created.body);
    assert.equal(metadata.title, 'Route Test');
    assert.equal(metadata.files.length, 1);
    assert.match(metadata.files[0].downloadUrl, /^\/artifacts\//);
  } finally {
    await new Promise<void>((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
  }
});
