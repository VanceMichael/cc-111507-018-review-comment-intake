import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createService } from '../src/server.mjs';

const intake = async (service, body) => {
  const response = await fetch(`${service.baseUrl}/v1/intakes`, {
    method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body),
  });
  return { status: response.status, body: await response.json() };
};

test('结构化接收：三种文字视图各自保留同一批注范围', async () => {
  const dataDir = await mkdtemp(join(tmpdir(), 'structured-views-'));
  try {
    for (const view of ['accepted', 'original', 'marked']) {
      const service = await createService({ dataDir, scenario: 'standard' });
      try {
        const { status, body: receipt } = await intake(service, { mode: 'structured', view });
        assert.equal(status, 201);
        assert.equal(receipt.state, 'complete');
        assert.equal(receipt.view, view);
        const [record] = receipt.records;
        assert.equal(record.id, 'comment-0');
        assert.equal(record.text, '请复查参数');
        assert.equal(record.anchor.locatable, true);
        assert.equal(record.anchor.text, record.anchor.views[view].text);
        assert.equal(record.anchor.views.accepted.text, '新参数\n确认');
        assert.equal(record.anchor.views.original.text, '旧参数\n确认');
        assert.equal(receipt.coverage.comments.anchored, 1);
        assert.equal(receipt.evidence.sourceVersion, '"source-v1"');
        assert.equal(receipt.blobs[0].digest.length, 64);
        const saved = await (await fetch(`${service.baseUrl}/v1/receipts/${receipt.id}`)).json();
        assert.deepEqual(saved.records, receipt.records);
        assert.deepEqual(saved.issues, receipt.issues);
        assert.deepEqual(saved.coverage, receipt.coverage);
      } finally { await service.close(); }
    }
  } finally { await rm(dataDir, { recursive: true, force: true }); }
});

test('结构化接收：缺失范围端点形成 partial，不把意见挂到相同文字', async () => {
  const dataDir = await mkdtemp(join(tmpdir(), 'structured-orphan-'));
  const service = await createService({ dataDir, scenario: 'missing-anchor' });
  try {
    const { status, body: receipt } = await intake(service, { mode: 'structured', view: 'accepted' });
    assert.equal(status, 201);
    assert.equal(receipt.state, 'partial');
    assert.equal(receipt.records.length, 1);
    const [record] = receipt.records;
    assert.equal(record.anchor.locatable, false);
    assert.equal(record.anchor.views, null);
    assert.match(record.anchor.reason, /范围终点/);
    assert.ok(receipt.issues.some(issue => issue.code === 'anchor-range-incomplete'));
    assert.equal(receipt.coverage.comments.unlocated, 1);
    assert.equal(receipt.coverage.structured, true);
  } finally { await service.close(); await rm(dataDir, { recursive: true, force: true }); }
});

test('结构化接收：回复关系不依赖批注顺序，并继承上级范围', async () => {
  const dataDir = await mkdtemp(join(tmpdir(), 'structured-replies-'));
  const service = await createService({ dataDir, scenario: 'reply-chain' });
  try {
    const { body: receipt } = await intake(service, { mode: 'structured', view: 'marked' });
    assert.equal(receipt.state, 'complete');
    const root = receipt.records.find(record => record.commentId === '0');
    const reply = receipt.records.find(record => record.commentId === '1');
    assert.equal(root.text, '请复查参数');
    assert.equal(reply.text, '已复核');
    assert.equal(reply.replyTo, '0');
    assert.equal(reply.threadRoot, '0');
    assert.equal(reply.anchor.inherited, true);
    assert.deepEqual(reply.anchor.range, root.anchor.range);
    assert.equal(receipt.coverage.parts.commentsExtended, 'word/commentsExtended.xml');
    assert.equal(receipt.coverage.comments.replies, 1);
  } finally { await service.close(); await rm(dataDir, { recursive: true, force: true }); }
});

test('结构化接收结果在新实例中可查询，原文摘要保持不变', async () => {
  const dataDir = await mkdtemp(join(tmpdir(), 'structured-persist-'));
  let id;
  const first = await createService({ dataDir, scenario: 'original-view' });
  try {
    const { body: receipt } = await intake(first, { mode: 'structured', view: 'original' });
    id = receipt.id;
    assert.equal(receipt.records[0].anchor.text, '旧参数\n确认');
  } finally { await first.close(); }
  const second = await createService({ dataDir, scenario: 'original-view' });
  try {
    const response = await fetch(`${second.baseUrl}/v1/receipts/${id}`);
    assert.equal(response.status, 200);
    const saved = await response.json();
    assert.equal(saved.mode, 'structured');
    assert.equal(saved.view, 'original');
    assert.equal(saved.records[0].anchor.views.accepted.text, '新参数\n确认');
    assert.equal(saved.records[0].anchor.views.original.text, '旧参数\n确认');
  } finally { await second.close(); await rm(dataDir, { recursive: true, force: true }); }
});

test('结构化接收拒绝未知文字视图，原始接收不受影响', async () => {
  const dataDir = await mkdtemp(join(tmpdir(), 'structured-invalid-'));
  const service = await createService({ dataDir });
  try {
    const invalid = await intake(service, { mode: 'structured', view: 'final' });
    assert.equal(invalid.status, 400);
    assert.match(invalid.body.error, /文字视图无效/);
    assert.equal((await fetch(`${service.baseUrl}/v1/receipts`).then(response => response.json())).length, 0);
    const raw = await intake(service, { mode: 'raw' });
    assert.equal(raw.status, 201);
    assert.equal(raw.body.state, 'received');
    assert.equal(raw.body.coverage.structured, false);
    assert.equal(raw.body.records.length, 0);
  } finally { await service.close(); await rm(dataDir, { recursive: true, force: true }); }
});
