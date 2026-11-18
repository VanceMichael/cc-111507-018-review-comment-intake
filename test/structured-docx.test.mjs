import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { zipSync, strToU8 } from 'fflate';
import { createService } from '../src/server.mjs';
import { parseReviewPackage, buildRecords } from '../src/docx-review.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

async function withService(scenario, run) {
  const dataDir = await mkdtemp(join(tmpdir(), 'structured-docx-'));
  const service = await createService({ dataDir, scenario });
  try {
    await run(service);
  } finally {
    await service.close();
    await rm(dataDir, { recursive: true, force: true });
  }
}

async function intake(service, body) {
  const response = await fetch(`${service.baseUrl}/v1/intakes`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(body)
  });
  return { status: response.status, body: await response.json() };
}

function docx({ document, comments, commentsExtended, extra = {} }) {
  const parts = { 'word/document.xml': strToU8(document) };
  if (comments) parts['word/comments.xml'] = strToU8(comments);
  if (commentsExtended) parts['word/commentsExtended.xml'] = strToU8(commentsExtended);
  for (const [name, text] of Object.entries(extra)) parts[name] = strToU8(text);
  return zipSync(parts);
}

test('结构化收件：accepted 视图覆盖文字保留插入、忽略删除，跨段锚点保留原始标记位置', async () => {
  await withService('standard', async (service) => {
    const { status, body } = await intake(service, { mode: 'structured', view: 'accepted' });
    assert.equal(status, 201);
    assert.equal(body.state, 'complete');
    assert.deepEqual(body.issues, []);
    assert.equal(body.records.length, 1);
    const [record] = body.records;
    assert.equal(record.commentId, '0');
    assert.equal(record.author, '合成审阅人');
    assert.equal(record.text, '请复查参数');
    assert.equal(record.parentId, null);
    assert.equal(record.view, 'accepted');
    assert.equal(record.coveredText, '新参数\n确认');
    assert.equal(record.anchor.start.marker, 'w:document/w:body[0]/w:p[0]/w:commentRangeStart[0]');
    assert.equal(record.anchor.start.path, 'w:document/w:body[0]/w:p[0]/w:r[0]/w:t[0]');
    assert.equal(record.anchor.start.offset, 2);
    assert.equal(record.anchor.end.marker, 'w:document/w:body[0]/w:p[1]/w:commentRangeEnd[0]');
    assert.equal(record.anchor.end.path, 'w:document/w:body[0]/w:p[1]/w:r[0]/w:t[0]');
    assert.equal(record.anchor.end.offset, 2);
    assert.equal(record.anchor.inheritedFrom, null);
    assert.deepEqual(record.views.accepted, {
      start: 2,
      end: 8,
      text: '新参数\n确认',
      segments: [
        { text: '新参数', status: 'inserted' },
        { text: '\n', status: 'normal' },
        { text: '确认', status: 'normal' }
      ]
    });
    assert.equal(record.views.original.text, '旧参数\n确认');
    assert.equal(record.views.marked.text, '旧参数新参数\n确认');
    assert.deepEqual(record.views.marked.segments.map((segment) => segment.status), ['deleted', 'inserted', 'normal', 'normal']);
    assert.equal(body.coverage.structured, true);
    assert.equal(body.coverage.view, 'accepted');
    assert.equal(body.coverage.comments, 1);
    assert.equal(body.coverage.anchored, 1);
    assert.equal(body.coverage.complete, true);
    assert.deepEqual(body.coverage.unsupportedParts, []);
    assert.equal(body.evidence.parts.find((part) => part.path === 'word/comments.xml').via, 'relationship');
    const saved = await (await fetch(`${service.baseUrl}/v1/receipts/${body.id}`)).json();
    assert.deepEqual(saved.records, body.records);
    assert.deepEqual(saved.coverage, body.coverage);
  });
});

test('结构化收件：original 视图保留删除文字，marked 视图同时保留并标注状态', async () => {
  await withService('original-view', async (service) => {
    const original = await intake(service, { mode: 'structured', view: 'original' });
    assert.equal(original.status, 201);
    assert.equal(original.body.records[0].coveredText, '旧参数\n确认');
    assert.deepEqual(original.body.records[0].views.original.segments.map((segment) => segment.status), ['deleted', 'normal', 'normal']);
    const marked = await intake(service, { mode: 'structured' });
    assert.equal(marked.status, 201);
    assert.equal(marked.body.coverage.view, 'marked');
    assert.equal(marked.body.records[0].coveredText, '旧参数新参数\n确认');
  });
});

test('结构化收件：缺少范围结束标记时标明无法定位，不靠文本匹配挂到相同文字', async () => {
  await withService('missing-anchor', async (service) => {
    const { status, body } = await intake(service, { mode: 'structured', view: 'accepted' });
    assert.equal(status, 201);
    assert.equal(body.state, 'partial');
    assert.equal(body.records.length, 1);
    const [record] = body.records;
    assert.equal(record.commentId, '0');
    assert.equal(record.anchor, null);
    assert.equal(record.coveredText, null);
    assert.equal(record.views, null);
    assert.ok(body.issues.some((issue) => issue.kind === 'range-endpoint-missing' && issue.commentId === '0'));
    assert.equal(body.coverage.complete, false);
    const saved = await (await fetch(`${service.baseUrl}/v1/receipts/${body.id}`)).json();
    assert.deepEqual(saved.records, body.records);
  });
});

test('结构化收件：回复关系经 commentsExtended 建立，回复继承父批注锚点', async () => {
  await withService('reply-chain', async (service) => {
    const { status, body } = await intake(service, { mode: 'structured', view: 'accepted' });
    assert.equal(status, 201);
    assert.equal(body.state, 'complete');
    assert.equal(body.records.length, 2);
    const parent = body.records.find((record) => record.commentId === '0');
    const reply = body.records.find((record) => record.commentId === '1');
    assert.equal(parent.parentId, null);
    assert.equal(parent.anchor.inheritedFrom, null);
    assert.equal(reply.parentId, '0');
    assert.equal(reply.text, '已复核');
    assert.equal(reply.anchor.inheritedFrom, '0');
    assert.equal(reply.coveredText, parent.coveredText);
    assert.deepEqual(reply.views, parent.views);
    assert.equal(body.coverage.replies, 1);
  });
});

test('结构化收件：无法解释的来源形成 rejected 收件并保留原文证据', async () => {
  await withService('corrupt-entry', async (service) => {
    const { status, body } = await intake(service, { mode: 'structured', view: 'accepted' });
    assert.equal(status, 201);
    assert.equal(body.state, 'rejected');
    assert.deepEqual(body.records, []);
    assert.equal(body.issues[0].kind, 'package-unreadable');
    assert.equal(body.coverage.complete, false);
    assert.equal(body.evidence.sha256, body.blobs[0].digest);
    const saved = await (await fetch(`${service.baseUrl}/v1/receipts/${body.id}`)).json();
    assert.equal(saved.state, 'rejected');
  });
});

test('结构化收件：无效视图返回 400，重新打开数据目录仍可查询', async () => {
  const dataDir = await mkdtemp(join(tmpdir(), 'structured-docx-'));
  let service = await createService({ dataDir, scenario: 'standard' });
  try {
    const invalid = await intake(service, { mode: 'structured', view: 'sideways' });
    assert.equal(invalid.status, 400);
    const { body } = await intake(service, { mode: 'structured', view: 'marked' });
    assert.equal(body.state, 'complete');
    await service.close();
    service = await createService({ dataDir, scenario: 'standard' });
    const saved = await (await fetch(`${service.baseUrl}/v1/receipts/${body.id}`)).json();
    assert.deepEqual(saved.records, body.records);
    assert.deepEqual(saved.coverage, body.coverage);
  } finally {
    await service.close();
    await rm(dataDir, { recursive: true, force: true });
  }
});

test('解析：非 ZIP 字节与缺少正文的组合判为无法解释', () => {
  assert.equal(parseReviewPackage(Buffer.from('junk')).status, 'unreadable');
  const noDocument = zipSync({ 'word/comments.xml': strToU8(`<w:comments xmlns:w="${W}"/>`) });
  const missing = parseReviewPackage(noDocument);
  assert.equal(missing.status, 'unreadable');
  assert.equal(missing.issues[0].kind, 'package-unreadable');
  const malformed = docx({ document: '<w:document' });
  assert.equal(parseReviewPackage(malformed).status, 'unreadable');
});

test('解析：相同文字重复出现时缺端点仍判无法定位，不做文本匹配', () => {
  const model = parseReviewPackage(docx({
    document: `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>相同</w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>相同</w:t></w:r></w:p></w:body></w:document>`,
    comments: `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="甲"><w:p><w:r><w:t>意见</w:t></w:r></w:p></w:comment></w:comments>`
  }));
  assert.equal(model.status, 'parsed');
  assert.ok(model.issues.some((issue) => issue.kind === 'range-endpoint-missing'));
  const [record] = buildRecords(model, 'marked');
  assert.equal(record.anchor, null);
  assert.equal(record.coveredText, null);
});

test('解析：端点顺序颠倒与重复端点都判无法定位', () => {
  const unordered = parseReviewPackage(docx({
    document: `<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeEnd w:id="0"/><w:r><w:t>文字</w:t></w:r><w:commentRangeStart w:id="0"/></w:p></w:body></w:document>`,
    comments: `<w:comments xmlns:w="${W}"><w:comment w:id="0"><w:p><w:r><w:t>意见</w:t></w:r></w:p></w:comment></w:comments>`
  }));
  assert.ok(unordered.issues.some((issue) => issue.kind === 'range-endpoints-unordered'));
  const duplicated = parseReviewPackage(docx({
    document: `<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>文字</w:t></w:r><w:commentRangeEnd w:id="0"/><w:commentRangeEnd w:id="0"/></w:p></w:body></w:document>`,
    comments: `<w:comments xmlns:w="${W}"><w:comment w:id="0"><w:p><w:r><w:t>意见</w:t></w:r></w:p></w:comment></w:comments>`
  }));
  assert.ok(duplicated.issues.some((issue) => issue.kind === 'range-endpoint-duplicate'));
  assert.equal(buildRecords(duplicated, 'marked')[0].anchor, null);
});

test('解析：点范围覆盖文字为空但仍可定位', () => {
  const model = parseReviewPackage(docx({
    document: `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>甲</w:t></w:r><w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/><w:r><w:t>乙</w:t></w:r></w:p></w:body></w:document>`,
    comments: `<w:comments xmlns:w="${W}"><w:comment w:id="0"><w:p><w:r><w:t>意见</w:t></w:r></w:p></w:comment></w:comments>`
  }));
  assert.deepEqual(model.issues, []);
  const [record] = buildRecords(model, 'accepted');
  assert.equal(record.coveredText, '');
  assert.equal(record.views.accepted.start, 1);
  assert.equal(record.views.accepted.end, 1);
});

test('解析：回复父级缺失、关系外目标与回复循环都标明关系不完整', () => {
  const comments = `<w:comments xmlns:w="${W}"><w:comment w:id="0"><w:p xmlns:w14="${W14}" w14:paraId="00000001"><w:r><w:t>甲</w:t></w:r></w:p></w:comment><w:comment w:id="1"><w:p xmlns:w14="${W14}" w14:paraId="00000002"><w:r><w:t>乙</w:t></w:r></w:p></w:comment></w:comments>`;
  const document = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p></w:body></w:document>`;
  const missingParent = parseReviewPackage(docx({
    document,
    comments,
    commentsExtended: `<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="00000001" w15:paraIdParent="00000009"/></w15:commentsEx>`
  }));
  assert.ok(missingParent.issues.some((issue) => issue.kind === 'reply-parent-unknown'));
  const outsideTarget = parseReviewPackage(docx({
    document,
    comments,
    commentsExtended: `<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="00000009" w15:paraIdParent="00000001"/></w15:commentsEx>`
  }));
  assert.ok(outsideTarget.issues.some((issue) => issue.kind === 'reply-target-unknown'));
  const cycle = parseReviewPackage(docx({
    document,
    comments,
    commentsExtended: `<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="00000001" w15:paraIdParent="00000002"/><w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000001"/></w15:commentsEx>`
  }));
  assert.ok(cycle.issues.some((issue) => issue.kind === 'reply-cycle'));
  assert.equal(buildRecords(cycle, 'marked').every((record) => record.anchor === null), true);
});

test('解析：主命名空间按 URI 识别，不依赖前缀写法', () => {
  const model = parseReviewPackage(docx({
    document: `<d:document xmlns:d="${W}"><d:body><d:p><d:commentRangeStart d:id="0"/><d:r><d:t>正文</d:t></d:r><d:commentRangeEnd d:id="0"/></d:p></d:body></d:document>`,
    comments: `<c:comments xmlns:c="${W}"><c:comment c:id="0"><c:p><c:r><c:t>意见</c:t></c:r></c:p></c:comment></c:comments>`
  }));
  assert.equal(model.status, 'parsed');
  assert.deepEqual(model.issues, []);
  assert.equal(buildRecords(model, 'marked')[0].coveredText, '正文');
});

test('解析：未列出的扩展部分保留原文并明确不支持', () => {
  const model = parseReviewPackage(docx({
    document: `<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>正文</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p></w:body></w:document>`,
    comments: `<w:comments xmlns:w="${W}"><w:comment w:id="0"><w:p><w:r><w:t>意见</w:t></w:r></w:p></w:comment></w:comments>`,
    extra: { 'word/commentsIds.xml': '<w16cid:commentsIds xmlns:w16cid="urn:x"/>' }
  }));
  assert.deepEqual(model.unsupportedParts, ['word/commentsIds.xml']);
  assert.ok(model.issues.some((issue) => issue.kind === 'part-unsupported' && issue.part === 'word/commentsIds.xml'));
  assert.equal(buildRecords(model, 'marked')[0].coveredText, '正文');
});
