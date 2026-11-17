import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { zipSync } from 'fflate';
import { extractReview } from '../src/wordprocessingml.mjs';

const fixture = name => readFile(new URL(`../fixtures/${name}.docx`, import.meta.url));
const encode = value => new TextEncoder().encode(value);

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const COMMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';

function packageOf({ document, comments, extended, relationships, extra = {} }) {
  const files = {};
  if (document !== undefined) files['word/document.xml'] = encode(document);
  if (comments !== undefined) files['word/comments.xml'] = encode(comments);
  if (extended !== undefined) files['word/commentsExtended.xml'] = encode(extended);
  if (relationships !== undefined) files['word/_rels/document.xml.rels'] = encode(relationships);
  for (const [name, value] of Object.entries(extra)) files[name] = encode(value);
  return zipSync(files);
}

const rels = (entries) => `<Relationships xmlns="${REL_NS}">${entries.map(([type, target]) =>
  `<Relationship Id="r" Type="${type}" Target="${target}"/>`).join('')}</Relationships>`;

const w = (body) => `<w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document>`;
const comment = (id, text) => `<w:comment w:id="${id}" w:author="合成审阅人"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
const commentsXml = (...items) => `<w:comments xmlns:w="${W_NS}">${items.join('')}</w:comments>`;

const base = (document, comments = commentsXml(comment(0, '请复查参数'))) => packageOf({
  document, comments, relationships: rels([[COMMENT_REL, 'comments.xml']]),
});

const anchored = (inner) => w(`<w:p><w:r><w:t>前缀</w:t></w:r><w:commentRangeStart w:id="0"/>${inner}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);

test('审阅文字按视图投影，批注范围随可见文字映射', async () => {
  const bytes = await fixture('review');
  for (const view of ['accepted', 'original', 'marked']) {
    const result = extractReview(bytes, { view });
    assert.equal(result.rejected, false);
    assert.equal(result.state, 'complete');
    assert.equal(result.records.length, 1);
    const [record] = result.records;
    assert.equal(record.commentId, '0');
    assert.equal(record.text, '请复查参数');
    assert.equal(record.anchor.locatable, true);
    assert.deepEqual(record.anchor.paragraphs, { start: 0, end: 1 });
    assert.equal(record.anchor.views.accepted.text, '新参数\n确认');
    assert.equal(record.anchor.views.original.text, '旧参数\n确认');
    assert.equal(record.anchor.views.marked.text, '旧参数新参数\n确认');
    assert.equal(record.anchor.text, record.anchor.views[view].text);
    assert.equal(record.anchor.views.accepted.includes.deleted, false);
    assert.equal(record.anchor.views.accepted.includes.inserted, true);
    assert.equal(record.anchor.views.accepted.omits.deleted, true);
    assert.equal(record.anchor.views.original.includes.deleted, true);
    assert.equal(record.anchor.views.original.includes.inserted, false);
    assert.equal(record.anchor.views.original.omits.inserted, true);
    assert.equal(record.anchor.views.accepted.end - record.anchor.views.accepted.start, 6);
    assert.equal(record.anchor.views.original.end - record.anchor.views.original.start, 6);
    assert.equal(record.anchor.views.marked.end - record.anchor.views.marked.start, 9);
    assert.equal(record.anchor.range.length, 9);
  }
});

test('原始锚点位置按 run/text 节点路径保留，不受视图影响', async () => {
  const bytes = await fixture('review');
  const result = extractReview(bytes);
  assert.deepEqual(result.records[0].anchor.markers.starts, [{ nodePath: 'document/body[0]/p[0]/commentRangeStart[0]', offset: 2, paragraph: 0 }]);
  assert.deepEqual(result.records[0].anchor.markers.ends, [{ nodePath: 'document/body[0]/p[1]/commentRangeEnd[0]', offset: 11, paragraph: 1 }]);
  assert.deepEqual(result.records[0].anchor.markers.references, [{ nodePath: 'document/body[0]/p[1]/r[1]/commentReference[0]', offset: 11, paragraph: 1 }]);
  assert.equal(result.records[0].anchor.range.start, 2);
  assert.equal(result.records[0].anchor.range.end, 11);
});

test('插入和删除文字在两种文字视图下都解释得通', async () => {
  const result = extractReview(await fixture('review'), { view: 'marked' });
  const [record] = result.records;
  assert.notEqual(record.anchor.views.accepted.text, record.anchor.views.original.text);
  assert.ok(record.anchor.views.accepted.text.includes('新参数'));
  assert.ok(!record.anchor.views.accepted.text.includes('旧参数'));
  assert.ok(record.anchor.views.original.text.includes('旧参数'));
  assert.ok(!record.anchor.views.original.text.includes('新参数'));
  assert.ok(record.anchor.views.marked.text.includes('旧参数'));
  assert.ok(record.anchor.views.marked.text.includes('新参数'));
});

test('只覆盖插入或删除文字时，另一种视图给出空范围而非错挂到相邻文字', () => {
  const document = w(`<w:p><w:r><w:t>设备</w:t></w:r><w:ins><w:r><w:t>新参数</w:t></w:r></w:ins></w:p>` +
    `<w:p><w:r><w:t>确认</w:t></w:r></w:p>`);
  const insDocument = w(`<w:p><w:r><w:t>设备</w:t></w:r><w:commentRangeStart w:id="0"/><w:ins><w:r><w:t>新参数</w:t></w:r></w:ins><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const inserted = extractReview(base(insDocument));
  const [record] = inserted.records;
  assert.equal(record.anchor.views.accepted.text, '新参数');
  assert.equal(record.anchor.views.original.text, '');
  assert.equal(record.anchor.views.original.empty, true);
  assert.equal(record.anchor.text, '新参数');
  const original = extractReview(base(insDocument), { view: 'original' });
  assert.equal(original.records[0].anchor.text, '');
  assert.deepEqual(extractReview(base(document)).records[0].text, '请复查参数');
});

test('范围端点缺失时标明无法定位，不按相同的可见文字搜索', () => {
  const document = w(`<w:p><w:r><w:t>新参数</w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>确认</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>新参数</w:t></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const result = extractReview(base(document));
  const [record] = result.records;
  assert.equal(result.state, 'partial');
  assert.equal(record.anchor.locatable, false);
  assert.equal(record.anchor.views, null);
  assert.equal(record.anchor.text, null);
  assert.deepEqual(record.anchor.markers.starts.map(marker => marker.offset), [3]);
  assert.deepEqual(record.anchor.markers.ends, []);
  assert.ok(result.issues.some(issue => issue.code === 'anchor-range-incomplete' && issue.missing.includes('end')));
});

test('只有范围起止标记缺少批注引用时仍按标记定位，并记录问题', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>`);
  const result = extractReview(base(document));
  const [record] = result.records;
  assert.equal(record.anchor.locatable, true);
  assert.equal(record.anchor.text, '参数');
  assert.deepEqual(record.anchor.markers.references, []);
  assert.ok(result.issues.some(issue => issue.code === 'anchor-reference-missing'));
});

test('重复端点不猜测范围', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>甲</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>乙</w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>丙</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const result = extractReview(base(document));
  assert.equal(result.records[0].anchor.locatable, false);
  assert.ok(result.issues.some(issue => issue.code === 'anchor-marker-duplicated'));
});

test('端点顺序颠倒时标明无法定位', () => {
  const document = w(`<w:p><w:commentRangeEnd w:id="0"/><w:r><w:t>参数</w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const result = extractReview(base(document));
  assert.equal(result.records[0].anchor.locatable, false);
  assert.ok(result.issues.some(issue => issue.code === 'anchor-order-invalid'));
});

test('跨段批注覆盖段间换行并保留两端段落', () => {
  const document = w(`<w:p><w:r><w:t>第一段</w:t></w:r><w:commentRangeStart w:id="0"/></w:p>` +
    `<w:p><w:r><w:t>第二段</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const result = extractReview(base(document));
  const [record] = result.records;
  assert.equal(record.anchor.locatable, true);
  assert.equal(record.anchor.text, '\n第二段');
  assert.deepEqual(record.anchor.paragraphs, { start: 0, end: 1 });
});

test('Unicode code point 偏移按字符而非无头代理项计算', () => {
  const document = w(`<w:p><w:r><w:t>😀评价</w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const result = extractReview(base(document));
  const [record] = result.records;
  assert.equal(record.anchor.range.start, 3);
  assert.equal(record.anchor.range.end, 5);
  assert.equal(record.anchor.text, '参数');
});

test('回复关系按 commentsExtended paraId 建立，不依赖批注 id 顺序', async () => {
  const bytes = await fixture('replies');
  const result = extractReview(bytes, { view: 'marked' });
  assert.equal(result.state, 'complete');
  const reply = result.records.find(record => record.commentId === '1');
  assert.equal(reply.replyTo, '0');
  assert.equal(reply.threadRoot, '0');
  assert.equal(reply.anchor.inherited, true);
  assert.equal(reply.anchor.parentId, '0');
  assert.equal(reply.anchor.text, result.records[0].anchor.text);
  assert.equal(reply.anchor.locatable, true);
});

test('回复链可以跨多层并解析到同一范围', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const comments = commentsXml(
    `<w:comment w:id="0" w:author="甲"><w:p xmlns:w14="${W14_NS}" w14:paraId="0000000A"><w:r><w:t>根意见</w:t></w:r></w:p></w:comment>`,
    `<w:comment w:id="1" w:author="乙"><w:p xmlns:w14="${W14_NS}" w14:paraId="0000000B"><w:r><w:t>回复</w:t></w:r></w:p></w:comment>`,
    `<w:comment w:id="2" w:author="丙"><w:p xmlns:w14="${W14_NS}" w14:paraId="0000000C"><w:r><w:t>再回复</w:t></w:r></w:p></w:comment>`,
  );
  const extended = `<w15:commentsEx xmlns:w15="${W15_NS}"><w15:commentEx w15:paraId="0000000A"/>` +
    `<w15:commentEx w15:paraId="0000000B" w15:paraIdParent="0000000A"/>` +
    `<w15:commentEx w15:paraId="0000000C" w15:paraIdParent="0000000B"/></w15:commentsEx>`;
  const bytes = packageOf({ document, comments, extended, relationships: rels([[COMMENT_REL, 'comments.xml'], [EXTENDED_REL, 'commentsExtended.xml']]) });
  const result = extractReview(bytes);
  assert.equal(result.state, 'complete');
  const [root, first, second] = result.records;
  assert.equal(first.replyTo, '0');
  assert.equal(second.replyTo, '1');
  assert.equal(second.threadRoot, '0');
  assert.equal(first.anchor.inherited, true);
  assert.equal(second.anchor.inherited, true);
  assert.equal(second.anchor.text, '参数');
});

test('回复目标缺失时不把回复挂到别处', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const comments = commentsXml(
    `<w:comment w:id="0" w:author="甲"><w:p xmlns:w14="${W14_NS}" w14:paraId="00000001"><w:r><w:t>根意见</w:t></w:r></w:p></w:comment>`,
    `<w:comment w:id="1" w:author="乙"><w:p xmlns:w14="${W14_NS}" w14:paraId="00000002"><w:r><w:t>回复</w:t></w:r></w:p></w:comment>`,
  );
  const extended = `<w15:commentsEx xmlns:w15="${W15_NS}"><w15:commentEx w15:paraId="00000001"/>` +
    `<w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000009"/></w15:commentsEx>`;
  const bytes = packageOf({ document, comments, extended, relationships: rels([[COMMENT_REL, 'comments.xml'], [EXTENDED_REL, 'commentsExtended.xml']]) });
  const result = extractReview(bytes);
  assert.equal(result.state, 'partial');
  const reply = result.records.find(record => record.commentId === '1');
  assert.equal(reply.replyTo, null);
  assert.equal(reply.anchor.locatable, false);
  assert.ok(result.issues.some(issue => issue.code === 'reply-target-missing' && issue.commentId === '1'));
});

test('回复关系循环不按可见文字猜测归属', () => {
  const document = w(`<w:p><w:r><w:t>参数</w:t></w:r></w:p>`);
  const comments = commentsXml(
    `<w:comment w:id="0" w:author="甲"><w:p xmlns:w14="${W14_NS}" w14:paraId="00000001"><w:r><w:t>甲</w:t></w:r></w:p></w:comment>`,
    `<w:comment w:id="1" w:author="乙"><w:p xmlns:w14="${W14_NS}" w14:paraId="00000002"><w:r><w:t>乙</w:t></w:r></w:p></w:comment>`,
  );
  const extended = `<w15:commentsEx xmlns:w15="${W15_NS}"><w15:commentEx w15:paraId="00000001" w15:paraIdParent="00000002"/>` +
    `<w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000001"/></w15:commentsEx>`;
  const bytes = packageOf({ document, comments, extended, relationships: rels([[COMMENT_REL, 'comments.xml'], [EXTENDED_REL, 'commentsExtended.xml']]) });
  const result = extractReview(bytes);
  assert.equal(result.state, 'partial');
  assert.equal(result.issues.filter(issue => issue.code === 'reply-cycle').length, 2);
  assert.ok(result.records.every(record => record.replyTo === null));
});

test('批注部件命名空间不在主命名空间时拒绝，不按本地名猜测', () => {
  const document = w(`<w:p><w:r><w:t>参数</w:t></w:r></w:p>`);
  const comments = `<c:comments xmlns:c="http://example.invalid/other">${comment(0, '别处批注').replace(/^<w:comment/, '<c:comment').replace(/<\/w:comment>$/, '</c:comment>')}</c:comments>`;
  const bytes = packageOf({ document, comments, relationships: rels([[COMMENT_REL, 'comments.xml']]) });
  const result = extractReview(bytes);
  assert.equal(result.rejected, true);
  assert.ok(result.issues.some(issue => issue.code === 'comments-unreadable'));
});

test('锚点指向不存在的批注时记录问题，不把意见挂到相同文字', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>`);
  const result = extractReview(base(document));
  assert.equal(result.records.length, 1);
  assert.equal(result.records[0].anchor.locatable, false);
  assert.ok(result.issues.some(issue => issue.code === 'anchor-target-missing' && issue.commentId === '7'));
});

test('不支持的关系外批注部件保留原文并明确列出', () => {
  const document = w(`<w:p><w:r><w:t>参数</w:t></w:r></w:p>`);
  const bytes = packageOf({
    document,
    comments: commentsXml(comment(0, '请复查')),
    relationships: rels([[COMMENT_REL, 'comments.xml']]),
    extra: { 'word/people.xml': `<w15:people xmlns:w15="${W15_NS}"/>` },
  });
  const result = extractReview(bytes);
  assert.equal(result.state, 'partial');
  assert.deepEqual(result.coverage.unsupportedParts, ['word/people.xml']);
  assert.ok(result.issues.some(issue => issue.code === 'part-unsupported' && issue.part === 'word/people.xml'));
});

test('批注部件缺失或关系目标越界时拒绝结构化收件', () => {
  const document = w(`<w:p><w:r><w:t>参数</w:t></w:r></w:p>`);
  const missing = extractReview(packageOf({ document, relationships: rels([[COMMENT_REL, 'comments.xml']]) }));
  assert.equal(missing.rejected, true);
  assert.equal(missing.issues.at(-1).code, 'comments-part-missing');
  const outside = extractReview(packageOf({
    document,
    relationships: rels([[COMMENT_REL, '../outside.xml']]),
    extra: { 'outside.xml': commentsXml(comment(0, '外部')) },
  }));
  assert.equal(outside.rejected, true);
});

test('非压缩包输入不抛异常而是拒绝', () => {
  const result = extractReview(Buffer.from('not a zip'));
  assert.equal(result.rejected, true);
  assert.equal(result.issues[0].code, 'package-unreadable');
});

test('文字视图取值受限', async () => {
  const bytes = await fixture('review');
  assert.throws(() => extractReview(bytes, { view: 'final' }), /文字视图无效/);
  const response = { accepted: '新参数\n确认', original: '旧参数\n确认', marked: '旧参数新参数\n确认' };
  for (const [view, text] of Object.entries(response)) {
    assert.equal(extractReview(bytes, { view }).records[0].anchor.text, text);
  }
});

test('重复批注 id 的每条记录都不定位，序号保证记录身份唯一', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const comments = commentsXml(comment(0, '甲'), comment(0, '乙'));
  const result = extractReview(base(document, comments));
  assert.equal(result.state, 'partial');
  assert.deepEqual(result.records.map(record => record.id), ['comment-0', 'comment-0-2']);
  assert.deepEqual(result.records.map(record => record.text), ['甲', '乙']);
  assert.ok(result.records.every(record => record.anchor.locatable === false));
  assert.deepEqual(result.issues.filter(issue => issue.code === 'comment-id-duplicated').map(issue => issue.count), [2]);
});

test('重复批注引用只作提示，起止唯一时仍定位', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="0"/>` +
    `<w:r><w:commentReference w:id="0"/></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const result = extractReview(base(document));
  assert.equal(result.state, 'partial');
  assert.equal(result.records[0].anchor.locatable, true);
  assert.equal(result.records[0].anchor.text, '参数');
  assert.equal(result.records[0].anchor.markers.references.length, 2);
  assert.ok(result.issues.some(issue => issue.code === 'anchor-reference-duplicated'));
});

test('标记 id 与批注 id 大小写和空白不同仍按同一身份处理', () => {
  const document = w(`<w:p><w:commentRangeStart w:id=" 7 "/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="7"/>` +
    `<w:r><w:commentReference w:id="7"/></w:r></w:p>`);
  const result = extractReview(base(document, commentsXml(comment(7, '意见'))));
  assert.equal(result.state, 'complete');
  assert.equal(result.records[0].anchor.locatable, true);
  assert.equal(result.records[0].anchor.markers.starts[0].nodePath, 'document/body[0]/p[0]/commentRangeStart[0]');
});

test('批量标记指向同一缺失批注时只记一条问题', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="9"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="9"/>` +
    `<w:r><w:commentReference w:id="9"/></w:r></w:p>`);
  const result = extractReview(base(document));
  const dangling = result.issues.filter(issue => issue.code === 'anchor-target-missing');
  assert.equal(dangling.length, 1);
  assert.equal(dangling[0].commentId, '9');
});

test('缺少关系部件时按约定路径读取，并保持 partial', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const bytes = packageOf({ document, comments: commentsXml(comment(0, '意见')) });
  const result = extractReview(bytes);
  assert.equal(result.records[0].anchor.text, '参数');
  assert.equal(result.state, 'partial');
  assert.ok(result.issues.some(issue => issue.code === 'relationships-missing'));
});

test('带 XML 声明与 AlternateContent 的部件只取一份可见文字', () => {
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><w:p>` +
    `<w:commentRangeStart w:id="0"/><mc:AlternateContent><mc:Choice Requires="x"><w:r><w:t>一次</w:t></w:r></mc:Choice>` +
    `<mc:Fallback><w:r><w:t>一次</w:t></w:r></mc:Fallback></mc:AlternateContent><w:commentRangeEnd w:id="0"/>` +
    `<w:r><w:commentReference w:id="0"/></w:r></w:p></w:body></w:document>`;
  const bytes = packageOf({ document, comments: commentsXml(comment(0, '意见')) });
  const result = extractReview(bytes);
  assert.equal(result.records[0].anchor.text, '一次');
});

test('批注文字中的修订不混入正文，回复段落 paraId 取末段', () => {
  const document = w(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>参数</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`);
  const comments = commentsXml(
    `<w:comment w:id="0" w:author="甲"><w:p xmlns:w14="${W14_NS}" w14:paraId="00000001"><w:r><w:t>保留</w:t></w:r>` +
    `<w:del><w:r><w:delText>删除</w:delText></w:r></w:del></w:p>` +
    `<w:p xmlns:w14="${W14_NS}" w14:paraId="00000002"><w:r><w:t>第二段</w:t></w:r></w:p></w:comment>`,
    `<w:comment w:id="1" w:author="乙"><w:p xmlns:w14="${W14_NS}" w14:paraId="00000003"><w:r><w:t>回复</w:t></w:r></w:p></w:comment>`,
  );
  const extended = `<w15:commentsEx xmlns:w15="${W15_NS}"><w15:commentEx w15:paraId="00000002"/>` +
    `<w15:commentEx w15:paraId="00000003" w15:paraIdParent="00000002"/></w15:commentsEx>`;
  const bytes = packageOf({ document, comments, extended, relationships: rels([[COMMENT_REL, 'comments.xml'], [EXTENDED_REL, 'commentsExtended.xml']]) });
  const result = extractReview(bytes);
  assert.equal(result.records[0].text, '保留\n第二段');
  assert.equal(result.records[1].replyTo, '0');
  assert.equal(result.records[1].anchor.inherited, true);
});
