import { receiveBytes } from './http-client.mjs';
import { inspect } from './source-metadata.mjs';
import { parseReviewPackage, buildRecords, VIEWS } from './docx-review.mjs';

// 结构化收件：领取入口 DOCX，按 docs/来源格式.md 的范围解析批注、
// 修订视图与回复关系。解析结果完整为 complete，存在可定位问题为 partial，
// 完全无法解释为 rejected；rejected 也保留原文与证据，不丢弃来源字节。
export async function collectStructured({ input, descriptor, sourceBase, store }) {
  const view = input.view ?? 'marked';
  if (!VIEWS.includes(view)) throw new Error('文字视图无效');
  const source = await receiveBytes(`${sourceBase}/files/${encodeURIComponent(descriptor.entry)}`);
  const blob = await store.blob(source.bytes);
  const base = {
    mode: 'structured',
    product: descriptor.product,
    source: descriptor.entry,
    sourceVersion: source.etag,
    metadata: inspect(source.bytes),
    blobs: [blob]
  };
  const evidence = { entry: descriptor.entry, etag: source.etag, sha256: blob.digest, bytes: source.bytes.length };
  const model = parseReviewPackage(source.bytes);
  if (model.status !== 'parsed') {
    const receipt = await store.save({
      ...base,
      state: 'rejected',
      records: [],
      issues: model.issues,
      coverage: { structured: true, view, views: VIEWS, comments: 0, anchored: 0, replies: 0, paragraphs: 0, parts: [], unsupportedParts: [], complete: false },
      evidence: { ...evidence, parts: [] }
    });
    return { receipt };
  }
  const records = buildRecords(model, view);
  const state = model.issues.length ? 'partial' : 'complete';
  const receipt = await store.save({
    ...base,
    state,
    records,
    issues: model.issues,
    coverage: {
      structured: true,
      view,
      views: VIEWS,
      comments: model.comments.length,
      anchored: model.comments.filter((comment) => comment.anchor).length,
      replies: model.comments.filter((comment) => comment.parentId !== null).length,
      paragraphs: model.paragraphs,
      parts: model.parts,
      unsupportedParts: model.unsupportedParts,
      complete: model.issues.length === 0
    },
    evidence: { ...evidence, parts: model.partSources }
  });
  return { receipt };
}
