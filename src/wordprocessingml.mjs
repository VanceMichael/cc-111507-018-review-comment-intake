import { unzipSync } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { posix } from 'node:path';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const COMMENTS_EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';

const MAX_PART_BYTES = 4 * 1024 * 1024;
const MAX_TOTAL_BYTES = 16 * 1024 * 1024;
const MAX_ENTRY_NAME = 256;
const SUPPORTED_PARTS = new Set([
  '[Content_Types].xml',
  '_rels/.rels',
  'word/document.xml',
  'word/comments.xml',
  'word/commentsExtended.xml',
  'word/_rels/document.xml.rels',
]);
const COMMENT_RELATED_PART = /(^|\/)(comments?|people|persons?|threads?)[^/]*\.xml$/i;

export const supportedViews = ['accepted', 'original', 'marked'];

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  processEntities: true,
});

function parseXml(bytes) {
  const text = Buffer.from(bytes).toString('utf8');
  if (/<!DOCTYPE/i.test(text)) throw new Error('XML 包含文档类型声明，不在支持范围内');
  return parser.parse(text);
}

function splitName(name) {
  const index = name.indexOf(':');
  return index === -1 ? { prefix: null, local: name } : { prefix: name.slice(0, index), local: name.slice(index + 1) };
}

function elementName(node) {
  const name = Object.keys(node).find(key => key !== ':@' && key !== '#text') ?? null;
  if (!name || name.startsWith('?') || name.startsWith('!') || name.startsWith('#')) return null;
  return name;
}

function childNamespaces(attributes, inherited) {
  const namespaces = { ...inherited };
  for (const [key, value] of Object.entries(attributes)) {
    if (typeof value !== 'string') continue;
    if (key === '@_xmlns') namespaces[''] = value;
    else if (key.startsWith('@_xmlns:')) namespaces[key.slice(8)] = value;
  }
  return namespaces;
}

function namespaceOf(name, attributes, inherited) {
  const { prefix } = splitName(name);
  const declared = attributes[`@_xmlns${prefix ? `:${prefix}` : ''}`];
  if (typeof declared === 'string') return declared;
  return inherited[prefix ?? ''] ?? null;
}

function attribute(node, local, namespace, inherited = {}) {
  const attributes = node?.[':@'] ?? {};
  for (const [key, value] of Object.entries(attributes)) {
    if (!key.startsWith('@_') || key.startsWith('@_xmlns')) continue;
    const name = key.slice(2);
    const { prefix, local: keyLocal } = splitName(name);
    if (keyLocal !== local) continue;
    if (namespace === null) {
      if (prefix === null) return value;
      continue;
    }
    if (prefix === null) continue;
    const declared = attributes[`@_xmlns:${prefix}`];
    if ((typeof declared === 'string' ? declared : inherited[prefix]) === namespace) return value;
  }
  return undefined;
}

function rootElement(document) {
  const node = document.find(item => item && typeof item === 'object' && item['#text'] === undefined && elementName(item));
  if (!node) return null;
  const name = elementName(node);
  return {
    name,
    local: splitName(name).local,
    children: node[name] ?? [],
    attributes: node[':@'] ?? {},
    namespaces: childNamespaces(node[':@'] ?? {}, {}),
  };
}

function elements(children) {
  if (!Array.isArray(children)) return [];
  return children.filter(child => child && typeof child === 'object' && child['#text'] === undefined && elementName(child));
}

function directText(children) {
  if (!Array.isArray(children)) return '';
  return children.filter(child => child && typeof child['#text'] === 'string').map(child => child['#text']).join('');
}

function canonicalId(value) {
  return String(value).trim();
}

function normalizedId(value) {
  return canonicalId(value).toLowerCase();
}

function markerRecord(state, kind, id, context, nodePath) {
  return {
    commentId: canonicalId(id),
    kind,
    paragraph: context.paragraph,
    offset: state.offset,
    nodePath,
  };
}

function storeMarker(state, map, record) {
  const key = normalizedId(record.commentId);
  const list = map.get(key) ?? [];
  list.push(record);
  map.set(key, list);
}

function newContext() {
  return { paragraph: 0, del: 0, ins: 0 };
}

function walkDocument(children, state, context, namespaces, path) {
  const siblingCounts = new Map();
  for (const node of children) {
    const name = elementName(node);
    if (!name) continue;
    const { local } = splitName(name);
    const index = siblingCounts.get(local) ?? 0;
    siblingCounts.set(local, index + 1);
    const attributes = node[':@'] ?? {};
    const here = childNamespaces(attributes, namespaces);
    const nodePath = [...path, `${local}[${index}]`].join('/');
    const body = node[name] ?? [];
    const namespace = namespaceOf(name, attributes, namespaces);
    const inWord = namespace === W;

    if (namespace === MC && local === 'AlternateContent') {
      const choice = elements(body).find(item => splitName(elementName(item)).local === 'Choice');
      if (choice) walkDocument(elements(choice[elementName(choice)] ?? []), state, context, childNamespaces(choice[':@'] ?? {}, here), [...path, `${local}[${index}]`]);
      continue;
    }

    if (inWord && local === 'p') {
      if (state.paragraphStarted) {
        state.pieces.push({ chars: ['\n'], kind: 'break', paragraph: context.paragraph, nodePath: path.join('/') });
        state.offset += 1;
        state.paragraphBreaks += 1;
      }
      context.paragraph = state.paragraphCount;
      state.paragraphCount += 1;
      state.paragraphStarted = true;
      walkDocument(elements(body), state, context, here, [...path, `${local}[${index}]`]);
      continue;
    }

    if (inWord && (local === 't' || local === 'delText')) {
      const text = directText(body);
      const kind = local === 'delText' || context.del > 0 ? 'deleted' : context.ins > 0 ? 'inserted' : 'text';
      if (text) {
        const chars = Array.from(text);
        state.pieces.push({ chars, kind, paragraph: context.paragraph, nodePath });
        state.offset += chars.length;
      }
      continue;
    }

    if (inWord && local === 'tab') {
      state.pieces.push({ chars: ['\t'], kind: context.del > 0 ? 'deleted' : context.ins > 0 ? 'inserted' : 'text', paragraph: context.paragraph, nodePath });
      state.offset += 1;
      continue;
    }

    if (inWord && (local === 'br' || local === 'cr')) {
      state.pieces.push({ chars: ['\n'], kind: context.del > 0 ? 'deleted' : context.ins > 0 ? 'inserted' : 'text', paragraph: context.paragraph, nodePath });
      state.offset += 1;
      continue;
    }

    if (inWord && local === 'del') {
      context.del += 1;
      walkDocument(elements(body), state, context, here, [...path, `${local}[${index}]`]);
      context.del -= 1;
      continue;
    }

    if (inWord && local === 'ins') {
      context.ins += 1;
      walkDocument(elements(body), state, context, here, [...path, `${local}[${index}]`]);
      context.ins -= 1;
      continue;
    }

    if (inWord && local === 'commentRangeStart') {
      const id = attribute(node, 'id', W, namespaces);
      if (id === undefined) state.malformedMarkers.push({ kind: '范围起点', nodePath });
      else storeMarker(state, state.starts, markerRecord(state, 'start', id, context, nodePath));
      continue;
    }

    if (inWord && local === 'commentRangeEnd') {
      const id = attribute(node, 'id', W, namespaces);
      if (id === undefined) state.malformedMarkers.push({ kind: '范围终点', nodePath });
      else storeMarker(state, state.ends, markerRecord(state, 'end', id, context, nodePath));
      continue;
    }

    if (inWord && local === 'commentReference') {
      const id = attribute(node, 'id', W, namespaces);
      if (id === undefined) state.malformedMarkers.push({ kind: '批注引用', nodePath });
      else storeMarker(state, state.refs, markerRecord(state, 'reference', id, context, nodePath));
      continue;
    }

    walkDocument(elements(body), state, context, here, [...path, `${local}[${index}]`]);
  }
}

function collectCommentText(children, namespaces) {
  const paragraphs = [];
  let current = null;

  const walk = (nodes, ns, deleted) => {
    for (const node of elements(nodes)) {
      const name = elementName(node);
      const attributes = node[':@'] ?? {};
      const here = childNamespaces(attributes, ns);
      const { local } = splitName(name);
      const inWord = namespaceOf(name, attributes, ns) === W;
      const body = node[name] ?? [];
      if (inWord && local === 'p') {
        current = { text: '', paraId: attribute(node, 'paraId', W14, here) ?? null };
        paragraphs.push(current);
        walk(body, here, deleted);
        current = null;
        continue;
      }
      if (inWord && local === 'del') { walk(body, here, true); continue; }
      if (inWord && local === 'ins') { walk(body, here, deleted); continue; }
      if (inWord && (local === 't' || local === 'delText')) {
        if (current && !deleted) current.text += directText(body);
        continue;
      }
      if (inWord && local === 'tab') {
        if (current && !deleted) current.text += '\t';
        continue;
      }
      if (inWord && (local === 'br' || local === 'cr')) {
        if (current && !deleted) current.text += '\n';
        continue;
      }
      walk(body, here, deleted);
    }
  };

  walk(children, namespaces, false);
  return paragraphs;
}

function parseComments(document) {
  const root = rootElement(document);
  if (!root || namespaceOf(root.name, root.attributes, root.namespaces) !== W) return { error: '批注部件命名空间无效' };
  const comments = [];
  const issues = [];
  const seen = new Set();
  for (const node of elements(root.children)) {
    const name = elementName(node);
    const attributes = node[':@'] ?? {};
    const here = childNamespaces(attributes, root.namespaces);
    if (splitName(name).local !== 'comment' || namespaceOf(name, attributes, root.namespaces) !== W) continue;
    const id = attribute(node, 'id', W, root.namespaces);
    if (id === undefined) {
      issues.push({ code: 'comment-id-missing', message: '批注缺少 id，无法建立稳定身份' });
      continue;
    }
    const key = normalizedId(id);
    const duplicate = seen.has(key);
    seen.add(key);
    const paragraphs = collectCommentText(node[name] ?? [], here);
    comments.push({
      commentId: canonicalId(id),
      duplicate,
      author: attribute(node, 'author', W, root.namespaces) ?? null,
      date: attribute(node, 'date', W, root.namespaces) ?? null,
      initials: attribute(node, 'initials', W, root.namespaces) ?? null,
      text: paragraphs.map(paragraph => paragraph.text).join('\n'),
      paraIds: paragraphs.map(paragraph => paragraph.paraId).filter(Boolean),
      paraId: paragraphs.map(paragraph => paragraph.paraId).filter(Boolean).at(-1) ?? null,
    });
  }
  const counts = new Map();
  for (const comment of comments) counts.set(normalizedId(comment.commentId), (counts.get(normalizedId(comment.commentId)) ?? 0) + 1);
  for (const comment of comments) {
    if (counts.get(normalizedId(comment.commentId)) > 1 && !comment.duplicate) comment.duplicate = true;
  }
  for (const [key, count] of counts) {
    if (count > 1) {
      const sample = comments.find(comment => normalizedId(comment.commentId) === key);
      issues.push({ code: 'comment-id-duplicated', commentId: sample.commentId, count, message: `批注 id ${sample.commentId} 重复，身份不唯一` });
    }
  }
  return { comments, issues };
}

function parseCommentsExtended(document) {
  const root = rootElement(document);
  if (!root) return { error: '扩展批注部件为空' };
  if (namespaceOf(root.name, root.attributes, root.namespaces) !== W15) return { error: '扩展批注部件命名空间无效' };
  const entries = [];
  const issues = [];
  const seen = new Set();
  for (const node of elements(root.children)) {
    const name = elementName(node);
    const attributes = node[':@'] ?? {};
    if (splitName(name).local !== 'commentEx' || namespaceOf(name, attributes, root.namespaces) !== W15) continue;
    const paraId = attribute(node, 'paraId', W15, root.namespaces);
    if (paraId === undefined) {
      issues.push({ code: 'reply-entry-incomplete', message: '扩展批注条目缺少 paraId' });
      continue;
    }
    const key = normalizedId(paraId);
    if (seen.has(key)) {
      issues.push({ code: 'reply-entry-duplicated', message: `扩展批注段落 ${paraId} 重复` });
      continue;
    }
    seen.add(key);
    const paraIdParent = attribute(node, 'paraIdParent', W15, root.namespaces);
    entries.push({ paraId: String(paraId), paraIdParent: paraIdParent === undefined ? null : String(paraIdParent) });
  }
  return { entries, issues };
}

function parseRelationships(document) {
  const root = rootElement(document);
  if (!root) return { error: '关系部件为空' };
  const relationships = [];
  for (const node of elements(root.children)) {
    const name = elementName(node);
    const attributes = node[':@'] ?? {};
    if (splitName(name).local !== 'Relationship') continue;
    const type = attributes['@_Type'];
    const target = attributes['@_Target'];
    if (typeof type === 'string' && typeof target === 'string') relationships.push({ type, target });
  }
  return { relationships };
}

function resolveTarget(base, target) {
  let decoded = target;
  try { decoded = decodeURIComponent(target); } catch { /* 保留原值，交由缺失部件处理 */ }
  const normalized = decoded.replace(/\\/g, '/');
  if (normalized.startsWith('/')) return posix.normalize(normalized).slice(1);
  return posix.normalize(posix.join(base, normalized)).replace(/^\.\//, '');
}

function viewProjection(pieces, view) {
  const visible = piece => view === 'marked' ? true : view === 'accepted' ? piece.kind !== 'deleted' : piece.kind !== 'inserted';
  const spans = [];
  const chars = [];
  let streamOffset = 0;
  let viewOffset = 0;
  for (const piece of pieces) {
    const length = piece.chars.length;
    const shown = visible(piece);
    spans.push({
      kind: piece.kind,
      streamStart: streamOffset,
      streamEnd: streamOffset + length,
      viewStart: viewOffset,
      viewEnd: viewOffset + (shown ? length : 0),
      visible: shown,
    });
    if (shown) {
      for (const character of piece.chars) chars.push(character);
      viewOffset += length;
    }
    streamOffset += length;
  }
  return { spans, chars, length: viewOffset, total: streamOffset };
}

function viewOffsetAt(projection, offset) {
  const { spans } = projection;
  let low = 0;
  let high = spans.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const span = spans[middle];
    if (offset < span.streamStart) high = middle - 1;
    else if (offset > span.streamEnd) low = middle + 1;
    else {
      const inside = Math.min(offset, span.streamEnd) - span.streamStart;
      return span.viewStart + (span.visible ? inside : 0);
    }
  }
  return offset >= projection.total ? projection.length : 0;
}

function anchorView(projection, streamStart, streamEnd) {
  const start = viewOffsetAt(projection, streamStart);
  const end = viewOffsetAt(projection, streamEnd);
  const text = projection.chars.slice(start, end).join('');
  const inside = projection.spans.filter(span => span.streamEnd > streamStart && span.streamStart < streamEnd);
  const visibleKinds = new Set(inside.filter(span => span.visible).map(span => span.kind));
  const omittedKinds = new Set(inside.filter(span => !span.visible).map(span => span.kind));
  return {
    start,
    end,
    text,
    empty: text.length === 0,
    includes: {
      text: visibleKinds.has('text'),
      inserted: visibleKinds.has('inserted'),
      deleted: visibleKinds.has('deleted'),
      breaks: visibleKinds.has('break'),
    },
    omits: {
      inserted: omittedKinds.has('inserted'),
      deleted: omittedKinds.has('deleted'),
    },
  };
}

function markerSummary(record) {
  if (!record) return null;
  return { nodePath: record.nodePath, offset: record.offset, paragraph: record.paragraph };
}

function cycleMembers(parents) {
  const finished = new Set();
  const members = new Set();
  for (const start of parents.keys()) {
    if (finished.has(start)) continue;
    const stack = [];
    const position = new Map();
    let current = start;
    while (current !== undefined && !finished.has(current)) {
      if (position.has(current)) {
        for (let index = position.get(current); index < stack.length; index += 1) members.add(stack[index]);
        break;
      }
      position.set(current, stack.length);
      stack.push(current);
      current = parents.get(current);
    }
    for (const id of stack) finished.add(id);
  }
  return members;
}

function threadRoot(commentId, parents) {
  const seen = new Set();
  let current = commentId;
  while (parents.has(current) && !seen.has(current)) {
    seen.add(current);
    current = parents.get(current);
  }
  return current;
}

function unsupportedParts(names) {
  return names.filter(name => !name.endsWith('/') && name.endsWith('.xml') && !SUPPORTED_PARTS.has(name)).sort();
}

function unlocatedAnchor(markers, reason) {
  return { locatable: false, reason: reason ?? null, text: null, range: null, views: null, markers };
}

function rejected(reason, code, issues = [], view = null) {
  return {
    rejected: true,
    reason,
    records: [],
    issues: [...issues, { code, message: reason }],
    coverage: { structured: true, supported: false, view, reason },
  };
}

export function extractReview(bytes, { view = 'accepted' } = {}) {
  if (!supportedViews.includes(view)) throw new Error('文字视图无效，可选 accepted、original 或 marked');
  const issues = [];
  const skipped = [];
  let archive;
  let budget = MAX_TOTAL_BYTES;
  try {
    archive = unzipSync(bytes, { filter: file => {
      const tooLarge = file.originalSize > MAX_PART_BYTES;
      const tooLong = file.name.length > MAX_ENTRY_NAME;
      const overBudget = file.originalSize > budget;
      if (tooLarge || tooLong || overBudget) {
        skipped.push({ name: file.name.slice(0, MAX_ENTRY_NAME), reason: tooLarge ? 'part-too-large' : tooLong ? 'name-too-long' : 'total-too-large' });
        return false;
      }
      budget -= file.originalSize;
      return true;
    } });
  } catch {
    return rejected('DOCX 压缩包无法读取', 'package-unreadable', [], view);
  }

  if (!archive['word/document.xml']) return rejected('DOCX 缺少正文部件', 'document-missing', [], view);
  let document;
  try { document = parseXml(archive['word/document.xml']); }
  catch { return rejected('正文部件无法解析', 'document-unreadable', [], view); }
  const documentRoot = rootElement(document);
  if (!documentRoot || namespaceOf(documentRoot.name, documentRoot.attributes, documentRoot.namespaces) !== W) {
    return rejected('正文部件不是 WordprocessingML 主命名空间', 'document-namespace-invalid', [], view);
  }

  for (const entry of skipped) {
    const reason = entry.reason === 'part-too-large' ? '超出单部件限额'
      : entry.reason === 'name-too-long' ? '名称超出长度限额' : '超出解压总量限额';
    issues.push({ code: 'part-skipped', part: entry.name, detail: entry.reason, message: `部件 ${entry.name} ${reason}，已跳过` });
  }
  const names = Object.keys(archive);
  const unsupported = unsupportedParts(names);
  for (const name of unsupported.filter(name => COMMENT_RELATED_PART.test(name))) {
    issues.push({ code: 'part-unsupported', part: name, message: `未支持的批注相关部件 ${name}，已保留原文但未解析` });
  }

  let relationships = null;
  if (archive['word/_rels/document.xml.rels']) {
    try {
      const parsed = parseRelationships(parseXml(archive['word/_rels/document.xml.rels']));
      if (parsed.error) issues.push({ code: 'relationships-unreadable', message: parsed.error });
      else relationships = parsed.relationships;
    } catch { issues.push({ code: 'relationships-unreadable', message: '关系部件无法解析' }); }
  } else {
    issues.push({ code: 'relationships-missing', message: '缺少文档关系部件，批注部件按约定路径读取' });
  }

  let commentsTarget = 'word/comments.xml';
  let extendedTarget = null;
  if (relationships) {
    const commentsRelationships = relationships.filter(item => item.type === `${R}/comments`);
    if (commentsRelationships.length > 1) {
      issues.push({ code: 'comments-relationship-duplicated', message: '批注关系重复，无法确认唯一批注部件' });
    }
    if (commentsRelationships.length) commentsTarget = resolveTarget('word', commentsRelationships[0].target);
    else issues.push({ code: 'comments-relationship-missing', message: '缺少批注关系，批注部件按约定路径读取' });
    const extendedRelationship = relationships.find(item => item.type === COMMENTS_EXTENDED_REL);
    if (extendedRelationship) extendedTarget = resolveTarget('word', extendedRelationship.target);
  }
  if (commentsTarget !== 'word/comments.xml') {
    return rejected(`批注关系指向 ${commentsTarget}，不在承诺范围内`, 'comments-part-unexpected-target', issues, view);
  }
  if (!archive[commentsTarget]) return rejected('批注部件缺失或关系目标不存在', 'comments-part-missing', issues, view);

  let parsedComments;
  try { parsedComments = parseComments(parseXml(archive[commentsTarget])); }
  catch { return rejected('批注部件无法解析', 'comments-unreadable', issues, view); }
  if (parsedComments.error) return rejected(parsedComments.error, 'comments-unreadable', issues, view);
  issues.push(...parsedComments.issues);

  const state = {
    pieces: [], offset: 0, paragraphCount: 0, paragraphStarted: false, paragraphBreaks: 0,
    starts: new Map(), ends: new Map(), refs: new Map(), malformedMarkers: [],
  };
  walkDocument(elements(documentRoot.children), state, newContext(), documentRoot.namespaces, [documentRoot.local]);
  for (const marker of state.malformedMarkers) {
    issues.push({ code: 'marker-id-missing', part: 'document', nodePath: marker.nodePath,
      message: `文档中的批注${marker.kind}标记缺少 id，无法绑定批注` });
  }

  let extended = null;
  if (extendedTarget && archive[extendedTarget]) {
    try {
      const parsed = parseCommentsExtended(parseXml(archive[extendedTarget]));
      if (parsed.error) issues.push({ code: 'reply-relation-unreadable', message: parsed.error });
      else {
        extended = parsed.entries;
        issues.push(...parsed.issues);
      }
    } catch { issues.push({ code: 'reply-relation-unreadable', message: '扩展批注部件无法解析' }); }
  } else if (extendedTarget) {
    issues.push({ code: 'reply-relation-incomplete', part: extendedTarget, message: '扩展批注关系指向的部件不存在' });
  }

  const projections = {
    accepted: viewProjection(state.pieces, 'accepted'),
    original: viewProjection(state.pieces, 'original'),
    marked: viewProjection(state.pieces, 'marked'),
  };

  const parents = new Map();
  const threadInfo = new Map();
  if (extended) {
    const byParaId = new Map();
    for (const comment of parsedComments.comments) {
      for (const paraId of comment.paraIds) {
        const key = normalizedId(paraId);
        if (byParaId.has(key) && byParaId.get(key) !== comment.commentId) {
          issues.push({ code: 'reply-paragraph-duplicated', commentId: comment.commentId,
            message: `批注段落 ${paraId} 对应多条批注，回复关系无法唯一确定` });
          continue;
        }
        byParaId.set(key, comment.commentId);
      }
    }
    for (const entry of extended) {
      if (entry.paraIdParent === null) continue;
      const childId = byParaId.get(normalizedId(entry.paraId));
      const parentId = byParaId.get(normalizedId(entry.paraIdParent));
      if (childId === undefined) {
        issues.push({ code: 'reply-source-missing', message: `回复关系中的段落 ${entry.paraId} 对应不到批注` });
        continue;
      }
      if (parentId === undefined) {
        issues.push({ code: 'reply-target-missing', commentId: childId, message: `批注 ${childId} 的上级批注不存在` });
        continue;
      }
      if (parentId === childId) {
        issues.push({ code: 'reply-cycle', commentId: childId, message: `批注 ${childId} 的回复关系指向自身` });
        continue;
      }
      if (parents.has(childId)) {
        issues.push({ code: 'reply-target-duplicated', commentId: childId, message: `批注 ${childId} 存在多个上级批注` });
        continue;
      }
      parents.set(childId, parentId);
    }
    for (const commentId of cycleMembers(parents)) {
      issues.push({ code: 'reply-cycle', commentId, message: `批注 ${commentId} 的回复关系形成循环` });
      parents.delete(commentId);
    }
    for (const [commentId, parentId] of parents) {
      threadInfo.set(commentId, { replyTo: parentId, threadRoot: threadRoot(commentId, parents) });
    }
  }

  const records = [];
  const byCommentId = new Map();
  const order = new Map();
  parsedComments.comments.forEach((comment, index) => {
    if (!comment.duplicate) order.set(normalizedId(comment.commentId), index);
  });
  const anchorOf = commentId => resolveAnchor(commentId, new Set());

  function resolveAnchor(commentId, visiting) {
    const cacheKey = normalizedId(commentId);
    if (byCommentId.has(cacheKey)) return byCommentId.get(cacheKey);
    const comment = parsedComments.comments[order.get(normalizedId(commentId))];
    if (!comment || visiting.has(commentId)) return null;
    visiting.add(commentId);
    const key = normalizedId(commentId);
    const starts = state.starts.get(key) ?? [];
    const ends = state.ends.get(key) ?? [];
    const references = state.refs.get(key) ?? [];
    const thread = threadInfo.get(commentId) ?? { replyTo: null, threadRoot: commentId };
    const counts = { start: starts.length, end: ends.length, reference: references.length };
    const missing = Object.keys(counts).filter(name => counts[name] === 0);
    const duplicated = Object.keys(counts).filter(name => counts[name] > 1 && name !== 'reference');
    const markers = { starts: starts.map(markerSummary), ends: ends.map(markerSummary), references: references.map(markerSummary) };
    let anchor;

    if (starts.length === 0 && ends.length === 0) {
      anchor = unlocatedAnchor(markers, null);
      if (thread.replyTo) {
        const parentAnchor = resolveAnchor(thread.replyTo, visiting);
        if (parentAnchor?.locatable) {
          anchor = {
            locatable: true,
            inherited: true,
            parentId: thread.replyTo,
            reason: null,
            text: parentAnchor.views[view].text,
            markers: parentAnchor.markers,
            ownMarkers: markers,
            range: parentAnchor.range,
            paragraphs: parentAnchor.paragraphs,
            views: parentAnchor.views,
          };
        } else {
          const message = `批注 ${commentId} 依附的上级批注 ${thread.replyTo} 无法定位`;
          issues.push({ code: 'reply-anchor-incomplete', commentId, message });
          anchor = unlocatedAnchor(markers, message);
        }
      } else {
        const message = `批注 ${commentId} 在正文中没有批注锚点`;
        issues.push({ code: 'comment-without-anchor', commentId, message });
        anchor = unlocatedAnchor(markers, message);
      }
    } else if (duplicated.length) {
      const message = `批注 ${commentId} 的${duplicated.map(labelOf).join('、')}标记重复，无法确定文字范围`;
      issues.push({ code: 'anchor-marker-duplicated', commentId, duplicated, message });
      anchor = unlocatedAnchor(markers, message);
    } else if (missing.includes('start') || missing.includes('end')) {
      const message = `批注 ${commentId} 缺少${missing.filter(name => name !== 'reference').map(labelOf).join('、')}，无法确定文字范围`;
      issues.push({ code: 'anchor-range-incomplete', commentId, missing, message });
      anchor = unlocatedAnchor(markers, message);
    } else if (starts[0].offset > ends[0].offset) {
      const message = `批注 ${commentId} 的范围终点位于起点之前，无法确定文字范围`;
      issues.push({ code: 'anchor-order-invalid', commentId, message });
      anchor = unlocatedAnchor(markers, message);
    } else {
      if (missing.includes('reference')) {
        issues.push({ code: 'anchor-reference-missing', commentId,
          message: `批注 ${commentId} 缺少批注引用，范围仍按起止标记定位` });
      }
      if (counts.reference > 1) {
        issues.push({ code: 'anchor-reference-duplicated', commentId,
          message: `批注 ${commentId} 的批注引用出现 ${counts.reference} 次，范围仍按起止标记定位` });
      }
      const streamStart = starts[0].offset;
      const streamEnd = ends[0].offset;
      anchor = {
        locatable: true,
        markers,
        range: { start: streamStart, end: streamEnd, length: streamEnd - streamStart },
        paragraphs: { start: starts[0].paragraph, end: ends[0].paragraph },
        views: {
          accepted: anchorView(projections.accepted, streamStart, streamEnd),
          original: anchorView(projections.original, streamStart, streamEnd),
          marked: anchorView(projections.marked, streamStart, streamEnd),
        },
      };
      anchor.text = anchor.views[view].text;
    }
    byCommentId.set(cacheKey, anchor);
    return anchor;
  }

  const idCounts = new Map();
  for (const comment of parsedComments.comments) {
    const thread = threadInfo.get(comment.commentId) ?? { replyTo: null, threadRoot: comment.commentId };
    const anchor = comment.duplicate
      ? unlocatedAnchor({
        starts: (state.starts.get(normalizedId(comment.commentId)) ?? []).map(markerSummary),
        ends: (state.ends.get(normalizedId(comment.commentId)) ?? []).map(markerSummary),
        references: (state.refs.get(normalizedId(comment.commentId)) ?? []).map(markerSummary),
      }, `批注 id ${comment.commentId} 重复，无法唯一绑定锚点`)
      : anchorOf(comment.commentId);
    const occurrence = (idCounts.get(comment.commentId) ?? 0) + 1;
    idCounts.set(comment.commentId, occurrence);
    records.push({
      id: occurrence === 1 ? `comment-${comment.commentId}` : `comment-${comment.commentId}-${occurrence}`,
      commentId: comment.commentId,
      author: comment.author,
      date: comment.date,
      initials: comment.initials,
      text: comment.text,
      replyTo: thread.replyTo,
      threadRoot: thread.threadRoot,
      anchor,
    });
  }

  const commentKeys = new Set(parsedComments.comments.map(comment => normalizedId(comment.commentId)));
  const dangling = new Map();
  const markerSets = [[state.starts, '范围起点'], [state.ends, '范围终点'], [state.refs, '批注引用']];
  for (const [map, label] of markerSets) {
    for (const [key, markers] of map) {
      if (commentKeys.has(key)) continue;
      const entry = dangling.get(key) ?? { commentId: markers[0].commentId, kinds: [] };
      entry.kinds.push(label);
      dangling.set(key, entry);
    }
  }
  for (const { commentId, kinds } of dangling.values()) {
    issues.push({ code: 'anchor-target-missing', commentId,
      message: `正文中的批注${kinds.join('、')} ${commentId} 在批注部件中不存在` });
  }

  const anchored = records.filter(record => record.anchor?.locatable).length;
  const coverage = {
    structured: true,
    supported: true,
    view,
    parts: {
      document: 'word/document.xml',
      comments: commentsTarget,
      commentsExtended: extendedTarget && archive[extendedTarget] ? extendedTarget : null,
      relationships: Boolean(relationships),
    },
    unsupportedParts: unsupported,
    views: Object.fromEntries(Object.entries(projections).map(([name, projection]) => [name, { codePoints: projection.length }])),
    text: {
      paragraphs: state.paragraphCount,
      codePoints: state.offset,
      breaks: state.paragraphBreaks,
    },
    comments: {
      total: records.length,
      anchored,
      inherited: records.filter(record => record.anchor?.inherited).length,
      unlocated: records.length - anchored,
      replies: records.filter(record => record.replyTo).length,
    },
    issues: issues.length,
  };

  return {
    rejected: false,
    state: issues.length ? 'partial' : 'complete',
    records,
    issues,
    coverage,
    reason: null,
  };
}

function labelOf(name) {
  return name === 'start' ? '范围起点' : name === 'end' ? '范围终点' : '批注引用';
}
