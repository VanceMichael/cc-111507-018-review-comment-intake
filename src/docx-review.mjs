// 采购评审 DOCX 审阅资料解析：批注锚点、修订视图与回复关系。
// 只读取 docs/来源格式.md 承诺的范围：word/document.xml、word/comments.xml、
// word/commentsExtended.xml 与关系文件，主命名空间按 URI 识别。
// 锚点绑定只依赖原始标记位置（commentRangeStart/End 所在节点路径），
// 不做可见文本匹配；端点缺失、重复或关系不完整时输出 issue，而不是猜测位置。

import { unzipSync, strFromU8 } from 'fflate';
import { XMLParser, XMLValidator } from 'fast-xml-parser';

export const VIEWS = ['accepted', 'original', 'marked'];

const NS_MAIN = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const NS_W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const KNOWN_PARTS = new Set([
  'word/document.xml',
  'word/comments.xml',
  'word/commentsExtended.xml',
  'word/_rels/document.xml.rels'
]);

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false
});

const codePoints = (text) => [...text].length;

function nodeTag(node) {
  for (const key of Object.keys(node)) if (key !== ':@') return key;
  return null;
}

function isElement(node) {
  const tag = nodeTag(node);
  return tag !== null && tag !== '#text' && !tag.startsWith('?');
}

function extendNsMap(nsMap, attrs) {
  if (!attrs) return nsMap;
  let extended = null;
  for (const [key, value] of Object.entries(attrs)) {
    if (key === '@_xmlns' || key.startsWith('@_xmlns:')) {
      if (!extended) extended = new Map(nsMap);
      extended.set(key === '@_xmlns' ? '' : key.slice(8), value);
    }
  }
  return extended || nsMap;
}

function resolveName(qname, nsMap) {
  const colon = qname.indexOf(':');
  const prefix = colon === -1 ? '' : qname.slice(0, colon);
  return { uri: nsMap.get(prefix) || '', local: colon === -1 ? qname : qname.slice(colon + 1) };
}

function attribute(node, nsMap, uri, local) {
  const attrs = node[':@'] || {};
  for (const [key, value] of Object.entries(attrs)) {
    if (key.startsWith('@_xmlns')) continue;
    const resolved = resolveName(key.slice(2), nsMap);
    if (resolved.uri === uri && resolved.local === local) return value;
  }
  return null;
}

function elementChildren(node, nsMap) {
  const tag = nodeTag(node);
  return (node[tag] || [])
    .filter(isElement)
    .map((child) => ({ node: child, nsMap: extendNsMap(nsMap, child[':@']) }));
}

function textOf(node) {
  const tag = nodeTag(node);
  return (node[tag] || [])
    .filter((child) => nodeTag(child) === '#text')
    .map((child) => child['#text'])
    .join('');
}

function parseXml(text) {
  if (XMLValidator.validate(text) !== true) return null;
  return parser.parse(text);
}

function rootElement(parsed) {
  const node = parsed.find(isElement);
  if (!node) return null;
  const nsMap = extendNsMap(new Map(), node[':@']);
  const tag = nodeTag(node);
  const { uri, local } = resolveName(tag, nsMap);
  return { node, tag, nsMap, uri, local };
}

// 关系文件：按 Type 后缀识别批注与回复关系部分，Target 相对 word/ 解析。
function readRelationships(parsed) {
  const found = {};
  const visit = (nodes) => {
    for (const node of nodes) {
      if (!isElement(node)) continue;
      const tag = nodeTag(node);
      const local = tag.includes(':') ? tag.slice(tag.indexOf(':') + 1) : tag;
      if (local === 'Relationship') {
        const attrs = node[':@'] || {};
        const target = attrs['@_Target'];
        const type = attrs['@_Type'] || '';
        if (target) {
          const path = target.startsWith('/') ? target.slice(1) : `word/${target}`;
          if (type.endsWith('/comments')) found.comments = { path, id: attrs['@_Id'] };
          else if (type.endsWith('/commentsExtended')) found.commentsExtended = { path, id: attrs['@_Id'] };
        }
      }
      visit(node[tag] || []);
    }
  };
  visit(parsed);
  return found;
}

// 正文遍历：把段落、文字段（含修订状态）、批注范围标记与引用
// 按文档顺序展开成事件流。未识别的元素透明下钻，不改变文字顺序。
function walkChildren(nodes, ctx, events, state) {
  const counters = new Map();
  for (const node of nodes) {
    if (!isElement(node)) continue;
    const qname = nodeTag(node);
    const count = counters.get(qname) || 0;
    counters.set(qname, count + 1);
    const path = ctx.path ? `${ctx.path}/${qname}[${count}]` : qname;
    const nsMap = extendNsMap(ctx.nsMap, node[':@']);
    const { uri, local } = resolveName(qname, nsMap);
    const children = node[qname] || [];
    if (uri === NS_MAIN) {
      if (local === 'p') {
        const index = state.paragraphs;
        state.paragraphs += 1;
        events.push({ type: 'paragraphStart', index });
        walkChildren(children, { path, nsMap, status: ctx.status, paragraph: index }, events, state);
        events.push({ type: 'paragraphEnd', index });
        continue;
      }
      if (local === 'ins' || local === 'del') {
        const status = local === 'ins' ? { ...ctx.status, ins: true } : { ...ctx.status, del: true };
        walkChildren(children, { path, nsMap, status, paragraph: ctx.paragraph }, events, state);
        continue;
      }
      if (local === 't' || local === 'delText') {
        const status = local === 'delText' || ctx.status.del ? 'deleted' : ctx.status.ins ? 'inserted' : 'normal';
        events.push({ type: 'segment', paragraph: ctx.paragraph, path, text: textOf(node), status });
        continue;
      }
      if (local === 'commentRangeStart' || local === 'commentRangeEnd') {
        events.push({
          type: 'boundary',
          kind: local === 'commentRangeStart' ? 'start' : 'end',
          id: attribute(node, nsMap, NS_MAIN, 'id'),
          marker: path,
          paragraph: ctx.paragraph
        });
        continue;
      }
      if (local === 'commentReference') {
        events.push({ type: 'reference', id: attribute(node, nsMap, NS_MAIN, 'id'), marker: path, paragraph: ctx.paragraph });
        continue;
      }
    }
    walkChildren(children, { path, nsMap, status: ctx.status, paragraph: ctx.paragraph }, events, state);
  }
}

function includedIn(status, view) {
  if (view === 'marked') return true;
  if (view === 'accepted') return status !== 'deleted';
  return status !== 'inserted';
}

// 把事件流展开成线性序列（段间插入换行段），并为每个事件记录
// 在 accepted / original / marked 三个视图中的 code point 偏移。
function indexDocument(parsed) {
  const events = [];
  const state = { paragraphs: 0 };
  walkChildren(parsed, { path: '', nsMap: new Map(), status: {}, paragraph: null }, events, state);
  const linear = [];
  let seenParagraph = false;
  for (const event of events) {
    if (event.type === 'paragraphStart') {
      if (seenParagraph) linear.push({ type: 'segment', paragraph: event.index, path: null, text: '\n', status: 'normal', separator: true });
      seenParagraph = true;
    } else if (event.type !== 'paragraphEnd') {
      linear.push(event);
    }
  }
  const offsets = { accepted: 0, original: 0, marked: 0 };
  const viewChars = { accepted: [], original: [], marked: [] };
  const segments = [];
  linear.forEach((event, index) => {
    event.order = index;
    if (event.type === 'segment') {
      event.views = {};
      for (const view of VIEWS) {
        if (includedIn(event.status, view)) {
          event.views[view] = offsets[view];
          const chars = [...event.text];
          offsets[view] += chars.length;
          viewChars[view].push(...chars);
        } else {
          event.views[view] = null;
        }
      }
      segments.push(event);
    } else {
      event.views = { ...offsets };
    }
  });
  return { linear, segments, viewChars, paragraphs: state.paragraphs };
}

// 锚点绑定到最近的可复查文字节点：优先取标记之前的文字末端，
// 否则取之后的文字起点；全篇无文字时保留 null，不用文本内容猜测位置。
function nearestText(linear, order) {
  for (let index = order - 1; index >= 0; index -= 1) {
    if (linear[index].type === 'segment' && linear[index].path) {
      return { path: linear[index].path, offset: codePoints(linear[index].text) };
    }
  }
  for (let index = order + 1; index < linear.length; index += 1) {
    if (linear[index].type === 'segment' && linear[index].path) {
      return { path: linear[index].path, offset: 0 };
    }
  }
  return null;
}

function pointOf(event, linear) {
  const binding = nearestText(linear, event.order);
  return {
    marker: event.marker,
    path: binding ? binding.path : null,
    offset: binding ? binding.offset : null,
    views: event.views,
    order: event.order
  };
}

// 范围端点按 id 配对；缺端点、重复端点、顺序颠倒都判为无法定位。
function assembleRanges(linear, issues) {
  const byId = new Map();
  for (const event of linear) {
    if (event.type !== 'boundary') continue;
    if (event.id === null || event.id === undefined) {
      issues.push({ kind: 'range-endpoint-unidentified', marker: event.marker, message: `正文中的批注范围标记缺少 id（${event.marker}），已忽略` });
      continue;
    }
    if (!byId.has(event.id)) byId.set(event.id, { starts: [], ends: [] });
    byId.get(event.id)[event.kind === 'start' ? 'starts' : 'ends'].push(event);
  }
  const ranges = new Map();
  for (const [id, { starts, ends }] of byId) {
    const broken = (kind, message) => {
      issues.push({ kind, commentId: id, message });
      ranges.set(id, null);
    };
    if (starts.length > 1) { broken('range-endpoint-duplicate', `批注 ${id} 的范围开始标记重复，无法定位`); continue; }
    if (ends.length > 1) { broken('range-endpoint-duplicate', `批注 ${id} 的范围结束标记重复，无法定位`); continue; }
    if (starts.length === 0) { broken('range-endpoint-missing', `批注 ${id} 的范围缺少开始标记，无法定位`); continue; }
    if (ends.length === 0) { broken('range-endpoint-missing', `批注 ${id} 的范围缺少结束标记，无法定位`); continue; }
    if (ends[0].order <= starts[0].order) { broken('range-endpoints-unordered', `批注 ${id} 的范围结束标记出现在开始标记之前，无法定位`); continue; }
    ranges.set(id, { start: pointOf(starts[0], linear), end: pointOf(ends[0], linear), inheritedFrom: null });
  }
  return ranges;
}

function gatherText(node, nsMap) {
  const tag = nodeTag(node);
  const { uri, local } = resolveName(tag, nsMap);
  if (uri === NS_MAIN && local === 't') return textOf(node);
  return elementChildren(node, nsMap).map((child) => gatherText(child.node, child.nsMap)).join('');
}

function readComments(root, issues) {
  const comments = [];
  const paraIdMap = new Map();
  const ambiguousParaIds = new Set();
  const seen = new Set();
  for (const child of elementChildren(root.node, root.nsMap)) {
    const { uri, local } = resolveName(nodeTag(child.node), child.nsMap);
    if (uri !== NS_MAIN || local !== 'comment') continue;
    const id = attribute(child.node, child.nsMap, NS_MAIN, 'id');
    if (id === null) {
      issues.push({ kind: 'comment-unidentified', message: '批注缺少 id，已忽略该批注' });
      continue;
    }
    if (seen.has(id)) {
      issues.push({ kind: 'comment-id-duplicate', commentId: id, message: `批注 ${id} 重复定义，仅保留第一条` });
      continue;
    }
    seen.add(id);
    const paragraphs = elementChildren(child.node, child.nsMap).filter((candidate) => {
      const name = resolveName(nodeTag(candidate.node), candidate.nsMap);
      return name.uri === NS_MAIN && name.local === 'p';
    });
    const comment = {
      id,
      author: attribute(child.node, child.nsMap, NS_MAIN, 'author'),
      date: attribute(child.node, child.nsMap, NS_MAIN, 'date'),
      initials: attribute(child.node, child.nsMap, NS_MAIN, 'initials'),
      text: paragraphs.map((paragraph) => gatherText(paragraph.node, paragraph.nsMap)).join('\n'),
      parentId: null,
      anchor: null,
      anchorBroken: false
    };
    for (const paragraph of paragraphs) {
      const paraId = attribute(paragraph.node, paragraph.nsMap, NS_W14, 'paraId');
      if (!paraId || ambiguousParaIds.has(paraId)) continue;
      if (paraIdMap.has(paraId) && paraIdMap.get(paraId) !== id) {
        paraIdMap.delete(paraId);
        ambiguousParaIds.add(paraId);
        issues.push({ kind: 'paraid-duplicate', paraId, message: `批注段落 paraId ${paraId} 被多条批注使用，相关回复关系无法确认` });
      } else {
        paraIdMap.set(paraId, id);
      }
    }
    comments.push(comment);
  }
  return { comments, paraIdMap, ambiguousParaIds };
}

// commentsExtended：paraIdParent 建立回复关系；关系外目标或缺失父级
// 只记录 issue，不靠文本匹配修复。
function readCommentsExtended(root, paraIdMap, ambiguousParaIds, issues) {
  const parentByCommentId = new Map();
  for (const child of elementChildren(root.node, root.nsMap)) {
    const { uri, local } = resolveName(nodeTag(child.node), child.nsMap);
    if (uri !== NS_W15 || local !== 'commentEx') continue;
    const paraId = attribute(child.node, child.nsMap, NS_W15, 'paraId');
    const paraIdParent = attribute(child.node, child.nsMap, NS_W15, 'paraIdParent');
    if (!paraId) {
      issues.push({ kind: 'reply-link-unidentified', message: '回复关系记录缺少 paraId，已忽略' });
      continue;
    }
    if (ambiguousParaIds.has(paraId)) {
      issues.push({ kind: 'reply-target-ambiguous', paraId, message: `回复关系指向的 paraId ${paraId} 对应多条批注，无法确认` });
      continue;
    }
    const commentId = paraIdMap.get(paraId);
    if (commentId === undefined) {
      issues.push({ kind: 'reply-target-unknown', paraId, message: `回复关系指向的 paraId ${paraId} 没有对应批注` });
      continue;
    }
    if (!paraIdParent) continue;
    if (ambiguousParaIds.has(paraIdParent)) {
      issues.push({ kind: 'reply-parent-ambiguous', commentId, paraId: paraIdParent, message: `批注 ${commentId} 的回复父级 paraId ${paraIdParent} 对应多条批注，回复关系无法确认` });
      continue;
    }
    const parentId = paraIdMap.get(paraIdParent);
    if (parentId === undefined) {
      issues.push({ kind: 'reply-parent-unknown', commentId, paraId: paraIdParent, message: `批注 ${commentId} 的回复父级 paraId ${paraIdParent} 没有对应批注，回复关系不完整` });
      continue;
    }
    if (parentByCommentId.has(commentId)) {
      issues.push({ kind: 'reply-link-duplicate', commentId, message: `批注 ${commentId} 的回复关系重复定义，仅保留第一条` });
      continue;
    }
    parentByCommentId.set(commentId, parentId);
  }
  return parentByCommentId;
}

// 回复批注沿 paraIdParent 链继承最近祖先的锚点；链断裂或成环时无法定位。
function resolveInheritedAnchor(comment, byCommentId) {
  const seen = new Set([comment.id]);
  let current = comment;
  while (current.parentId !== null) {
    if (seen.has(current.parentId)) {
      return { ok: false, kind: 'reply-cycle', message: `批注 ${comment.id} 的回复关系形成循环，无法定位` };
    }
    seen.add(current.parentId);
    const parent = byCommentId.get(current.parentId);
    if (!parent) {
      return { ok: false, kind: 'reply-parent-unknown', message: `批注 ${comment.id} 的回复父级 ${current.parentId} 不存在，回复关系不完整` };
    }
    if (parent.anchor) return { ok: true, anchor: parent.anchor, from: parent.anchor.inheritedFrom ?? parent.id };
    if (parent.anchorBroken) {
      return { ok: false, kind: 'reply-anchor-unlocatable', message: `批注 ${comment.id} 的父批注 ${parent.id} 无法定位，回复同样无法定位` };
    }
    current = parent;
  }
  return { ok: false, kind: 'reply-anchor-unlocatable', message: `批注 ${comment.id} 的回复链没有可定位的锚点` };
}

// 解析入口：返回 { status: 'parsed', ... } 或 { status: 'unreadable', issues }。
export function parseReviewPackage(bytes) {
  const issues = [];
  const unreadable = (message) => ({ status: 'unreadable', issues: [{ kind: 'package-unreadable', message }, ...issues] });
  let entries;
  try {
    entries = unzipSync(bytes);
  } catch {
    return unreadable('来源字节不是可读取的 DOCX ZIP 组合');
  }
  const names = Object.keys(entries).filter((name) => !name.endsWith('/'));
  const unsupportedParts = names.filter((name) => name.startsWith('word/') && !KNOWN_PARTS.has(name));
  for (const part of unsupportedParts) {
    issues.push({ kind: 'part-unsupported', part, message: `DOCX 部分 ${part} 不在支持范围内，保留原文但未解析` });
  }
  if (!entries['word/document.xml']) return unreadable('DOCX 组合中缺少 word/document.xml，无法读取正文');
  const documentParsed = parseXml(strFromU8(entries['word/document.xml']));
  if (!documentParsed) return unreadable('word/document.xml 不是可解析的 XML');
  const documentRoot = rootElement(documentParsed);
  if (!documentRoot || documentRoot.uri !== NS_MAIN || documentRoot.local !== 'document') {
    return unreadable('word/document.xml 不是可识别的 WordprocessingML 文档');
  }

  let relationships = {};
  if (entries['word/_rels/document.xml.rels']) {
    const relsParsed = parseXml(strFromU8(entries['word/_rels/document.xml.rels']));
    if (relsParsed) relationships = readRelationships(relsParsed);
    else issues.push({ kind: 'xml-malformed', part: 'word/_rels/document.xml.rels', message: '关系文件 word/_rels/document.xml.rels 无法解析，批注部分按固定位置读取' });
  }

  const parts = ['word/document.xml'];
  const partSources = [{ path: 'word/document.xml', via: 'fixed' }];

  const commentsPath = relationships.comments ? relationships.comments.path : 'word/comments.xml';
  let comments = [];
  let paraIdMap = new Map();
  let ambiguousParaIds = new Set();
  if (entries[commentsPath]) {
    const commentsParsed = parseXml(strFromU8(entries[commentsPath]));
    const commentsRoot = commentsParsed && rootElement(commentsParsed);
    if (!commentsRoot || commentsRoot.uri !== NS_MAIN || commentsRoot.local !== 'comments') {
      return unreadable(`批注部分 ${commentsPath} 不是可识别的 WordprocessingML 批注`);
    }
    parts.push(commentsPath);
    partSources.push({ path: commentsPath, via: relationships.comments ? 'relationship' : 'fixed', relationshipId: relationships.comments?.id ?? null });
    ({ comments, paraIdMap, ambiguousParaIds } = readComments(commentsRoot, issues));
  } else if (relationships.comments) {
    issues.push({ kind: 'referenced-part-missing', part: commentsPath, message: `关系文件引用的批注部分 ${commentsPath} 在包中不存在` });
  }

  const extendedPath = relationships.commentsExtended ? relationships.commentsExtended.path : 'word/commentsExtended.xml';
  if (entries[extendedPath]) {
    const extendedParsed = parseXml(strFromU8(entries[extendedPath]));
    const extendedRoot = extendedParsed && rootElement(extendedParsed);
    if (extendedRoot && extendedRoot.uri === NS_W15 && extendedRoot.local === 'commentsEx') {
      parts.push(extendedPath);
      partSources.push({ path: extendedPath, via: relationships.commentsExtended ? 'relationship' : 'fixed', relationshipId: relationships.commentsExtended?.id ?? null });
      const parentByCommentId = readCommentsExtended(extendedRoot, paraIdMap, ambiguousParaIds, issues);
      for (const comment of comments) comment.parentId = parentByCommentId.get(comment.id) ?? null;
    } else {
      issues.push({ kind: 'xml-malformed', part: extendedPath, message: `回复关系部分 ${extendedPath} 无法解析，回复关系未读取` });
    }
  } else if (relationships.commentsExtended) {
    issues.push({ kind: 'referenced-part-missing', part: extendedPath, message: `关系文件引用的回复关系部分 ${extendedPath} 在包中不存在` });
  }

  const { linear, segments, viewChars, paragraphs } = indexDocument(documentParsed);
  const ranges = assembleRanges(linear, issues);
  const byCommentId = new Map(comments.map((comment) => [comment.id, comment]));

  for (const id of ranges.keys()) {
    if (!byCommentId.has(id)) {
      issues.push({ kind: 'range-unknown-comment', commentId: id, message: `正文中的批注范围 ${id} 没有对应的批注，无法定位` });
    }
  }
  for (const event of linear) {
    if (event.type !== 'reference') continue;
    if (event.id === null || event.id === undefined) {
      issues.push({ kind: 'reference-unidentified', marker: event.marker, message: `正文中的批注引用缺少 id（${event.marker}），已忽略` });
    } else if (!byCommentId.has(event.id)) {
      issues.push({ kind: 'reference-unknown-comment', commentId: event.id, message: `正文中的批注引用 ${event.id} 没有对应的批注` });
    }
  }

  for (const comment of comments) {
    if (ranges.has(comment.id)) {
      comment.anchor = ranges.get(comment.id);
      comment.anchorBroken = comment.anchor === null;
    }
  }
  for (const comment of comments) {
    if (comment.anchor || comment.anchorBroken) continue;
    if (comment.parentId === null) {
      issues.push({ kind: 'comment-unanchored', commentId: comment.id, message: `批注 ${comment.id} 在正文中没有可定位的锚点范围` });
      continue;
    }
    const resolved = resolveInheritedAnchor(comment, byCommentId);
    if (resolved.ok) comment.anchor = { ...resolved.anchor, inheritedFrom: resolved.from };
    else issues.push({ kind: resolved.kind, commentId: comment.id, message: resolved.message });
  }

  return { status: 'parsed', paragraphs, parts, partSources, unsupportedParts, comments, segments, viewChars, issues };
}

function sliceView(model, anchor, view) {
  const start = anchor.start.views[view];
  const end = anchor.end.views[view];
  const text = model.viewChars[view].slice(start, end).join('');
  const segments = model.segments
    .filter((segment) => segment.views[view] !== null
      && segment.views[view] >= start
      && segment.views[view] + codePoints(segment.text) <= end)
    .map((segment) => ({ text: segment.text, status: segment.status }));
  return { start, end, text, segments };
}

function compareIds(a, b) {
  const numericA = /^\d+$/.test(a) ? Number(a) : null;
  const numericB = /^\d+$/.test(b) ? Number(b) : null;
  if (numericA !== null && numericB !== null) return numericA - numericB;
  if (a < b) return -1;
  if (a > b) return 1;
  return 0;
}

// 领域记录：每条批注一条，保留 comment id、该视图覆盖文字及原始锚点；
// 各视图范围映射一并给出，无法定位的批注 anchor/views 为 null。
export function buildRecords(model, view) {
  const orderOf = (comment) => (comment.anchor ? comment.anchor.start.order : Number.POSITIVE_INFINITY);
  const sorted = [...model.comments].sort((a, b) => orderOf(a) - orderOf(b) || compareIds(a.id, b.id));
  return sorted.map((comment) => {
    const views = comment.anchor
      ? Object.fromEntries(VIEWS.map((name) => [name, sliceView(model, comment.anchor, name)]))
      : null;
    return {
      commentId: comment.id,
      author: comment.author,
      date: comment.date,
      initials: comment.initials,
      text: comment.text,
      parentId: comment.parentId,
      view,
      coveredText: views ? views[view].text : null,
      anchor: comment.anchor
        ? {
            start: { marker: comment.anchor.start.marker, path: comment.anchor.start.path, offset: comment.anchor.start.offset },
            end: { marker: comment.anchor.end.marker, path: comment.anchor.end.path, offset: comment.anchor.end.offset },
            inheritedFrom: comment.anchor.inheritedFrom
          }
        : null,
      views
    };
  });
}
