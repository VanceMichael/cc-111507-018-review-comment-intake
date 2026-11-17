import { receiveBytes } from './http-client.mjs';
import { inspect } from './source-metadata.mjs';
import { extractReview, supportedViews } from './wordprocessingml.mjs';

export async function collect({ input, descriptor, sourceBase, store }) {
  if (!['raw', 'structured'].includes(input.mode)) return { httpStatus: 400, error: '接收方式无效' };
  const view = input.view ?? 'accepted';
  if (input.mode === 'structured' && !supportedViews.includes(view)) {
    return { httpStatus: 400, error: '文字视图无效，可选 accepted、original 或 marked' };
  }
  const source = await receiveBytes(`${sourceBase}/files/${encodeURIComponent(descriptor.entry)}`);
  const blob = await store.blob(source.bytes);
  const metadata = inspect(source.bytes);
  const common = { product: descriptor.product, source: descriptor.entry, sourceVersion: source.etag, metadata, blobs: [blob] };
  if (input.mode === 'raw') {
    const receipt = await store.save({ ...common, state: 'received', mode: 'raw', records: [], issues: [],
      coverage: { structured: false, downloadedBytes: source.bytes.length } });
    return { receipt };
  }

  let extracted;
  try {
    extracted = extractReview(source.bytes, { view });
  } catch (error) {
    return { httpStatus: 400, error: error.message };
  }
  const evidence = {
    sourceBytes: source.bytes.length,
    sourceVersion: source.etag,
    blobDigest: blob.digest,
    relationships: 'word/_rels/document.xml.rels',
  };
  if (extracted.rejected) {
    const receipt = await store.save({ ...common, state: 'rejected', mode: 'structured', view,
      records: [], issues: extracted.issues,
      coverage: { ...extracted.coverage, downloadedBytes: source.bytes.length },
      evidence });
    return { httpStatus: 201, receipt };
  }
  const receipt = await store.save({ ...common, state: extracted.state, mode: 'structured', view,
    records: extracted.records, issues: extracted.issues,
    coverage: { ...extracted.coverage, downloadedBytes: source.bytes.length },
    evidence });
  return { receipt };
}
