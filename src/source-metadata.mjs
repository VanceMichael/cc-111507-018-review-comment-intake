function parseJson(bytes) { try { return JSON.parse(bytes.toString()); } catch { return null; } }
export function inspect(bytes) { return { format: 'WordprocessingML', zip: bytes.subarray(0,2).toString() === 'PK', sourceBytes: bytes.length }; }
