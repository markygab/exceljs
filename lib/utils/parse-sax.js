const {SaxesParser} = require('saxes');
const {PassThrough} = require('readable-stream');
const {bufferToString} = require('./browser-buffer-decode');

const spreadsheetmlNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function normalizeElementName(node) {
  if (node && node.uri === spreadsheetmlNamespace && node.local) {
    return node.local;
  }

  return node && node.name;
}

function flattenAttributes(attributes = {}) {
  return Object.keys(attributes).reduce((result, key) => {
    result[key] = attributes[key].value;
    return result;
  }, {});
}

function normalizeOpenTag(node) {
  return {
    ...node,
    name: normalizeElementName(node),
    attributes: flattenAttributes(node.attributes),
  };
}

function normalizeCloseTag(node) {
  return typeof node === 'object' && node !== null
    ? {
        ...node,
        name: normalizeElementName(node),
      }
    : node;
}

module.exports = async function* (iterable) {
  // TODO: Remove once node v8 is deprecated
  // Detect and upgrade old streams
  if (iterable.pipe && !iterable[Symbol.asyncIterator]) {
    iterable = iterable.pipe(new PassThrough());
  }
  const saxesParser = new SaxesParser({xmlns: true});
  let error;
  saxesParser.on('error', err => {
    error = err;
  });
  let events = [];
  saxesParser.on('opentag', value =>
    events.push({eventType: 'opentag', value: normalizeOpenTag(value)})
  );
  saxesParser.on('text', value => events.push({eventType: 'text', value}));
  saxesParser.on('closetag', value =>
    events.push({eventType: 'closetag', value: normalizeCloseTag(value)})
  );
  for await (const chunk of iterable) {
    saxesParser.write(bufferToString(chunk));
    // saxesParser.write and saxesParser.on() are synchronous,
    // so we can only reach the below line once all events have been emitted
    if (error) throw error;
    // As a performance optimization, we gather all events instead of passing
    // them one by one, which would cause each event to go through the event queue
    yield events;
    events = [];
  }
};
