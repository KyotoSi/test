// src/config/xml.js
//
// Тонкая обёртка над @xmldom/xmldom для чтения конфигурации и метаданных.
// Работаем по local-name (без учёта префиксов пространств имён),
// чтобы одинаково читать <md:EntityDescriptor> и <EntityDescriptor>.

const { DOMParser } = require('@xmldom/xmldom');

function parseXml(text, sourceName = 'XML') {
  const errors = [];
  let doc;
  try {
    doc = new DOMParser({
      onError: (level, msg) => {
        if (level !== 'warning') errors.push(String(msg).trim());
      },
    }).parseFromString(String(text), 'text/xml');
  } catch (e) {
    errors.push(e.message);
  }
  if (errors.length || !doc || !doc.documentElement) {
    const err = new Error(`${sourceName}: некорректный XML — ${errors[0] || 'пустой документ'}`);
    err.xmlErrors = errors;
    throw err;
  }
  return doc;
}

function localName(node) {
  return node.localName || String(node.nodeName).replace(/^.*:/, '');
}

/** Дочерние элементы (только непосредственные) с заданным local-name. */
function children(el, name) {
  const out = [];
  if (!el) return out;
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (!name || localName(n) === name)) out.push(n);
  }
  return out;
}

function child(el, name) {
  return children(el, name)[0] || null;
}

/** Все потомки с заданным local-name (обход в глубину). */
function descendants(el, name) {
  const out = [];
  const walk = (node) => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1) {
        if (localName(n) === name) out.push(n);
        walk(n);
      }
    }
  };
  if (el) walk(el);
  return out;
}

function attr(el, name) {
  if (!el || !el.hasAttribute(name)) return undefined;
  return el.getAttribute(name);
}

function text(el) {
  if (!el) return undefined;
  return String(el.textContent || '').trim();
}

module.exports = { parseXml, children, child, descendants, attr, text, localName };
