// Minimal XML DOM, just enough to read and rewrite WordprocessingML parts
// such as word/document.xml. Uses no dependencies, so every line can be audited.
//
// Supported: elements, attributes, namespaces (xmlns / xmlns:prefix), text,
// the five predefined entities and numeric character references, CDATA,
// comments, processing instructions and a DOCTYPE without an internal subset.
// Comments, processing instructions and DOCTYPEs are kept byte for byte.

export const ELEMENT = 1;
export const TEXT = 3;
export const RAW = 0; // comment, processing instruction or DOCTYPE

const XMLNS_URI = 'http://www.w3.org/2000/xmlns/';
const XML_URI = 'http://www.w3.org/XML/1998/namespace';

class Node {
    constructor(type) {
        this.type = type;
        this.parent = null;
    }
}

export class Text extends Node {
    constructor(value) {
        super(TEXT);
        this.value = value;
    }
    clone() {
        return new Text(this.value);
    }
}

export class Raw extends Node {
    constructor(source) {
        super(RAW);
        this.source = source;
    }
    clone() {
        return new Raw(this.source);
    }
}

export class Element extends Node {
    /**
     * @param {string} name  qualified name, e.g. "w:t"
     * @param {string|null} ns  resolved namespace URI
     * @param {Map<string,string>} scope  prefix -> namespace URI bindings in effect
     */
    constructor(name, ns, scope) {
        super(ELEMENT);
        this.name = name;
        this.ns = ns;
        this.scope = scope;
        this.attrs = new Map(); // keeps source order
        this.children = [];
    }

    get localName() {
        const colon = this.name.indexOf(':');
        return colon < 0 ? this.name : this.name.slice(colon + 1);
    }

    getAttribute(name) {
        return this.attrs.has(name) ? this.attrs.get(name) : null;
    }

    setAttribute(name, value) {
        this.attrs.set(name, String(value));
    }

    /** Concatenated text of all descendant text nodes. */
    get textContent() {
        let out = '';
        for (const child of this.children) {
            if (child.type === TEXT) out += child.value;
            else if (child.type === ELEMENT) out += child.textContent;
        }
        return out;
    }

    set textContent(value) {
        for (const child of this.children) child.parent = null;
        this.children = [];
        if (value) this.appendChild(new Text(value));
    }

    appendChild(child) {
        if (child.parent) child.parent.removeChild(child);
        child.parent = this;
        this.children.push(child);
        return child;
    }

    insertBefore(child, reference) {
        if (child.parent) child.parent.removeChild(child);
        const index = this.children.indexOf(reference);
        if (index < 0) throw new Error('insertBefore: reference is not a child');
        child.parent = this;
        this.children.splice(index, 0, child);
        return child;
    }

    removeChild(child) {
        const index = this.children.indexOf(child);
        if (index < 0) throw new Error('removeChild: not a child');
        this.children.splice(index, 1);
        child.parent = null;
        return child;
    }

    /** Deep copy. The copy has no parent. */
    clone() {
        const copy = new Element(this.name, this.ns, this.scope);
        for (const [key, value] of this.attrs) copy.attrs.set(key, value);
        for (const child of this.children) copy.appendChild(child.clone());
        return copy;
    }

    /** Every descendant element (document order) whose namespace and local name match. */
    findAll(ns, localName, out = []) {
        for (const child of this.children) {
            if (child.type !== ELEMENT) continue;
            if (child.ns === ns && child.localName === localName) out.push(child);
            child.findAll(ns, localName, out);
        }
        return out;
    }
}

export class XmlDocument {
    constructor() {
        this.children = []; // prolog nodes, the root element, trailing nodes
    }

    get root() {
        return this.children.find((child) => child.type === ELEMENT);
    }

    /** The prefix the root element binds to a namespace URI, or null. */
    prefixFor(ns) {
        for (const [prefix, uri] of this.root.scope) {
            if (uri === ns && prefix !== '') return prefix;
        }
        return null;
    }

    createElement(qualifiedName) {
        const root = this.root;
        const colon = qualifiedName.indexOf(':');
        const prefix = colon < 0 ? '' : qualifiedName.slice(0, colon);
        return new Element(qualifiedName, root.scope.get(prefix) ?? null, root.scope);
    }
}

// ---------------------------------------------------------------- parsing

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

function decodeEntities(text) {
    if (!text.includes('&')) return text;
    return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (whole, body) => {
        if (body[0] === '#') {
            const code = body[1] === 'x' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
            return String.fromCodePoint(code);
        }
        if (body in ENTITIES) return ENTITIES[body];
        throw new Error(`Unsupported XML entity ${whole}`);
    });
}

const NAME_END = /[\s/>=]/;

export function parseXml(source) {
    if (source.charCodeAt(0) === 0xfeff) source = source.slice(1); // byte order mark

    const doc = new XmlDocument();
    const rootScope = new Map([['xml', XML_URI], ['xmlns', XMLNS_URI]]);
    const stack = []; // open elements
    const append = (node) => {
        if (stack.length) stack[stack.length - 1].appendChild(node);
        else doc.children.push(node);
    };
    const fail = (message, at) => {
        throw new Error(`XML parse error at offset ${at}: ${message}`);
    };

    let i = 0;
    const length = source.length;
    while (i < length) {
        if (source[i] !== '<') {
            let end = source.indexOf('<', i);
            if (end < 0) end = length;
            const value = decodeEntities(source.slice(i, end));
            if (stack.length) append(new Text(value));
            else if (value.trim()) fail('text outside the root element', i);
            else doc.children.push(new Raw(source.slice(i, end))); // whitespace between prolog nodes
            i = end;
            continue;
        }

        if (source.startsWith('<?', i)) {
            const end = source.indexOf('?>', i);
            if (end < 0) fail('unterminated processing instruction', i);
            append(new Raw(source.slice(i, end + 2)));
            i = end + 2;
        } else if (source.startsWith('<!--', i)) {
            const end = source.indexOf('-->', i);
            if (end < 0) fail('unterminated comment', i);
            append(new Raw(source.slice(i, end + 3)));
            i = end + 3;
        } else if (source.startsWith('<![CDATA[', i)) {
            const end = source.indexOf(']]>', i);
            if (end < 0) fail('unterminated CDATA section', i);
            append(new Text(source.slice(i + 9, end)));
            i = end + 3;
        } else if (source.startsWith('<!', i)) {
            const end = source.indexOf('>', i);
            if (end < 0) fail('unterminated declaration', i);
            if (source.slice(i, end).includes('[')) fail('DOCTYPE internal subsets are not supported', i);
            append(new Raw(source.slice(i, end + 1)));
            i = end + 1;
        } else if (source[i + 1] === '/') {
            const end = source.indexOf('>', i);
            if (end < 0) fail('unterminated end tag', i);
            const name = source.slice(i + 2, end).trim();
            const open = stack.pop();
            if (!open || open.name !== name) fail(`unexpected </${name}>`, i);
            i = end + 1;
        } else {
            // Start tag: name, then attributes, then ">" or "/>".
            let j = i + 1;
            while (j < length && !NAME_END.test(source[j])) j++;
            const name = source.slice(i + 1, j);
            if (!name) fail('missing element name', i);

            const attrs = [];
            let selfClosing = false;
            for (;;) {
                while (j < length && /\s/.test(source[j])) j++;
                if (j >= length) fail('unterminated start tag', i);
                if (source[j] === '>') {
                    j++;
                    break;
                }
                if (source.startsWith('/>', j)) {
                    selfClosing = true;
                    j += 2;
                    break;
                }
                const nameStart = j;
                while (j < length && !NAME_END.test(source[j])) j++;
                const attrName = source.slice(nameStart, j);
                while (j < length && /\s/.test(source[j])) j++;
                if (source[j] !== '=') fail(`attribute ${attrName} has no value`, j);
                j++;
                while (j < length && /\s/.test(source[j])) j++;
                const quote = source[j];
                if (quote !== '"' && quote !== "'") fail(`attribute ${attrName} is not quoted`, j);
                const valueEnd = source.indexOf(quote, j + 1);
                if (valueEnd < 0) fail(`unterminated value for ${attrName}`, j);
                attrs.push([attrName, decodeEntities(source.slice(j + 1, valueEnd))]);
                j = valueEnd + 1;
            }

            // Namespace declarations on this element extend the parent's scope.
            const parentScope = stack.length ? stack[stack.length - 1].scope : rootScope;
            let scope = parentScope;
            for (const [attrName, value] of attrs) {
                if (attrName === 'xmlns' || attrName.startsWith('xmlns:')) {
                    if (scope === parentScope) scope = new Map(parentScope);
                    scope.set(attrName === 'xmlns' ? '' : attrName.slice(6), value);
                }
            }
            const colon = name.indexOf(':');
            const prefix = colon < 0 ? '' : name.slice(0, colon);
            const element = new Element(name, scope.get(prefix) ?? null, scope);
            for (const [attrName, value] of attrs) element.attrs.set(attrName, value);

            append(element);
            if (!selfClosing) stack.push(element);
            i = j;
        }
    }

    if (stack.length) fail(`<${stack[stack.length - 1].name}> is never closed`, length);
    if (!doc.root) fail('no root element', 0);
    return doc;
}

// ---------------------------------------------------------- serializing

const escapeText = (value) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const escapeAttr = (value) =>
    value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/"/g, '&quot;')
        .replace(/\t/g, '&#9;')
        .replace(/\n/g, '&#10;')
        .replace(/\r/g, '&#13;');

function serializeNode(node, out) {
    if (node.type === TEXT) {
        out.push(escapeText(node.value));
    } else if (node.type === RAW) {
        out.push(node.source);
    } else {
        let tag = '<' + node.name;
        for (const [key, value] of node.attrs) tag += ` ${key}="${escapeAttr(value)}"`;
        if (node.children.length === 0) {
            out.push(tag + '/>');
        } else {
            out.push(tag + '>');
            for (const child of node.children) serializeNode(child, out);
            out.push(`</${node.name}>`);
        }
    }
}

export function serializeXml(doc) {
    const out = [];
    for (const node of doc.children) serializeNode(node, out);
    return out.join('');
}
