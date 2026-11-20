'use strict';

/**
 * test_rich_text.js (rcballot-cck.3 / .1)
 *
 * Unit tests for the linkify / first-paragraph helpers in static-pages/src/index.html. The
 * helpers live inline in that page (no module system), so they are extracted from between the
 * <rich-text-helpers> markers and evaluated in a vm sandbox with a minimal fake DOM.
 */

const assert = require('assert');
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'static-pages', 'src', 'index.html'), 'utf8');
const start = html.indexOf('// <rich-text-helpers>');
const end = html.indexOf('// </rich-text-helpers>');
assert.ok(start !== -1 && end > start, 'helper markers present in index.html');

function fakeNode(tag, text) {
  return {
    tag, text: text || '', children: [], attrs: {},
    set textContent(v) { this.text = v; this.children = []; },
    get textContent() { return this.text; },
    set innerHTML(v) { throw new Error('innerHTML must never be used for admin text'); },
    appendChild(c) { this.children.push(c); return c; },
    setAttribute(k, v) { this.attrs[k] = v; },
  };
}
const sandbox = {
  document: {
    createElement: (t) => fakeNode(t),
    createTextNode: (t) => fakeNode('#text', t),
  },
};
vm.createContext(sandbox);
vm.runInContext(html.slice(start, end) + '\nthis.api = { parseRichText_, setRich_, setRichWithLink_, safeHttpUrl_, firstParagraph_, remainingParagraphs_ };', sandbox);
const { parseRichText_, setRich_, setRichWithLink_, safeHttpUrl_, firstParagraph_, remainingParagraphs_ } = sandbox.api;

function plain(v) { return JSON.parse(JSON.stringify(v)); }

function testPlainTextIsOneTextSegment() {
  assert.deepEqual(plain(parseRichText_('Hello\n\nworld')), [{ type: 'text', value: 'Hello\n\nworld' }]);
  assert.deepEqual(plain(parseRichText_('')), []);
  assert.deepEqual(plain(parseRichText_(null)), []);
}

function testBareUrlsBecomeLinks() {
  assert.deepEqual(plain(parseRichText_('See https://example.com/a?b=1 now')), [
    { type: 'text', value: 'See ' }, { type: 'link', value: 'https://example.com/a?b=1' }, { type: 'text', value: ' now' },
  ]);
  assert.deepEqual(plain(parseRichText_('go http://a.b/c.')), [
    { type: 'text', value: 'go ' }, { type: 'link', value: 'http://a.b/c' }, { type: 'text', value: '.' },
  ]);
  assert.deepEqual(plain(parseRichText_('(https://x.org/p)')), [
    { type: 'text', value: '(' }, { type: 'link', value: 'https://x.org/p' }, { type: 'text', value: ')' },
  ]);
  const two = plain(parseRichText_('https://a.com\nhttps://b.com'));
  assert.deepEqual(two.map((s) => s.type), ['link', 'text', 'link']);
}

function testOnlyHttpSchemesAreLinked() {
  ['javascript:alert(1)', 'ftp://x.com', 'data:text/html,<b>x</b>', 'www.example.com'].forEach((t) => {
    assert.deepEqual(plain(parseRichText_(t)).filter((s) => s.type === 'link'), [], t);
  });
}

function testCrlfNormalised() {
  assert.deepEqual(plain(parseRichText_('a\r\nb')), [{ type: 'text', value: 'a\nb' }]);
}

function testSetRichBuildsNodesAndNeverInjectsHtml() {
  const el = fakeNode('p');
  const evil = '<img src=x onerror=alert(1)> <script>x</script> https://ok.example/p?a=1&b=2"><b>';
  setRich_(el, evil);
  const links = el.children.filter((c) => c.tag === 'a');
  assert.equal(links.length, 1);
  assert.equal(links[0].href, 'https://ok.example/p?a=1&b=2');
  assert.equal(links[0].text, 'https://ok.example/p?a=1&b=2');
  assert.equal(links[0].target, '_blank');
  assert.equal(links[0].rel, 'noopener noreferrer');
  const texts = el.children.filter((c) => c.tag === '#text').map((c) => c.text).join('');
  assert.ok(texts.indexOf('<img src=x onerror=alert(1)>') !== -1, 'markup stays literal text');
  assert.ok(texts.indexOf('<script>') !== -1);
  // Re-render replaces rather than appends.
  setRich_(el, 'plain');
  assert.equal(el.children.length, 1);
}

function testFirstParagraph() {
  assert.equal(firstParagraph_('One\nstill one\n\nTwo\n\nThree'), 'One\nstill one');
  assert.equal(firstParagraph_('Only paragraph'), 'Only paragraph');
  assert.equal(firstParagraph_('  A  \r\n\r\nB'), 'A');
  assert.equal(firstParagraph_('A\n   \nB'), 'A'); // whitespace-only line counts as blank
  assert.equal(firstParagraph_('\n\nLeading blank\n\nNext'), 'Leading blank');
  assert.equal(firstParagraph_(''), '');
  assert.equal(firstParagraph_(undefined), '');
}

function testRemainingParagraphs() {
  assert.equal(remainingParagraphs_('One\n\nTwo\n\nThree'), 'Two\n\nThree');
  assert.equal(remainingParagraphs_('Only'), '');
  assert.equal(remainingParagraphs_('A\r\n\r\nB'), 'B');
}

/** rcballot-9rc: link after the Description; label falls back to URL; only http(s); DOM-built. */
function testDescriptionLink() {
  const el = fakeNode('p');
  setRichWithLink_(el, 'Intro text', 'Original responses', 'https://docs.example.com/d/1');
  const a = el.children[el.children.length - 1];
  assert.equal(a.tag, 'a');
  assert.equal(a.href, 'https://docs.example.com/d/1');
  assert.equal(a.textContent, 'Original responses');
  assert.equal(a.target, '_blank');
  assert.equal(a.rel, 'noopener noreferrer');
  assert.equal(el.children[el.children.length - 2].text, '\n'); // own line after the text

  // Blank text falls back to the URL; blank Description gets no leading separator.
  const el2 = fakeNode('p');
  setRichWithLink_(el2, '', '  ', 'http://a.example/x');
  assert.equal(el2.children.length, 1);
  assert.equal(el2.children[0].textContent, 'http://a.example/x');

  // Missing/blank/non-http(s) URL: page unchanged (identical to setRich_).
  ['', null, undefined, 'javascript:alert(1)', 'ftp://x.com', 'data:text/html,x', 'www.example.com', 'https://a b.com', 'https://x.com/"><b>'].forEach((u) => {
    const e = fakeNode('p');
    setRichWithLink_(e, 'Intro', 'Label', u);
    assert.equal(e.children.length, 1, String(u));
    assert.equal(e.children[0].tag, '#text', String(u));
  });
  assert.equal(safeHttpUrl_('  https://ok.example/p  '), 'https://ok.example/p');
  assert.equal(safeHttpUrl_('javascript:alert(1)'), '');
  // Re-render replaces rather than appends.
  setRichWithLink_(el, 'again', '', '');
  assert.equal(el.children.length, 1);
}

testDescriptionLink();
testPlainTextIsOneTextSegment();
testBareUrlsBecomeLinks();
testOnlyHttpSchemesAreLinked();
testCrlfNormalised();
testSetRichBuildsNodesAndNeverInjectsHtml();
testFirstParagraph();
testRemainingParagraphs();
console.log('test_rich_text: all tests passed');
