/* dc-lite: a very small renderer for PattEcks! pages.
   It reads the page template (with {{ holes }}, <sc-for> and <sc-if>), asks the page's
   Component for its values, builds a light tree and patches only what changed in the DOM. */
(function () {
  'use strict';
  var HOLE = /\{\{\s*([^}]+?)\s*\}\}/g;
  var ONE = /^\s*\{\{\s*([^}]+?)\s*\}\}\s*$/;

  function lookup(scope, path) {
    path = path.trim();
    if (path === 'true') return true;
    if (path === 'false') return false;
    if (path === 'null') return null;
    if (/^-?\d+(\.\d+)?$/.test(path)) return +path;
    if (/^'.*'$|^".*"$/.test(path)) return path.slice(1, -1);
    var v = scope, parts = path.split('.');
    for (var i = 0; i < parts.length; i++) { if (v == null) return undefined; v = v[parts[i]]; }
    return v;
  }
  function value(str, scope) {
    var m = str.match(ONE);
    if (m) return lookup(scope, m[1]);
    if (str.indexOf('{{') < 0) return str;
    return str.replace(HOLE, function (_, p) { var v = lookup(scope, p); return v == null ? '' : String(v); });
  }
  function child(scope, extra) { var s = Object.create(scope); for (var k in extra) s[k] = extra[k]; return s; }

  // template node -> list of virtual nodes
  function build(node, scope, out) {
    if (node.nodeType === 3) { out.push({ t: '#', text: value(node.nodeValue, scope) }); return; }
    if (node.nodeType !== 1) return;
    var tag = node.localName;
    if (tag === 'sc-for') {
      var list = value(node.getAttribute('list') || '', scope) || [];
      var as = node.getAttribute('as') || 'item';
      for (var i = 0; i < list.length; i++) {
        var extra = { $index: i }; extra[as] = list[i];
        var s = child(scope, extra);
        for (var c = node.firstChild; c; c = c.nextSibling) build(c, s, out);
      }
      return;
    }
    if (tag === 'sc-if') {
      if (value(node.getAttribute('value') || '', scope)) for (var c2 = node.firstChild; c2; c2 = c2.nextSibling) build(c2, scope, out);
      return;
    }
    var v = { t: tag, ns: node.namespaceURI, a: {}, ev: {}, c: [] };
    var attrs = node.attributes;
    for (var j = 0; j < attrs.length; j++) {
      var n = attrs[j].name;
      if (n.indexOf('hint-') === 0) continue;
      var val = value(attrs[j].value, scope);
      if (n.length > 2 && n[0] === 'o' && n[1] === 'n') { if (typeof val === 'function') v.ev[n.slice(2).toLowerCase()] = val; continue; }
      v.a[n] = val;
    }
    var kids = node.localName === 'template' ? node.content : node;
    for (var k = kids.firstChild; k; k = k.nextSibling) build(k, scope, v.c);
    out.push(v);
  }

  function setAttr(el, n, val) {
    if (n === 'style') { el.style.cssText = val == null ? '' : String(val); return; }
    if (val == null || val === false) { if (/^aria-/.test(n) && val === false) el.setAttribute(n, 'false'); else el.removeAttribute(n); return; }
    el.setAttribute(n, val === true ? (/^aria-/.test(n) ? 'true' : '') : String(val));
  }
  function listen(el, name) {
    el.addEventListener(name, function (e) { var h = el.__h && el.__h[name]; if (h) h(e); });
  }
  function create(v) {
    if (v.t === '#') return document.createTextNode(v.text);
    var el = v.ns && v.ns !== 'http://www.w3.org/1999/xhtml' ? document.createElementNS(v.ns, v.t) : document.createElement(v.t);
    for (var n in v.a) setAttr(el, n, v.a[n]);
    el.__h = {};
    for (var e in v.ev) { el.__h[e] = v.ev[e]; listen(el, e); }
    for (var i = 0; i < v.c.length; i++) el.appendChild(create(v.c[i]));
    el.__v = v;
    return el;
  }
  function patch(parent, el, o, v) {
    if (o.t !== v.t || o.ns !== v.ns) { var nw = create(v); parent.replaceChild(nw, el); return nw; }
    if (v.t === '#') { if (o.text !== v.text) el.nodeValue = v.text; return el; }
    var n;
    for (n in o.a) if (!(n in v.a)) setAttr(el, n, null);
    for (n in v.a) if (o.a[n] !== v.a[n]) setAttr(el, n, v.a[n]);
    for (n in v.ev) { if (!el.__h[n]) listen(el, n); el.__h[n] = v.ev[n]; }
    for (n in el.__h) if (!(n in v.ev)) el.__h[n] = null;
    patchKids(el, o.c, v.c);
    return el;
  }
  function patchKids(el, oc, vc) {
    var nodes = el.childNodes, i;
    for (i = 0; i < vc.length; i++) {
      if (i < oc.length) patch(el, nodes[i], oc[i], vc[i]);
      else el.appendChild(create(vc[i]));
    }
    for (i = oc.length - 1; i >= vc.length; i--) el.removeChild(nodes[i]);
  }

  function DCLogic(props) { this.props = props || {}; this.state = {}; }
  DCLogic.prototype.setState = function (u) {
    var n = typeof u === 'function' ? u(this.state, this.props) : u;
    var s = {}, k;
    for (k in this.state) s[k] = this.state[k];
    for (k in n) s[k] = n[k];
    this.state = s;
    this.__schedule && this.__schedule();
  };
  DCLogic.prototype.forceUpdate = function () { this.__schedule && this.__schedule(); };
  window.DCLogic = DCLogic;

  window.DC = {
    mount: function (tpl, root, makeClass, props) {
      var Comp = makeClass(DCLogic);
      var inst = new Comp(props || {});
      if (!inst.state) inst.state = {};
      var old = [], pending = false;
      function draw() {
        pending = false;
        var vals = inst.renderVals ? inst.renderVals() : {};
        var next = [];
        for (var c = tpl.content.firstChild; c; c = c.nextSibling) build(c, vals, next);
        patchKids(root, old, next);
        old = next;
      }
      inst.__schedule = function () { if (!pending) { pending = true; Promise.resolve().then(draw); } };
      draw();
      if (inst.componentDidMount) inst.componentDidMount();
      return inst;
    }
  };
})();
