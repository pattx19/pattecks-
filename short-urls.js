/* PattEcks! short URLs: show /collection, /archive, /henley, /denim in the address bar instead of #hashes. */
(function () {
  var HOME = { collection: '/collection', archive: '/archive' };
  var PROD = { henley: '/henley', denim: '/denim' };
  var path = location.pathname.replace(/\.html$/, '').replace(/\/index$/, '/');
  var onHome = path === '/' || path === '';
  var onProduct = path === '/product';

  function whenReady(id, fn, tries) {
    var el = document.getElementById(id);
    if (el) return fn(el);
    if ((tries || 0) < 60) setTimeout(function () { whenReady(id, fn, (tries || 0) + 1); }, 50);
  }
  function tidy() {
    var h = location.hash.replace('#', '');
    if (onHome && HOME[h]) {
      whenReady(h, function (el) {
        el.scrollIntoView();
        history.replaceState(history.state, '', HOME[h]);
      });
    } else if (onProduct && PROD[h]) {
      // let the product page read the hash first, then shorten the address
      setTimeout(function () { history.replaceState(history.state, '', PROD[h]); }, 600);
    }
  }
  tidy();
  window.addEventListener('hashchange', tidy);

  if (onHome) {
    document.addEventListener('click', function (e) {
      var a = e.target.closest && e.target.closest('a[href]');
      if (!a) return;
      var href = a.getAttribute('href').replace(/^\/?#/, '').replace(/^\//, '');
      if (!HOME[href]) return;
      var el = document.getElementById(href);
      if (!el) return;
      e.preventDefault();
      el.scrollIntoView({ behavior: 'smooth' });
      history.pushState({ s: href }, '', HOME[href]);
    });
    window.addEventListener('popstate', function () {
      var p = location.pathname.replace('/', '');
      var el = HOME[p] && document.getElementById(p);
      if (el) el.scrollIntoView({ behavior: 'smooth' }); else if (!location.hash) window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }
})();
