// Menü auf dem Handy öffnen und schließen
var knopf = document.querySelector('.nav-knopf');
var nav = document.getElementById('nav');
if (knopf && nav) {
  knopf.addEventListener('click', function () {
    var offen = nav.classList.toggle('offen');
    knopf.setAttribute('aria-expanded', offen ? 'true' : 'false');
  });
  nav.addEventListener('click', function (e) {
    if (e.target.tagName === 'A') {
      nav.classList.remove('offen');
      knopf.setAttribute('aria-expanded', 'false');
    }
  });
}

// Noch nicht eingetragene Links zeigen "Link folgt" statt ins Leere zu führen
document.querySelectorAll('a[href^="HIER-"]').forEach(function (a) {
  a.classList.add('link-offen');
  a.setAttribute('aria-disabled', 'true');
  a.removeAttribute('target');
  a.addEventListener('click', function (e) { e.preventDefault(); });
  var hinweis = document.createElement('span');
  hinweis.className = 'link-hinweis';
  hinweis.textContent = 'Link folgt';
  var icon = a.querySelector('svg');
  if (icon) { icon.replaceWith(hinweis); } else { a.appendChild(hinweis); }
});

// IBAN in die Zwischenablage kopieren
var kopieren = document.getElementById('iban-kopieren');
if (kopieren) {
  kopieren.addEventListener('click', function () {
    var iban = document.getElementById('iban').textContent.replace(/\s/g, '');
    var fertig = function () {
      kopieren.textContent = 'IBAN kopiert';
      setTimeout(function () { kopieren.textContent = 'IBAN kopieren'; }, 2500);
    };
    var auswaehlen = function () {
      var r = document.createRange();
      r.selectNodeContents(document.getElementById('iban'));
      var s = window.getSelection();
      s.removeAllRanges();
      s.addRange(r);
      kopieren.textContent = 'IBAN markiert';
    };
    if (navigator.clipboard) {
      navigator.clipboard.writeText(iban).then(fertig, auswaehlen);
    } else {
      auswaehlen();
    }
  });
}
