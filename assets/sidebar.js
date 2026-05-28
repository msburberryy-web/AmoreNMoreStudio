(function () {
  'use strict';

  var R = './';
  var sc = document.currentScript;
  if (sc && sc.getAttribute('data-root')) { R = sc.getAttribute('data-root'); }

  // ── Language (persisted) ──
  var lang = localStorage.getItem('am-lang') || 'en';
  document.documentElement.setAttribute('data-lang', lang);

  // ── Inject CSS ──
  var style = document.createElement('style');
  style.textContent = '\
    #am-sb-btn {\
      position:fixed; top:14px; right:18px; z-index:1100;\
      width:40px; height:40px;\
      background:rgba(247,242,234,.94);\
      border:1px solid rgba(201,169,110,.3);\
      border-radius:2px; cursor:pointer;\
      display:flex; flex-direction:column;\
      align-items:center; justify-content:center; gap:5px;\
      backdrop-filter:blur(8px); -webkit-backdrop-filter:blur(8px);\
      transition:background .2s, border-color .2s;\
      box-shadow:0 2px 12px rgba(0,0,0,.08);\
    }\
    #am-sb-btn:hover { background:rgba(253,250,245,1); border-color:rgba(201,169,110,.55); }\
    #am-sb-btn span {\
      display:block; width:16px; height:1.5px;\
      background:#6B5240; border-radius:1px;\
      transition:transform .28s ease, opacity .28s ease;\
      transform-origin:center;\
    }\
    #am-sb-btn.open span:nth-child(1) { transform:translateY(6.5px) rotate(45deg); }\
    #am-sb-btn.open span:nth-child(2) { opacity:0; transform:scaleX(0); }\
    #am-sb-btn.open span:nth-child(3) { transform:translateY(-6.5px) rotate(-45deg); }\
    #am-sb-ov {\
      position:fixed; inset:0; z-index:1050;\
      background:rgba(26,18,8,.28);\
      opacity:0; pointer-events:none;\
      transition:opacity .3s;\
      backdrop-filter:blur(2px); -webkit-backdrop-filter:blur(2px);\
    }\
    #am-sb-ov.open { opacity:1; pointer-events:all; }\
    #am-sb-dr {\
      position:fixed; top:0; right:0;\
      width:min(275px, 88vw); height:100%;\
      z-index:1080;\
      background:#FDFAF5;\
      border-left:1px solid rgba(201,169,110,.2);\
      display:flex; flex-direction:column;\
      transform:translateX(100%);\
      transition:transform .32s cubic-bezier(.4,0,.2,1);\
      overflow-y:auto;\
      box-shadow:-6px 0 36px rgba(0,0,0,.1);\
    }\
    #am-sb-dr.open { transform:translateX(0); }\
    .am-sb-head {\
      padding:22px 22px 16px;\
      border-bottom:1px solid rgba(201,169,110,.14);\
      display:flex; align-items:center; justify-content:space-between;\
      flex-shrink:0;\
    }\
    .am-sb-logo { max-width:112px; height:auto; mix-blend-mode:multiply; opacity:.88; }\
    .am-sb-cls {\
      background:none; border:none; cursor:pointer;\
      font-size:16px; color:#8B7A60; line-height:1;\
      padding:6px 8px; transition:color .2s; flex-shrink:0;\
    }\
    .am-sb-cls:hover { color:#1A1208; }\
    .am-sb-ey {\
      padding:18px 22px 6px;\
      font-size:9px; letter-spacing:4px; text-transform:uppercase;\
      color:#A8854A; opacity:.7;\
    }\
    .am-sb-ey .sb-en { font-family:"Cormorant Garamond",serif; }\
    .am-sb-ey .sb-my { font-family:"Noto Sans Myanmar",sans-serif; letter-spacing:1px; }\
    .am-sb-nav { list-style:none; padding:0 0 6px; }\
    .am-sb-nav li a {\
      display:block;\
      padding:13px 22px;\
      font-family:"Cormorant Garamond",serif;\
      font-size:13px; letter-spacing:2px; text-transform:uppercase;\
      text-decoration:none; color:#4A3018;\
      border-left:2px solid transparent;\
      transition:color .2s, background .2s, border-color .2s;\
    }\
    .am-sb-nav li a:hover, .am-sb-nav li a.am-cur {\
      color:#A8854A; background:rgba(201,169,110,.07);\
      border-left-color:rgba(201,169,110,.5);\
    }\
    .am-sb-nav .sb-en { font-family:"Cormorant Garamond",serif; letter-spacing:2px; text-transform:uppercase; }\
    .am-sb-nav .sb-my { font-family:"Noto Sans Myanmar",sans-serif; letter-spacing:0; text-transform:none; font-size:13px; }\
    .am-sb-rule {\
      height:1px; margin:6px 22px 2px;\
      background:linear-gradient(90deg,rgba(201,169,110,.2),transparent);\
      flex-shrink:0;\
    }\
    .am-sb-lang {\
      padding:12px 22px;\
      display:flex; align-items:center; gap:8px;\
      border-top:1px solid rgba(201,169,110,.12);\
      flex-shrink:0;\
    }\
    .am-sb-ll {\
      font-size:9px; letter-spacing:3px; text-transform:uppercase;\
      color:#A8854A; opacity:.7; margin-right:2px;\
    }\
    .am-sb-ll .sb-en { font-family:"Cormorant Garamond",serif; }\
    .am-sb-ll .sb-my { font-family:"Noto Sans Myanmar",sans-serif; letter-spacing:0.5px; text-transform:none; }\
    .am-sb-lb {\
      font-size:11px;\
      background:none; border:1px solid transparent;\
      padding:5px 10px; cursor:pointer;\
      color:#8B7A60; border-radius:1px;\
      transition:color .2s, border-color .2s, background .2s;\
    }\
    .am-sb-lb[data-l="en"] { font-family:"Cormorant Garamond",serif; letter-spacing:1px; }\
    .am-sb-lb[data-l="my"] { font-family:"Noto Sans Myanmar",sans-serif; letter-spacing:0; }\
    .am-sb-lb.active { color:#A8854A; border-color:rgba(201,169,110,.4); background:rgba(201,169,110,.09); }\
    .am-sb-ls { color:rgba(180,150,100,.35); font-size:9px; }\
    .am-sb-foot {\
      margin-top:auto; padding:18px 22px 26px;\
      border-top:1px solid rgba(201,169,110,.1);\
      text-align:center; flex-shrink:0;\
    }\
    .am-sb-foot img {\
      max-width:72px; height:auto;\
      mix-blend-mode:multiply; opacity:.4;\
      display:block; margin:0 auto 7px;\
    }\
    .am-sb-foot p {\
      font-family:"Cormorant Garamond",serif;\
      font-size:9px; letter-spacing:3px; text-transform:uppercase;\
      color:rgba(107,82,64,.38);\
    }\
    #am-pg-foot {\
      text-align:center; padding:44px 24px 32px;\
      border-top:1px solid rgba(201,169,110,.14);\
      background:rgba(201,169,110,.022);\
    }\
    #am-pg-foot img {\
      max-width:80px; height:auto;\
      mix-blend-mode:multiply; opacity:.35;\
      display:block; margin:0 auto 10px;\
    }\
    #am-pg-foot p {\
      font-size:9px; letter-spacing:4px; text-transform:uppercase;\
      color:rgba(107,82,64,.38);\
    }\
    #am-pg-foot p .sb-en { font-family:"Cormorant Garamond",serif; letter-spacing:4px; }\
    #am-pg-foot p .sb-my { font-family:"Noto Sans Myanmar",sans-serif; letter-spacing:0.5px; text-transform:none; }\
    html[data-lang="my"] .sb-en { display:none !important; }\
    html[data-lang="en"] .sb-my { display:none !important; }\
    .sb-my {\
      font-family:"Noto Sans Myanmar",sans-serif !important;\
      letter-spacing:0 !important;\
      text-transform:none !important;\
    }\
    html[data-lang="my"] .headline,\
    html[data-lang="my"] .hero-title,\
    html[data-lang="my"] .cta-title,\
    html[data-lang="my"] .zi-title,\
    html[data-lang="my"] .zi-intro-title,\
    html[data-lang="my"] .services-title,\
    html[data-lang="my"] .section-title,\
    html[data-lang="my"] .hero-sub .sb-my,\
    html[data-lang="my"] .brand-sub .sb-my {\
      font-family:"Noto Serif Myanmar",serif !important;\
      letter-spacing:0 !important;\
      text-transform:none !important;\
    }\
  ';
  document.head.appendChild(style);

  // ── Inject HTML ──
  var wrap = document.createElement('div');
  wrap.id = 'am-sb-wrap';
  wrap.innerHTML =
    '<button id="am-sb-btn" aria-label="Menu"><span></span><span></span><span></span></button>' +
    '<div id="am-sb-ov"></div>' +
    '<nav id="am-sb-dr" role="navigation" aria-label="Site navigation">' +
      '<div class="am-sb-head">' +
        '<img class="am-sb-logo" src="' + R + 'amore-logo.PNG" alt="Amorè N\' More">' +
        '<button class="am-sb-cls" id="am-sb-cls" aria-label="Close">&#10005;</button>' +
      '</div>' +
      '<div class="am-sb-ey"><span class="sb-en">Navigation</span><span class="sb-my">လမ်းညွှန်</span></div>' +
      '<ul class="am-sb-nav">' +
        '<li><a href="' + R + '"><span class="sb-en">Home</span><span class="sb-my">မူလစာမျက်နှာ</span></a></li>' +
        '<li><a href="' + R + 'bookings/"><span class="sb-en">Book Appointment</span><span class="sb-my">ချိန်းဆိုရန်</span></a></li>' +
        '<li><a href="' + R + 'studio-room/"><span class="sb-en">Our Studio</span><span class="sb-my">ကျွန်ုပ်တို့ Studio</span></a></li>' +
        '<li><a href="' + R + 'about/"><span class="sb-en">About &amp; Partners</span><span class="sb-my">ကျွန်ုပ်တို့အကြောင်း</span></a></li>' +
        '<li><a href="' + R + 'magazine/"><span class="sb-en">The Magazine</span><span class="sb-my">မဂ္ဂဇင်း</span></a></li>' +
      '</ul>' +
      '<div class="am-sb-rule"></div>' +
      '<div class="am-sb-lang">' +
        '<span class="am-sb-ll"><span class="sb-en">Lang</span><span class="sb-my">ဘာသာ</span></span>' +
        '<button class="am-sb-lb" data-l="en">English</button>' +
        '<span class="am-sb-ls">&middot;</span>' +
        '<button class="am-sb-lb" data-l="my">မြန်မာ</button>' +
      '</div>' +
      '<div class="am-sb-foot">' +
        '<img src="' + R + 'amore-logo.PNG" alt="">' +
        '<p>Amorè N\' More &nbsp;&middot;&nbsp; Tokyo</p>' +
      '</div>' +
    '</nav>';
  document.body.appendChild(wrap);

  // Page footer
  var foot = document.createElement('footer');
  foot.id = 'am-pg-foot';
  foot.innerHTML =
    '<img src="' + R + 'amore-logo.PNG" alt="Amorè N\' More">' +
    '<p>' +
      '<span class="sb-en">&copy; 2025 Amorè N\' More Wedding Studio &nbsp;&middot;&nbsp; Sengoku, Tokyo</span>' +
      '<span class="sb-my">&copy; 2025 Amorè N\' More Wedding Studio &nbsp;&middot;&nbsp; Sengoku, Tokyo</span>' +
    '</p>';
  document.body.appendChild(foot);

  // ── Controls ──
  var btn = document.getElementById('am-sb-btn');
  var ov  = document.getElementById('am-sb-ov');
  var dr  = document.getElementById('am-sb-dr');
  var cls = document.getElementById('am-sb-cls');

  function openMenu()  {
    btn.classList.add('open');
    ov.classList.add('open');
    dr.classList.add('open');
    document.body.style.overflow = 'hidden';
  }
  function closeMenu() {
    btn.classList.remove('open');
    ov.classList.remove('open');
    dr.classList.remove('open');
    document.body.style.overflow = '';
  }

  btn.addEventListener('click', function () { dr.classList.contains('open') ? closeMenu() : openMenu(); });
  ov.addEventListener('click', closeMenu);
  cls.addEventListener('click', closeMenu);
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeMenu(); });

  // ── Active page highlight ──
  dr.querySelectorAll('.am-sb-nav a').forEach(function (a) {
    try {
      var href = new URL(a.getAttribute('href'), window.location.href);
      var curr = window.location.pathname.replace(/\/index\.html$/, '').replace(/\/+$/, '') || '/';
      var dest = href.pathname.replace(/\/index\.html$/, '').replace(/\/+$/, '') || '/';
      if (curr === dest) { a.classList.add('am-cur'); }
    } catch (e) {}
  });

  // ── Language toggle ──
  function applyLang(l) {
    lang = l;
    localStorage.setItem('am-lang', l);
    document.documentElement.setAttribute('data-lang', l);
    dr.querySelectorAll('.am-sb-lb').forEach(function (b) {
      b.classList.toggle('active', b.getAttribute('data-l') === l);
    });
  }
  dr.querySelectorAll('.am-sb-lb').forEach(function (b) {
    b.addEventListener('click', function () { applyLang(b.getAttribute('data-l')); });
  });
  applyLang(lang);

})();
