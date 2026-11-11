/* Referral finder: search and grouping rules, shared by the public page and the staff view.
   Loads one JSON file per Microsoft List from data/ (see data/README.md), joins them, and exposes window.FinderCore.
   FinderCore.ready is a promise that settles once the files are loaded. */
(function () {
  'use strict';
  var FILES = ['meta', 'areas', 'topics', 'topic_words', 'kinds', 'counties', 'organizations', 'listings', 'phones',
    'coverage', 'assignments', 'loj_areas', 'reasons', 'reason_areas', 'presets', 'pins', 'search_words', 'examples'];
  var DATA_DIR = (document.currentScript && document.currentScript.getAttribute('data-dir')) || 'data/';

  var D, AREA, TOPIC, KIND, LISTING, REASON;
  var FC = window.FinderCore = { loaded: false, error: '' };

  // ---------- Reading the List exports ----------
  // A file may be a plain array of rows, or a Power Automate / SharePoint response with the rows under "value".
  function rows(x) { return Array.isArray(x) ? x : (x && Array.isArray(x.value) ? x.value : []); }
  function str(v) { return v == null ? '' : String(v).trim(); }
  function yes(v) { return v === true || /^(yes|true|1)$/i.test(str(v)); }
  function num(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
  function byOrder(a, b) { return num(a.order) - num(b.order); }
  function group(list, key) { var m = {}; list.forEach(function (r) { var k = str(r[key]); (m[k] = m[k] || []).push(r); }); return m; }

  // Join the flat tables into the shape the rules below use.
  function join(T) {
    var meta = rows(T.meta)[0] || {};
    var words = group(rows(T.topic_words), 'topic_id');
    var examples = group(rows(T.examples).slice().sort(byOrder), 'topic_id');
    var phones = group(rows(T.phones), 'listing_id');
    var cover = group(rows(T.coverage), 'listing_id');
    var assign = group(rows(T.assignments), 'listing_id');
    var reasonAreas = group(rows(T.reason_areas), 'reason_id');
    var regions = {}, counties = [];
    rows(T.counties).forEach(function (r) {
      var c = str(r.county), g = str(r.region);
      if (!c) return;
      counties.push(c);
      if (g) (regions[g] = regions[g] || []).push(c);
    });
    counties.sort();
    var orgs = {};
    rows(T.organizations).forEach(function (r) { orgs[str(r.org_id)] = { name: str(r.name), website: str(r.website), desc: str(r.description) }; });
    var listings = rows(T.listings).map(function (r) {
      var id = str(r.listing_id), cov = cover[id] || [], as = assign[id] || [];
      var scopes = cov.map(function (c) { return str(c.scope).toLowerCase(); });
      var scope = scopes.indexOf('national') >= 0 ? 'national' : scopes.indexOf('statewide') >= 0 ? 'statewide'
        : scopes.indexOf('region') >= 0 ? 'region' : scopes.indexOf('county') >= 0 ? 'county' : '';
      function placesOf(s) { return cov.filter(function (c) { return str(c.scope).toLowerCase() === s; }).map(function (c) { return str(c.place); }).filter(Boolean); }
      function idsOf(t) { return as.filter(function (a) { return str(a.type).toLowerCase() === t; }).map(function (a) { return str(a.id); }); }
      return {
        id: id, org: str(r.org_id), role: str(r.role), program: str(r.program), office: str(r.office),
        website: str(r.website), email: str(r.email), intake: str(r.intake_link), desc: str(r.description), alert: str(r.alert),
        checked: str(r.last_checked).slice(0, 10), source: str(r.source),
        phone: (phones[id] || []).slice().sort(byOrder).map(function (p) {
          var digits = str(p.phone).replace(/[^0-9]/g, '');
          return { shown: str(p.shown_as) || str(p.phone), tel: digits.length === 10 ? '+1' + digits : digits, label: str(p.label) };
        }),
        scope: scope, counties: placesOf('county'), regions: placesOf('region'),
        places: scope === 'region' ? placesOf('region') : placesOf('county'),
        topics: idsOf('topic'), kinds: idsOf('kind'), areas: idsOf('area')
      };
    });
    return {
      built: str(meta.built).slice(0, 10), sample: yes(meta.sample), generalTopic: str(meta.general_topic_id),
      areas: rows(T.areas).slice().sort(byOrder).map(function (r) { return { id: str(r.area_id), name: str(r.name), hint: str(r.hint), start: yes(r.is_start) }; }),
      topics: rows(T.topics).slice().sort(byOrder).map(function (r) {
        var id = str(r.topic_id);
        return { id: id, area: str(r.area_id), name: str(r.name), mayHandle: yes(r.may_handle),
          // examples for the "NYDHR may handle this" box (LEGAL CHECK the wording before it goes live)
          examples: (examples[id] || []).map(function (e) { return { text: str(e.text), why: str(e.why) }; }), words: (words[id] || []).map(function (w) { return str(w.word); }).filter(Boolean) };
      }),
      kinds: rows(T.kinds).slice().sort(byOrder).map(function (r) { return { id: str(r.kind_id), name: str(r.name) }; }),
      counties: counties, regions: regions, orgs: orgs, listings: listings,
      lojAreas: rows(T.loj_areas).slice().sort(byOrder).map(function (r) { return str(r.area); }),
      reasons: rows(T.reasons).map(function (r) {
        var id = str(r.reason_id);
        // description, group and group_order: the line under each reason and its heading in the staff view's reason list
        return { id: id, name: str(r.name), desc: str(r.description), group: str(r.group), groupOrder: num(r.group_order),
          areas: (reasonAreas[id] || []).map(function (a) { return str(a.area); }) };
      }),
      // Shared search words: place words (with a bias or characteristic word, open their topic first), the bias and
      // characteristic words that set that off, and generic words that topic names don't match on
      searchWords: rows(T.search_words).map(function (r) { return { word: str(r.word), role: str(r.role), topic: str(r.topic_id) }; }),
      presets: rows(T.presets).map(function (r) { return { reason: str(r.reason_id), area: str(r.area), topic: str(r.topic_id) }; }),
      pins: rows(T.pins).map(function (r) { return { reason: str(r.reason_id), area: str(r.area), order: num(r.order), listing: str(r.listing_id), any: yes(r.any_county) }; })
    };
  }

  function byId(list) { var m = {}; list.forEach(function (x) { m[x.id] = x; }); return m; }
  function norm(s) {
    return String(s || '').toLowerCase().replace(/[’‘]/g, "'").replace(/[^a-z0-9'+ ]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  var STOP = {};   // filler words the search skips ("the", "my"): search_words.json, role "filler"

  function init(data) {
    D = data;
    AREA = byId(D.areas); TOPIC = byId(D.topics); KIND = byId(D.kinds); LISTING = byId(D.listings); REASON = byId(D.reasons);
    // Text each listing can be found by: its names, description, and its topics' names and hidden words.
    D.listings.forEach(function (l) {
      var org = D.orgs[l.org] || {};
      var bits = [org.name, l.program, l.office, l.desc, org.desc];
      l.topics.forEach(function (t) { var tp = TOPIC[t]; if (tp) { bits.push(tp.name); bits = bits.concat(tp.words); } });
      l._hay = ' ' + norm(bits.join(' ')) + ' ';
    });
    buildSearch();
    FC.data = D; FC.AREA = AREA; FC.TOPIC = TOPIC; FC.KIND = KIND; FC.LISTING = LISTING; FC.REASON = REASON;
  }
  init(join({}));

  FC.ready = Promise.all(FILES.map(function (f) {
    return fetch(DATA_DIR + f + '.json', { cache: 'no-cache' }).then(function (r) {
      if (!r.ok) throw new Error(f + '.json: ' + r.status);
      return r.json();
    });
  })).then(function (list) {
    var T = {};
    FILES.forEach(function (f, i) { T[f] = list[i]; });
    init(join(T));
    FC.loaded = true;
  }).catch(function (err) {
    FC.error = String(err && err.message || err);
    console.error('[finder] could not load the listings:', err);
    throw err;
  });

  // "Kings (Brooklyn)" -> "Kings County (Brooklyn)"; "Monroe" -> "Monroe County"
  function countyLabel(c) { var m = /^(.*?) \((.*)\)$/.exec(c || ''); return m ? m[1] + ' County (' + m[2] + ')' : c + ' County'; }
  // A county from a link: the full name ("Kings (Brooklyn)") or the short one ("Kings"), any capitals
  function countyFrom(v) {
    var n = norm(v), hit = '';
    if (!n) return '';
    D.counties.some(function (c) { if (norm(c) === n || norm(c.split(' (')[0]) === n) { hit = c; return true; } return false; });
    return hit;
  }
  function isWide(l) { return l.scope === 'statewide' || l.scope === 'national'; }
  function inCounty(l, county) {
    if (!county) return false;
    if (l.counties.indexOf(county) >= 0) return true;
    return l.regions.some(function (r) { return (D.regions[r] || []).indexOf(county) >= 0; });
  }
  function whereLine(l) {
    if (l.office) return l.office;
    if (l.scope === 'national') return 'Federal';
    if (l.scope === 'statewide') return 'All of New York';
    if (l.scope === 'region') return 'Serves the ' + l.places.join(' and ') + ' region';
    if (l.scope === 'county') return l.places.length === 1 ? countyLabel(l.places[0]) : 'Serves ' + l.places.join(', ') + ' counties';
    return '';
  }
  function websiteOf(l) { return l.website || (D.orgs[l.org] || {}).website || ''; }

  // ---------- Suggestions while typing ----------
  // A topic matches on its hidden words and on the words of its name. Rules (search audit, 2026-11-09 and 11-11):
  // - Whole words only: "fire" does not match "fired", "trans" does not match "transportation" (a plural s or es is fine).
  // - Generic words ("work", "insurance", "retaliation") in a topic name don't match; a topic matches them only if it has
  //   the word itself as a hidden word.
  // - A place word plus a bias or characteristic word ("my boss is racist") puts the place's topic first.
  // - A misspelt word ("preganant") is read as the nearest known word, one letter off (two for long words).
  var GENERIC = {}, PLACES = [], TRIGGERS = [], VOCAB = {}, KNOWN = {};

  function phraseRe(w) {   // whole words, allowing a plural s or es, or 's, on the last word ("workers" finds "workers'")
    return new RegExp('(^| )' + w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "(e?s|'s?)?(?= |$)");
  }
  function buildSearch() {
    GENERIC = {}; PLACES = []; TRIGGERS = []; VOCAB = {}; KNOWN = {}; STOP = {};
    D.searchWords.forEach(function (w) {
      var n = norm(w.word);
      if (!n) return;
      if (w.role === 'filler') STOP[n] = 1;
      else if (w.role === 'generic') GENERIC[n] = 1;
      else if (w.role === 'place') { if (TOPIC[w.topic]) PLACES.push({ n: n, re: phraseRe(n), topic: w.topic, word: w.word }); }
      else TRIGGERS.push({ n: n, re: phraseRe(n), word: w.word });
    });
    D.topics.forEach(function (t) {
      t._name = norm(t.name);
      t._words = t.words.map(function (w) { var n = norm(w); return { w: w, n: n, re: phraseRe(n) }; }).filter(function (x) { return x.n; });
      // the words of the name that count: not generic, not filler
      // (short words like "at", "day" or "off" are left out: they match too much inside a sentence)
      t._nameWords = t._name.split(' ').filter(function (x) { return x.length >= 4 && !STOP[x] && !GENERIC[x]; });
      (t._name + ' ' + t._words.map(function (x) { return x.n; }).join(' ')).split(' ').forEach(function (x) { if (x.length >= 4) VOCAB[x] = 1; });
    });
    PLACES.concat(TRIGGERS).forEach(function (x) { x.n.split(' ').forEach(function (y) { if (y.length >= 4) VOCAB[y] = 1; }); });
    Object.keys(VOCAB).forEach(function (x) { KNOWN[x] = 1; });
    D.listings.forEach(function (l) { l._hay.split(' ').forEach(function (x) { if (x) KNOWN[x] = 1; }); });
  }

  // Edits needed to turn a into b (letters added, dropped, changed or swapped), stopping early once over max
  function editDistance(a, b, max) {
    if (Math.abs(a.length - b.length) > max) return max + 1;
    var prev2 = [], prev = [], cur, i, j;
    for (j = 0; j <= b.length; j++) prev[j] = j;
    for (i = 1; i <= a.length; i++) {
      cur = [i];
      var low = i;
      for (j = 1; j <= b.length; j++) {
        var cost = a[i - 1] === b[j - 1] ? 0 : 1;
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
        if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) cur[j] = Math.min(cur[j], prev2[j - 2] + 1);
        if (cur[j] < low) low = cur[j];
      }
      if (low > max) return max + 1;
      prev2 = prev; prev = cur;
    }
    return prev[b.length];
  }
  // The query with misspelt words swapped for the nearest search word. Words of 5 or more letters only.
  function fixTypos(qn) {
    return qn.split(' ').map(function (x) {
      if (x.length < 5 || KNOWN[x] || STOP[x]) return x;
      var max = x.length >= 9 ? 2 : 1, best = '', bestD = max + 1;
      Object.keys(VOCAB).forEach(function (v) {
        var d = editDistance(x, v, max);
        if (d < bestD) { bestD = d; best = v; }
      });
      return best || x;
    }).join(' ');
  }

  // Topics a place word plus a bias or characteristic word points to, in the order the place words were typed
  function intentTopics(qn) {
    var trigger = null;
    TRIGGERS.some(function (x) { if (x.re.test(qn)) { trigger = x; return true; } return false; });
    if (!trigger) return [];
    var hits = [];
    PLACES.forEach(function (p) {
      var m = p.re.exec(qn);
      if (m) hits.push({ at: m.index, topic: p.topic, word: p.word + '” and “' + trigger.word });
    });
    hits.sort(function (a, b) { return a.at - b.at; });
    var seen = {}, out = [];
    hits.forEach(function (h) { if (!seen[h.topic]) { seen[h.topic] = 1; out.push(h); } });
    return out;
  }
  function intentTopic(q) {
    var qn = fixTypos(norm(q));
    return intentTopics(qn).map(function (h) { return TOPIC[h.topic]; });
  }

  // Does typed text (possibly an unfinished word) start a word in this text, and is that more than generic words?
  function startsWord(text, qn) {
    var at = (' ' + text).indexOf(' ' + qn);
    if (at < 0) return false;
    var typed = qn.split(' '), covered = text.slice(at).split(' ').slice(0, typed.length);
    // a generic word typed in full ("work") is not the start of a longer word ("worker")
    if (GENERIC[typed[typed.length - 1]] && covered[covered.length - 1] !== typed[typed.length - 1]) return false;
    return covered.some(function (x) { return !STOP[x] && !GENERIC[x]; }) || text === qn;
  }

  function scoreTopics(qn) {
    var out = [], content = qn.split(' ').filter(function (x) { return x && !STOP[x]; });
    var intent = intentTopics(qn), rank = {};
    intent.forEach(function (h, i) { rank[h.topic] = i; });
    D.topics.forEach(function (t) {
      var score = -1, via = '', hits = 0;
      if (t.id in rank) { score = 0; via = intent[rank[t.id]].word; }
      else if (t._name.indexOf(qn) === 0 && startsWord(t._name, qn)) score = 1;
      else if (startsWord(t._name, qn)) score = 2;
      // hidden words typed (or being typed), and words of the name inside a longer search
      t._words.forEach(function (x) {
        var whole = x.n === qn || x.re.test(qn);
        if (whole || startsWord(x.n, qn)) {
          // the whole search beats a phrase in it, a longer phrase beats a shorter one, and a whole word beats an unfinished one
          hits += x.n === qn ? 20 : whole ? 1 + x.n.split(' ').length : 1;
          if (!via) via = x.w;
        }
      });
      // words of the name inside a longer search only break ties ("landlord" puts Landlord harassment first),
      // or find something when nothing else matched
      var named = t._nameWords.filter(function (y) { return content.indexOf(y) >= 0 || content.indexOf(y + 's') >= 0 || content.indexOf(y.replace(/s$/, '')) >= 0; }).length;
      if (score < 0 && hits > 0) score = 3;
      if (score < 0 && named > 0) score = 4;
      hits += named;
      if (score >= 0) out.push({ topic: t, score: score, hits: hits, via: score === 1 || score === 2 ? '' : via, order: D.topics.indexOf(t) });
    });
    if (out.some(function (x) { return x.score < 4; })) out = out.filter(function (x) { return x.score < 4; });
    out.sort(function (a, b) { return a.score - b.score || b.hits - a.hits || a.order - b.order; });
    return out;
  }
  function suggest(q, max) {
    var qn = norm(q);
    if (qn.length < 2) return [];
    var out = scoreTopics(qn);
    if (!out.length) {
      var fixed = fixTypos(qn);
      if (fixed !== qn) out = scoreTopics(fixed);
    }
    return out.slice(0, max || 5).map(function (s) {
      var area = AREA[s.topic.area];
      return { id: s.topic.id, name: s.topic.name, sub: (area ? area.name : '') + (s.via ? ' · matches “' + s.via + '”' : '') };
    });
  }
  // Organizations whose name has a word starting with what was typed ("camba", "legal aid", "NAACP")
  function suggestOrgs(q, max) {
    var qn = norm(q);
    if (qn.length < 2) return [];
    var byOrg = {};
    D.listings.forEach(function (l) { (byOrg[l.org] = byOrg[l.org] || []).push(l); });
    var out = [];
    Object.keys(byOrg).forEach(function (id) {
      var o = D.orgs[id];
      if (!o) return;
      var nn = norm(o.name), at = (' ' + nn).indexOf(' ' + qn);
      if (at < 0) return;
      out.push({ id: id, name: o.name, score: at === 0 ? 0 : 1, sub: orgWhere(byOrg[id]) });
    });
    out.sort(function (a, b) { return a.score - b.score || a.name.localeCompare(b.name); });
    return out.slice(0, max || 3).map(function (o) { return { id: o.id, name: o.name, sub: o.sub }; });
  }
  // Where an organization helps, in a few words: "All of New York", or its first counties and regions
  function orgWhere(ls) {
    if (ls.some(function (l) { return l.scope === 'statewide'; })) return 'All of New York';
    if (ls.some(function (l) { return l.scope === 'national'; })) return 'Federal';
    var places = [];
    ls.forEach(function (l) { l.places.forEach(function (p) { if (places.indexOf(p) < 0) places.push(p); }); });
    if (!places.length) return '';
    return places.slice(0, 2).join(', ') + (places.length > 2 ? ' and ' + (places.length - 2) + ' more' : '');
  }
  function exactOrg(q) {
    var qn = norm(q), hit = null;
    if (!qn) return null;
    Object.keys(D.orgs).some(function (id) { if (norm(D.orgs[id].name) === qn) { hit = id; return true; } return false; });
    return hit && D.listings.some(function (l) { return l.org === hit; }) ? hit : null;
  }
  // A typed phrase that is exactly a topic's name or one of its words opens that topic.
  function exactTopic(q) {
    var qn = norm(q);
    if (!qn) return null;
    var hit = null;
    D.topics.some(function (t) { if (norm(t.name) === qn) { hit = t; return true; } return false; });
    return hit;
  }

  // ---------- Which listings match ----------
  function matches(query) {
    if (!query) return [];
    if (query.type === 'topic') return D.listings.filter(function (l) { return l.topics.indexOf(query.value) >= 0; });
    if (query.type === 'org') return D.listings.filter(function (l) { return l.org === query.value; });
    if (query.type === 'area') return D.listings.filter(function (l) { return l.topics.some(function (t) { return TOPIC[t] && TOPIC[t].area === query.value; }); });
    var toks = norm(query.value).split(' ').filter(function (t) { return t.length > 1 && !STOP[t]; });
    if (!toks.length) return [];
    // Same rules as the suggestions: whole words (a plural s or es is fine), and misspelt words read as the nearest search word
    var find = function (ts) {
      var res = ts.map(phraseRe);
      return D.listings.filter(function (l) { return res.every(function (re) { return re.test(l._hay); }); });
    };
    var found = find(toks);
    if (found.length) return found;
    var fixed = fixTypos(toks.join(' ')).split(' ');
    return fixed.join(' ') === toks.join(' ') ? found : find(fixed);
  }
  function queryLabel(query) {
    if (!query) return '';
    if (query.type === 'topic') return (TOPIC[query.value] || {}).name || query.value;
    if (query.type === 'area') return (AREA[query.value] || {}).name || query.value;
    if (query.type === 'org') return (D.orgs[query.value] || {}).name || query.value;
    return '“' + query.value + '”';
  }
  function queryAreaId(query) {
    if (!query) return '';
    if (query.type === 'topic') return (TOPIC[query.value] || {}).area || '';
    if (query.type === 'area') return query.value;
    return '';
  }

  // ---------- Cards: one per organization ----------
  function tagsFor(list, query) {
    var first = [], rest = [], seen = {};
    list.forEach(function (l) {
      l.topics.forEach(function (t) {
        if (seen[t] || !TOPIC[t]) return; seen[t] = 1;
        (query && query.type === 'topic' && query.value === t ? first : rest).push(TOPIC[t].name);
      });
    });
    var kinds = [];
    list.forEach(function (l) { l.kinds.forEach(function (k) { if (!seen[k] && KIND[k]) { seen[k] = 1; kinds.push(KIND[k].name); } }); });
    return first.concat(rest).slice(0, 4).concat(kinds).map(function (n) { return { name: n }; });
  }
  function phoneOf(l) {
    var p = l.phone[0];
    return p ? { shown: p.shown, href: 'tel:' + (p.tel || p.shown.replace(/[^0-9+]/g, '')), label: p.label || '' } : null;
  }
  function unique(arr) { var s = {}; return arr.filter(function (x) { if (s[x]) return false; s[x] = 1; return true; }); }

  function buildCards(list, query, county) {
    var byOrg = {}, order = [];
    list.forEach(function (l) { if (!byOrg[l.org]) { byOrg[l.org] = []; order.push(l.org); } byOrg[l.org].push(l); });
    return order.map(function (orgId) {
      var ls = byOrg[orgId], org = D.orgs[orgId] || { name: orgId };
      var mains = ls.filter(function (l) { return !l.program; });
      var progs = {}, progOrder = [];
      ls.filter(function (l) { return l.program; }).forEach(function (l) { if (!progs[l.program]) { progs[l.program] = []; progOrder.push(l.program); } progs[l.program].push(l); });
      var placeName = county ? countyLabel(county) : 'New York';
      var c = {
        key: orgId, name: org.name, url: org.website || websiteOf(ls[0]), listings: ls,
        kinds: unique([].concat.apply([], ls.map(function (l) { return l.kinds; }))),
        simple: false, grouped: false, sites: false, hasPrograms: progOrder.length > 0,
        phone: null, where: '', desc: '', alert: '', noPhone: false, tags: [], offices: [], programs: []
      };
      if (mains.length === 1 && !progOrder.length) {
        var m = mains[0];
        c.simple = true; c.url = websiteOf(m); c.phone = phoneOf(m); c.where = whereLine(m);
        c.desc = m.desc || org.desc || ''; c.alert = m.alert || ''; c.noPhone = !c.phone; c.tags = tagsFor([m], query);
      } else {
        c.desc = org.desc || (mains[0] && mains[0].desc) || '';
        if (mains.length === 1) {
          c.phone = phoneOf(mains[0]); c.where = whereLine(mains[0]); c.noPhone = !c.phone; c.alert = mains[0].alert || '';
          c.tags = tagsFor(mains, query);
        } else if (mains.length > 1) {
          var ownSites = mains.some(function (l) { return l.website && l.website !== org.website; });
          c.sites = ownSites; c.grouped = !ownSites;
          c.officeHead = mains.length + ' ' + (ownSites ? 'branches' : 'offices') + ' in ' + placeName;
          c.offices = mains.map(function (l) {
            var p = phoneOf(l);
            return { key: l.id, name: l.office || whereLine(l), url: websiteOf(l), phone: p, hasPhone: !!p, where: whereLine(l), tags: tagsFor([l], query) };
          });
          if (!ownSites) c.tags = tagsFor(mains, query);
        }
        c.programs = progOrder.map(function (name) {
          var pl = progs[name], one = pl.length === 1, p = phoneOf(pl[0]);
          return {
            key: orgId + name, name: name, url: websiteOf(pl[0]), one: one, many: !one,
            phone: one ? p : null, hasPhone: one && !!p, noPhone: one && !p, where: one ? whereLine(pl[0]) : '',
            desc: pl[0].desc || '', alert: pl[0].alert || '', tags: tagsFor(pl, query),
            siteHead: pl.length + ' sites in ' + placeName,
            sites: pl.map(function (l) { var sp = phoneOf(l); return { key: l.id, phone: sp, hasPhone: !!sp, where: whereLine(l) }; })
          };
        });
        c.programHead = progOrder.length + (progOrder.length === 1 ? ' program' : ' programs');
      }
      c.hasPhone = !!c.phone; c.hasDesc = !!c.desc; c.hasAlert = !!c.alert; c.hasTags = c.tags.length > 0;
      c.hasWhere = !!c.where;
      return c;
    });
  }

  // ---------- Picks ("Often helps with this") ----------
  function pinsFor(reason, lojArea, county) {
    if (!reason) return [];
    var rows = D.pins.filter(function (p) { return p.reason === reason && (p.area === 'All' || !lojArea || p.area === lojArea); });
    rows.sort(function (a, b) { return a.order - b.order; });
    var out = [], orgs = {};
    rows.forEach(function (p) {
      var l = LISTING[p.listing];
      if (!l || orgs[l.org] || out.length >= 3) return;
      var ok = county ? (p.any || isWide(l) || inCounty(l, county)) : isWide(l);
      if (ok) { orgs[l.org] = 1; out.push(l); }
    });
    return out;
  }
  function presetFor(reason, lojArea) {
    var hit = null;
    D.presets.some(function (p) { if (p.reason === reason && p.area === lojArea) { hit = p; return true; } return false; });
    if (!hit) D.presets.some(function (p) { if (p.reason === reason && p.area === 'All') { hit = p; return true; } return false; });
    if (!hit) D.presets.some(function (p) { if (p.reason === reason) { hit = p; return true; } return false; });
    return hit ? hit.topic : '';
  }
  function reasonsFor(lojArea) {
    return D.reasons.filter(function (r) { return !lojArea || r.areas.indexOf('All') >= 0 || r.areas.indexOf(lojArea) >= 0; });
  }

  // ---------- The whole results page ----------
  // opts: {query, county, kinds: {kindId: true}, reason, lojArea}
  function results(opts) {
    var query = opts.query, county = opts.county || '', checked = opts.kinds || {};
    var all = matches(query);
    var pinned = pinsFor(opts.reason, opts.lojArea, county);
    var pinnedOrgs = {}; pinned.forEach(function (l) { pinnedOrgs[l.org] = 1; });
    var rest = all.filter(function (l) { return !pinnedOrgs[l.org]; });
    var local = county ? rest.filter(function (l) { return inCounty(l, county); }) : [];
    var wide = rest.filter(isWide);
    // An organization search shows all of that organization, so its offices elsewhere get their own group
    var elsewhere = query && query.type === 'org' ? rest.filter(function (l) { return !isWide(l) && !inCounty(l, county); }) : [];

    // Broader matches: listings that help with the whole area in general (listing.areas), for a topic search
    var broad = [];
    var areaId = queryAreaId(query);
    if (query && query.type === 'topic' && areaId) {
      var shown = {}; all.forEach(function (l) { shown[l.id] = 1; });
      broad = D.listings.filter(function (l) {
        return !shown[l.id] && !pinnedOrgs[l.org] && (l.areas || []).indexOf(areaId) >= 0 && (isWide(l) || inCounty(l, county));
      });
    }

    var cardsLocal = buildCards(local, query, county), cardsWide = buildCards(wide, query, county), cardsBroad = buildCards(broad, query, county);
    var cardsPinned = buildCards(pinned, query, county), cardsElse = buildCards(elsewhere, query, '');
    var localCount = cardsLocal.length;

    // Kinds of help, with counts across the three groups (not the picks)
    var canNarrow = !!query && query.type !== 'text' && query.type !== 'org';
    var pool = cardsLocal.concat(cardsWide, cardsBroad);
    var kindList = D.kinds.map(function (k) {
      var n = pool.filter(function (c) { return c.kinds.indexOf(k.id) >= 0; }).length;
      return { id: k.id, name: k.name, count: n, checked: !!checked[k.id] };
    }).filter(function (k) { return k.count > 0 || k.checked; });
    var anyChecked = canNarrow && kindList.some(function (k) { return k.checked; });
    function keep(c) { return !anyChecked || c.kinds.some(function (k) { return checked[k]; }); }
    var before = pool.length;
    cardsLocal = cardsLocal.filter(keep); cardsWide = cardsWide.filter(keep); cardsBroad = cardsBroad.filter(keep);
    var after = cardsLocal.length + cardsWide.length + cardsBroad.length + cardsElse.length;

    var groups = [];
    if (cardsPinned.length) groups.push({ key: 'g0', cls: 'grp g0', title: 'Often helps with this', cards: cardsPinned });
    if (county && cardsLocal.length) groups.push({ key: 'g1', cls: 'grp g1', title: 'In ' + countyLabel(county), cards: cardsLocal });
    if (cardsWide.length) groups.push({ key: 'g2', cls: 'grp g2', title: 'Statewide and federal help', cards: cardsWide });
    if (cardsBroad.length) {
      var an = (AREA[areaId] || {}).name || '';
      groups.push({ key: 'g3', cls: 'grp g3', title: 'Broader matches', cards: cardsBroad,
        sub: 'These help with ' + an.toLowerCase() + ' in general, not “' + queryLabel(query) + '” in particular.' });
    }
    if (cardsElse.length) groups.push({ key: 'g4', cls: 'grp g3', title: county ? 'Other places it helps' : 'Where it helps', cards: cardsElse });
    groups.forEach(function (g) { g.count = g.cards.length + (g.cards.length === 1 ? ' result' : ' results'); g.hasSub = !!g.sub; });

    var total = cardsPinned.length + after;
    // Warn only when nothing local turned up at all (a local pick counts as local help)
    var note = '';
    var localPicks = pinned.filter(function (l) { return inCounty(l, county); }).length;
    if (query && query.type !== 'org' && county && total > 0 && localCount === 0 && localPicks === 0) {
      note = 'There are no local organizations listed in ' + countyLabel(county) + ' for this search'
        + (cardsWide.length ? ', but we’ve included statewide and federal help instead.' : '.');
    }
    var kindNames = kindList.filter(function (k) { return k.checked; }).map(function (k) { return k.name.toLowerCase(); });
    return {
      groups: groups, total: total, note: note, kinds: canNarrow ? kindList : [],
      hidden: anyChecked ? before - after : 0, before: before + cardsPinned.length,
      narrowedTo: kindNames.join(' or '),
      pinnedListings: pinned, restListings: rest.filter(function (l) { return !county || inCounty(l, county) || isWide(l); })
    };
  }
  function labelQuoted(query) { return query.type === 'text' ? queryLabel(query) : '“' + queryLabel(query) + '”'; }

  // Areas that share a word with a search that found nothing
  function areasLike(q) {
    var toks = norm(q).split(' ').filter(function (t) { return t.length > 2 && !STOP[t]; });
    var scored = D.areas.filter(function (a) { return !a.start; }).map(function (a) {
      var hay = norm([a.name, a.hint].concat(D.topics.filter(function (t) { return t.area === a.id; }).map(function (t) { return t.name + ' ' + (t.words || []).join(' '); })).join(' '));
      var n = toks.filter(function (t) { return hay.indexOf(t.slice(0, Math.max(4, t.length - 2))) >= 0; }).length;
      return { a: a, n: n };
    }).filter(function (x) { return x.n > 0; });
    scored.sort(function (x, y) { return y.n - x.n; });
    return scored.slice(0, 3).map(function (x) { return x.a; });
  }

  var api = {
    norm: norm, suggest: suggest, intentTopic: intentTopic, suggestOrgs: suggestOrgs, exactOrg: exactOrg, exactTopic: exactTopic, matches: matches, results: results, buildCards: buildCards,
    queryLabel: queryLabel, queryAreaId: queryAreaId, whereLine: whereLine, websiteOf: websiteOf, phoneOf: phoneOf,
    isWide: isWide, countyLabel: countyLabel, countyFrom: countyFrom, inCounty: inCounty, pinsFor: pinsFor, presetFor: presetFor, reasonsFor: reasonsFor, areasLike: areasLike
  };
  Object.keys(api).forEach(function (k) { FC[k] = api[k]; });
})();
