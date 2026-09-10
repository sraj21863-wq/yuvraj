/* Schema check for the dashboard payload.

   The pipeline runs this same check on the way out and the page runs it on the
   way in, so a payload that is malformed, truncated or from a different build
   of the pipeline is rejected at both ends rather than half-rendered. The
   dashboard shows nothing at all rather than a figure it cannot stand behind.

   validate()     -> array of hard errors. Non-empty means: do not render.
   softWarnings() -> array of things a reader should know about data that is
                     good enough to render (unmatched sales people, orders with
                     no owner, status flags that are neither open nor invoiced). */
window.OB_VALIDATE = (function () {
  'use strict';

  var GRID_KEY = /^(\d{1,2})\|([^|]+)\|([^|]+)$/;
  var STATUS_KEY = /^(\d{1,2})\|([^|]+)\|([^|]+)\|(.*)$/;
  var PERSON_KEY = /^(\d{1,2})\|([^|]+)\|([^|]+)\|(.+)$/;

  function isObj(v) { return v && typeof v === 'object' && !Array.isArray(v); }

  function strArray(p, name, errs, min) {
    var v = p[name];
    if (!Array.isArray(v)) { errs.push(name + ': expected an array'); return null; }
    if (min && v.length < min) {
      errs.push(name + ': expected at least ' + min + ' entries, got ' + v.length);
      return null;
    }
    for (var i = 0; i < v.length; i++) {
      if (typeof v[i] !== 'string' || !v[i]) {
        errs.push(name + '[' + i + ']: expected a non-empty string');
        return null;
      }
    }
    return v;
  }

  function numberMap(p, name, errs, keyRe, required) {
    var v = p[name];
    if (v === undefined || v === null) {
      if (required) errs.push(name + ': missing');
      return null;
    }
    if (!isObj(v)) { errs.push(name + ': expected an object'); return null; }
    var k, n = 0, bad = 0;
    for (k in v) {
      if (!Object.prototype.hasOwnProperty.call(v, k)) continue;
      n++;
      if (typeof v[k] !== 'number' || !isFinite(v[k])) {
        if (bad++ < 3) errs.push(name + '["' + k + '"]: ' + v[k] + ' is not a finite number');
        continue;
      }
      if (keyRe && !keyRe.test(k)) {
        if (bad++ < 3) errs.push(name + ': key "' + k + '" does not fit ' + keyRe);
      }
    }
    return n;
  }

  function validate(p) {
    var errs = [];
    if (!isObj(p)) return ['payload is not an object'];

    if (p.schema && String(p.schema).split('/')[0] !== 'ob-dashboard') {
      errs.push('schema: "' + p.schema + '" is not an order-booking payload');
    }
    ['generatedAt', 'sourceFile', 'sourceModified'].forEach(function (k) {
      if (typeof p[k] !== 'string' || !p[k]) errs.push(k + ': missing');
    });
    if (p.generatedAt && isNaN(new Date(p.generatedAt))) {
      errs.push('generatedAt: "' + p.generatedAt + '" is not a date');
    }

    var months = strArray(p, 'months', errs, 12);
    var regions = strArray(p, 'regions', errs, 1);
    var segments = strArray(p, 'segments', errs, 1);

    numberMap(p, 'actual', errs, GRID_KEY, true);
    numberMap(p, 'target', errs, GRID_KEY, true);
    numberMap(p, 'count', errs, GRID_KEY, true);
    numberMap(p, 'status', errs, STATUS_KEY, true);
    numberMap(p, 'soip', errs, GRID_KEY, false);
    numberMap(p, 'invoicedAmount', errs, GRID_KEY, false);
    numberMap(p, 'pactual', errs, PERSON_KEY, true);
    numberMap(p, 'pcount', errs, PERSON_KEY, true);
    numberMap(p, 'ptarget', errs, null, true);
    numberMap(p, 'excludedByMonth', errs, GRID_KEY, false);

    if (!Array.isArray(p.people)) errs.push('people: expected an array');
    if (!Array.isArray(p.openStatuses) || !p.openStatuses.length) {
      errs.push('openStatuses: expected a non-empty array');
    }
    if (!Array.isArray(p.invoicedStatuses) || !p.invoicedStatuses.length) {
      errs.push('invoicedStatuses: expected a non-empty array');
    }

    var lm = p.latestMonth;
    if (typeof lm !== 'number' || lm < 1 || lm > 12 || lm !== Math.floor(lm)) {
      errs.push('latestMonth: expected a whole number 1-12, got ' + lm);
    }

    /* A payload can be well formed and still be useless: a target of zero makes
       every completion percentage on the page infinite. Catch it here rather
       than letting the page print it. */
    if (isObj(p.target)) {
      var t = 0, k;
      for (k in p.target) if (typeof p.target[k] === 'number') t += p.target[k];
      if (!(t > 0)) errs.push('target: the whole year totals ' + t + ' -- nothing to measure against');
    }

    /* Keys must name a region and segment the page can actually show, or the
       figure silently lands nowhere. */
    if (regions && segments && isObj(p.actual)) {
      var known = {}, i, j, miss = 0;
      for (i = 0; i < regions.length; i++)
        for (j = 0; j < segments.length; j++) known[regions[i] + '|' + segments[j]] = 1;
      for (var key in p.actual) {
        var m = GRID_KEY.exec(key);
        if (m && !known[m[2] + '|' + m[3]] && miss++ < 3) {
          errs.push('actual: "' + key + '" is outside the region x segment grid');
        }
      }
    }
    if (months && months.length !== 12) errs.push('months: expected 12 names');
    return errs;
  }

  function softWarnings(p) {
    var out = [];
    if (!isObj(p)) return out;
    if (Array.isArray(p.warnings)) {
      p.warnings.forEach(function (w) { out.push(String(w)); });
    }
    var other = 0, k;
    if (isObj(p.status)) {
      var open = {}, inv = {};
      (p.openStatuses || []).forEach(function (s) { open[s] = 1; });
      (p.invoicedStatuses || []).forEach(function (s) { inv[s] = 1; });
      for (k in p.status) {
        var flag = k.split('|').slice(3).join('|');
        if (!open[flag] && !inv[flag]) other += p.status[k];
      }
    }
    if (Math.abs(other) > 0.005) {
      out.push(other.toFixed(2) + ' Cr carries a status that is neither open nor invoiced');
    }
    if (p.excludedTotal) {
      out.push(Number(p.excludedTotal).toFixed(2) + ' Cr sits outside the reporting grid');
    }
    if (isObj(p.fuzzyMatchedPeople)) {
      var n = Object.keys(p.fuzzyMatchedPeople).length;
      if (n) out.push(n + ' sales ' + (n === 1 ? 'name was' : 'names were') +
                      ' matched to a target row by word overlap');
    }
    return out;
  }

  return { validate: validate, softWarnings: softWarnings };
})();
