/**
 * metric_engine.js — the dashboard twin of scripts/metric_engine.py.
 *
 * Implements references/data-contract.md §5 (metric DSL) literally and mirrors every
 * implementation choice of scripts/metric_engine.py (see IMPLEMENTATION_NOTES there), so
 * that for every metric and every filter combination the dashboard reproduces the
 * authoritative Python results in 05_metrics.json: same `value`, `status`, `n_included`
 * (checked by the parity cases through scripts/check_dashboard.mjs), and the same
 * included_ids / excluded / auxiliary numbers.
 *
 * Mirroring rules that matter for parity:
 *  - records are ordered by id ascending using Python's sorted() order (Unicode code
 *    points; identical to plain `a < b` for every id without astral-plane characters —
 *    never localeCompare), and every sum is accumulated in that order;
 *  - check order: period outside spec.valid_periods -> not_comparable; unresolved
 *    placeholder -> not_computable; growth not_comparable; |R| = 0 -> empty (with
 *    empty_cause, and empty_as_zero for sum/count); then the kind rules;
 *  - eq/in/has use strict equality (no bool/number coercion; lists/objects never equal);
 *  - a field missing from a record has status null (known() false, reason "absent");
 *  - a not_computable result with no included or excluded record lists its scope in scope_ids;
 *  - Python truthiness is reproduced where the Python code relies on `or` / `if x`.
 *
 * Dependency-free ES2017; no eval, no network, no DOM access.
 * Browser: defines window.createEngine (and window.APAEngine with helpers).
 * Node: module.exports = {createEngine, ...}; without CommonJS it sets globalThis.createEngine.
 *
 * API:  const eng = createEngine(data)   // data = analysis.json ({records, config, metrics, ...})
 *   eng.compute(spec, filters)          -> metric result (data-contract §5 result object)
 *   eng.computeBreakdown(spec, filters) -> breakdown result
 *   eng.computeTrend(filters)           -> trend result
 *   eng.computeAll(filters)             -> {filters, metrics:{id:res}, breakdowns:{id:res}, trend, records:[ids], record_ids:[ids]}
 *   eng.filterRecords(filters)          -> records passing base_where AND dashboard filters (id order)
 *   eng.resolveField(name, filters)     -> resolved field name for the filters' period, or null
 *   eng.normalizeFilters(filters)       -> {period, dim: value, ...} (throws on an unknown period)
 * filters = {"period": "2025", "asset_type": null | "...", "location": null | "...", "use": null | "..."}
 */
(function () {
  'use strict';

  var KNOWN_STATUSES = ['value', 'derived', 'zero_recorded', 'zero_confirmed'];
  var TEXT_KEYS = ['label', 'short_label', 'definition', 'scope_text', 'basis', 'valuation_date',
    'period_text', 'source_text', 'unavailable_reason', 'empty_text'];
  var PLACEHOLDERS = ['{REV}', '{REV_PREV}', '{EXP}', '{VAL}'];
  var NO_EXPENSE_REASON = 'لا توجد بيانات مصروفات';
  var DEFAULT_UNAVAILABLE = 'الحقل غير متاح لهذه الفترة';
  var INVALID_PERIOD_REASON = 'لا يُحسب هذا المؤشر للفترة المختارة (خارج الفترات الصالحة له)';
  var EMPTY_FILTERS_REASON = 'لا تطابق التصفية الحالية أي سجل في نطاق هذا المؤشر';
  var EMPTY_SCOPE_REASON = 'لا توجد سجلات مؤهلة لهذا المؤشر في الملف';
  var REVENUE_START_ONLY_REASON = 'كل السجلات المعروف إيرادها في الفترتين بدأ إيرادها في الفترة الحالية (إيراد سابق صفر)؛ ' +
    'يُعرض «بدء إيراد» والتغير بالريال دون نسبة نمو';
  var PREV_NONPOSITIVE_REASON = 'مجموع إيراد الفترة السابقة للسجلات المعروفة في الفترتين صفر أو سالب؛ ' +
    'لا تُحسب نسبة نمو';
  var NO_COVERAGE_FIELD_REASON = 'الحقول المطلوبة «لا تنطبق» على أي سجل ضمن النطاق';
  var NONE_KEY = '__none__';
  var OVERLAP_NOTE = 'قد تتداخل الفئات';
  var hasOwn = Object.prototype.hasOwnProperty;

  // -------------------------------------------------------------------------
  // Python-compatible primitives
  // -------------------------------------------------------------------------
  function isObj(x) { return x !== null && typeof x === 'object' && !Array.isArray(x); }

  function own(obj, key) {
    // dict.get(key) -> value or null (never inherited properties)
    if (obj === null || obj === undefined || typeof obj !== 'object') return null;
    if (!hasOwn.call(obj, key)) return null;
    var v = obj[key];
    return v === undefined ? null : v;
  }

  function nn(x) { return x === undefined ? null : x; }

  /** Python truthiness (None, False, 0, '', [], {} are falsy). */
  function truthy(x) {
    if (x === null || x === undefined || x === false) return false;
    if (typeof x === 'number') return x !== 0 && !isNaN(x);
    if (typeof x === 'string') return x.length > 0;
    if (Array.isArray(x)) return x.length > 0;
    if (typeof x === 'object') return Object.keys(x).length > 0;
    return true;
  }

  /** Python `a or b`. */
  function or(a, b) { return truthy(a) ? a : b; }

  function isNum(x) { return typeof x === 'number'; }

  /** Python str() for the scalar keys used by breakdowns. */
  function keyStr(k) {
    if (typeof k === 'boolean') return k ? 'true' : 'false';
    if (k === null || k === undefined) return 'None';
    return String(k);
  }

  /** Python string ordering (by Unicode code point). Returns -1, 0, 1. */
  function cmpStr(a, b) {
    a = String(a); b = String(b);
    if (a === b) return 0;
    var n = Math.min(a.length, b.length);
    for (var i = 0; i < n; i++) {
      var ca = a.charCodeAt(i), cb = b.charCodeAt(i);
      if (ca !== cb) {
        // surrogate halves (0xD800-0xDFFF) encode code points above 0xFFFF: they sort after BMP chars
        var sa = ca >= 0xD800 && ca <= 0xDFFF, sb = cb >= 0xD800 && cb <= 0xDFFF;
        if (sa !== sb) return sa ? 1 : -1;
        return ca < cb ? -1 : 1;
      }
    }
    return a.length < b.length ? -1 : 1;
  }

  function byId(a, b) { return cmpStr(a.id, b.id); }

  function sortedStrings(arr) { return arr.slice().sort(cmpStr); }

  /** Equality without coercion (mirrors metric_engine.strict_eq). */
  function strictEq(a, b) {
    a = nn(a); b = nn(b);
    if ((a !== null && typeof a === 'object') || (b !== null && typeof b === 'object')) return false;
    if (typeof a === 'boolean' || typeof b === 'boolean') {
      return typeof a === 'boolean' && typeof b === 'boolean' && a === b;
    }
    if (isNum(a) || isNum(b)) return isNum(a) && isNum(b) && a === b;
    if (a === null || b === null) return a === null && b === null;
    return typeof a === typeof b && a === b;
  }

  function inList(st, list) {
    // Python `st in (list or [])`
    if (!truthy(list)) return false;
    if (typeof list === 'string') return typeof st === 'string' && list.indexOf(st) >= 0;
    if (Array.isArray(list)) {
      for (var i = 0; i < list.length; i++) if (list[i] === st) return true;
      return false;
    }
    if (isObj(list)) return typeof st === 'string' && hasOwn.call(list, st);
    return false;
  }

  // -------------------------------------------------------------------------
  // Record access
  // -------------------------------------------------------------------------
  function getAttr(rec, attr) {
    if (attr === 'id' || attr === 'record_type' || attr === 'scope') return own(rec, attr);
    var dims = or(own(rec, 'dims'), {});
    return own(dims, attr);
  }

  function getTv(rec, field) {
    var fields = or(own(rec, 'fields'), {});
    var tv = own(fields, field);
    return isObj(tv) ? tv : null;
  }

  function fieldStatus(rec, field) {
    var tv = getTv(rec, field);
    return tv === null ? null : own(tv, 's');
  }

  function isKnown(rec, field) {
    var s = fieldStatus(rec, field);
    return typeof s === 'string' && KNOWN_STATUSES.indexOf(s) >= 0;
  }

  function fv(rec, field) { return own(getTv(rec, field), 'v'); }

  function hasFieldKey(rec, field) {
    var fields = or(own(rec, 'fields'), {});
    return isObj(fields) && hasOwn.call(fields, field);
  }

  // -------------------------------------------------------------------------
  // Placeholders
  // -------------------------------------------------------------------------
  function Unresolved(placeholder) {
    this.placeholder = placeholder;
    this.message = 'unresolved placeholder: ' + placeholder;
  }

  function placeholderContext(config, periodId) {
    var period = null;
    var periods = or(own(config, 'periods'), []);
    for (var i = 0; i < periods.length; i++) {
      if (own(periods[i], 'id') === periodId) {
        period = periods[i];
        break;
      }
    }
    var val = or(own(config, 'valuation'), {});
    var p = or(period, {});
    return {
      '{REV}': own(p, 'rev_field'),
      '{REV_PREV}': own(p, 'prev_rev_field'),
      '{EXP}': own(p, 'exp_field'),
      '{VAL}': own(val, 'field')
    };
  }

  function resolveField(name, ctx) {
    if (typeof name === 'string' && name.charAt(0) === '{' && name.charAt(name.length - 1) === '}') {
      var resolved = hasOwn.call(ctx, name) ? ctx[name] : null;
      if (!truthy(resolved)) throw new Unresolved(name);
      return resolved;
    }
    if (typeof name !== 'string' || !name) throw new Unresolved(keyStr(name));
    return name;
  }

  function singleOp(where) {
    if (!isObj(where)) throw new Error('Where غير صالح: ' + JSON.stringify(where));
    var keys = Object.keys(where);
    if (keys.length !== 1) throw new Error('Where غير صالح: ' + JSON.stringify(where));
    return [keys[0], where[keys[0]]];
  }

  function resolveWhere(where, ctx) {
    if (where === null || where === undefined) return null;
    var oa = singleOp(where), op = oa[0], arg = oa[1];
    var out = {};
    if (op === 'all' || op === 'any') {
      out[op] = arg.map(function (w) { return resolveWhere(w, ctx); });
      return out;
    }
    if (op === 'not') return { not: resolveWhere(arg, ctx) };
    if (op === 'known') return { known: resolveField(arg, ctx) };
    if (op === 'status' || op === 'gt') {
      out[op] = [resolveField(arg[0], ctx), arg[1]];
      return out;
    }
    if (op === 'eq' || op === 'in' || op === 'has') {
      out[op] = arg.slice();
      return out;
    }
    throw new Error('عامل Where غير معروف: ' + op);
  }

  // -------------------------------------------------------------------------
  // Where predicate
  // -------------------------------------------------------------------------
  function evalWhere(where, rec) {
    if (where === null || where === undefined) return true;
    var oa = singleOp(where), op = oa[0], arg = oa[1];
    var i, val;
    switch (op) {
      case 'all':
        for (i = 0; i < arg.length; i++) if (!evalWhere(arg[i], rec)) return false;
        return true;
      case 'any':
        for (i = 0; i < arg.length; i++) if (evalWhere(arg[i], rec)) return true;
        return false;
      case 'not':
        return !evalWhere(arg, rec);
      case 'eq':
        return strictEq(getAttr(rec, arg[0]), arg[1]);
      case 'in': {
        val = getAttr(rec, arg[0]);
        var opts = or(arg[1], []);
        for (i = 0; i < opts.length; i++) if (strictEq(val, opts[i])) return true;
        return false;
      }
      case 'has':
        val = getAttr(rec, arg[0]);
        if (Array.isArray(val)) {
          for (i = 0; i < val.length; i++) if (strictEq(val[i], arg[1])) return true;
          return false;
        }
        return strictEq(val, arg[1]);
      case 'known':
        return isKnown(rec, arg);
      case 'status': {
        var st = fieldStatus(rec, arg[0]);
        return st !== null && inList(st, arg[1]);
      }
      case 'gt':
        if (!isKnown(rec, arg[0])) return false;
        return fv(rec, arg[0]) > arg[1];
      default:
        throw new Error('عامل Where غير معروف: ' + op);
    }
  }

  function andWhere() {
    var items = [];
    for (var i = 0; i < arguments.length; i++) {
      if (arguments[i] !== null && arguments[i] !== undefined) items.push(arguments[i]);
    }
    if (!items.length) return null;
    if (items.length === 1) return items[0];
    return { all: items };
  }

  // -------------------------------------------------------------------------
  // Result helpers
  // -------------------------------------------------------------------------
  function texts(spec, period) {
    var out = {};
    TEXT_KEYS.forEach(function (k) { out[k] = own(spec, k); });
    var tbp = (period !== null && period !== undefined) ? own(or(own(spec, 'text_by_period'), {}), period) : null;
    if (isObj(tbp)) {
      TEXT_KEYS.forEach(function (k) { if (own(tbp, k) !== null) out[k] = tbp[k]; });
    }
    return out;
  }

  function baseResult(spec, period) {
    var res = {
      id: own(spec, 'id'), kind: own(spec, 'kind'), period: nn(period),
      value: null, status: null, unit: own(spec, 'unit'),
      n_scope: 0, n_included: 0, included_ids: [], excluded: [],
      resolved_fields: {}, reason: null, empty_cause: null, note: null,
      card: truthy(hasOwn.call(spec, 'card') ? spec.card : false),
      card_order: own(spec, 'card_order'),
      card_group: own(spec, 'card_group'), kind_tag: own(spec, 'kind_tag')
    };
    var t = texts(spec, period);
    TEXT_KEYS.forEach(function (k) { res[k] = t[k]; });
    // analyst flag «يحتاج تحقق» (IMPLEMENTATION_NOTES review): copied unchanged under every filter and period
    var rv = own(spec, 'review');
    if (isObj(rv)) res.review = JSON.parse(JSON.stringify(rv));
    return res;
  }

  function ids(recs) { return recs.map(function (r) { return r.id; }); }

  function finishStatus(res, R, included, excluded, emptyReason) {
    res.n_included = included.length;
    res.included_ids = ids(included);
    res.excluded = excluded;
    if (!R.length) {
      res.status = 'empty';
    } else if (!included.length) {
      res.status = 'not_computable';
      res.reason = or(res.reason, emptyReason);
    } else if (excluded.length) {
      res.status = 'partial';
    } else {
      res.status = 'complete';
    }
  }

  function statusOrAbsent(rec, field) {
    return or(fieldStatus(rec, field), 'absent');
  }

  // -------------------------------------------------------------------------
  // Kinds
  // -------------------------------------------------------------------------
  function kCount(eng, spec, R, ctx, res) {
    res.value = R.length;
    res.n_included = R.length;
    res.included_ids = ids(R);
    res.status = R.length ? 'complete' : 'empty';
  }

  function kSum(eng, spec, R, ctx, res) {
    var field = resolveField(own(spec, 'field'), ctx);
    res.resolved_fields = { field: field };
    var K = R.filter(function (r) { return isKnown(r, field); });
    var U = R.filter(function (r) { return !isKnown(r, field); });
    var total = 0;
    K.forEach(function (r) { total = total + fv(r, field); });
    var excluded = U.map(function (r) { return { id: r.id, reason: statusOrAbsent(r, field) }; });
    finishStatus(res, R, K, excluded, 'لا توجد قيم معروفة للحقل ضمن النطاق');
    res.n_unknown = U.length;
    res.value = (R.length && K.length) ? total : null;
  }

  function kRatio(eng, spec, R, ctx, res) {
    var numF = resolveField(own(spec, 'num'), ctx);
    var denF = resolveField(own(spec, 'den'), ctx);
    var scale = hasOwn.call(spec, 'scale') ? spec.scale : 1;
    if (scale === null || scale === undefined) scale = 1;
    res.resolved_fields = { num: numF, den: denF };
    var I = [], excluded = [];
    R.forEach(function (r) {
      if (!isKnown(r, numF)) excluded.push({ id: r.id, reason: 'num_unknown' });
      else if (!isKnown(r, denF)) excluded.push({ id: r.id, reason: 'den_unknown' });
      else if (!(fv(r, denF) > 0)) excluded.push({ id: r.id, reason: 'den_nonpositive' });
      else I.push(r);
    });
    var numSum = 0, denSum = 0;
    I.forEach(function (r) {
      numSum = numSum + fv(r, numF);
      denSum = denSum + fv(r, denF);
    });
    var numKnownTotal = 0;
    R.forEach(function (r) { if (isKnown(r, numF)) numKnownTotal = numKnownTotal + fv(r, numF); });
    finishStatus(res, R, I, excluded, 'لا توجد سجلات يتوافر لها البسط والمقام معًا');
    res.num = numSum;
    res.den = denSum;
    res.value = I.length ? (numSum / denSum * scale) : null;
    res.num_known_total = numKnownTotal;
    res.coverage_num_share = numKnownTotal !== 0 ? (numSum / numKnownTotal) : null;
    res.scale = scale;
  }

  function kDiff(eng, spec, R, ctx, res) {
    var aF = resolveField(own(spec, 'a'), ctx);
    var bF = resolveField(own(spec, 'b'), ctx);
    res.resolved_fields = { a: aF, b: bF };
    if (!R.length) { res.status = 'empty'; return; }
    if (!R.some(function (r) { return hasFieldKey(r, bF); })) {
      res.status = 'not_computable';
      res.reason = NO_EXPENSE_REASON;
      return;
    }
    var I = [], excluded = [];
    R.forEach(function (r) {
      if (!isKnown(r, aF)) excluded.push({ id: r.id, reason: 'a_unknown' });
      else if (!isKnown(r, bF)) excluded.push({ id: r.id, reason: 'b_unknown' });
      else I.push(r);
    });
    var value = 0, aSum = 0, bSum = 0;
    I.forEach(function (r) {
      value = value + (fv(r, aF) - fv(r, bF));
      aSum = aSum + fv(r, aF);
      bSum = bSum + fv(r, bF);
    });
    finishStatus(res, R, I, excluded, 'لا توجد سجلات يتوافر لها الإيراد والمصروف معًا');
    res.value = I.length ? value : null;
    res.a_sum = aSum;
    res.b_sum = bSum;
  }

  function kGrowth(eng, spec, R, ctx, res, f) {
    var prevF = resolveField(hasOwn.call(spec, 'prev') ? nn(spec.prev) : '{REV_PREV}', ctx);
    var currF = resolveField(hasOwn.call(spec, 'curr') ? nn(spec.curr) : '{REV}', ctx);
    res.resolved_fields = { prev: prevF, curr: currF };
    var period = eng._period(f.period);
    var comp = or(own(period, 'comparable_with_prev'), { status: 'unverified', reason: null });
    res.comparability = comp;
    res.prev_period = own(period, 'prev_id');
    if (own(comp, 'status') === 'not_comparable') {
      res.status = 'not_comparable';
      res.value = null;
      res.reason = or(own(comp, 'reason'), 'الفترتان غير قابلتين للمقارنة');
      return;
    }
    if (!R.length) { res.status = 'empty'; return; }
    var I = [], excluded = [], onlyCurr = [], onlyPrev = [], starts = [];
    R.forEach(function (r) {
      var kp = isKnown(r, prevF), kc = isKnown(r, currF);
      if (kp && kc) {
        if (fv(r, prevF) === 0 && fv(r, currF) > 0) {
          // revenue start: shown as «بدء إيراد» with its riyal change, never inside the growth %
          starts.push(r);
          excluded.push({ id: r.id, reason: 'revenue_start' });
        } else {
          I.push(r);
        }
      } else if (kc) { onlyCurr.push(r.id); excluded.push({ id: r.id, reason: 'prev_unknown' }); }
      else if (kp) { onlyPrev.push(r.id); excluded.push({ id: r.id, reason: 'curr_unknown' }); }
      else excluded.push({ id: r.id, reason: 'both_unknown' });
    });
    var prevSum = 0, currSum = 0;
    I.forEach(function (r) {
      prevSum = prevSum + fv(r, prevF);
      currSum = currSum + fv(r, currF);
    });
    var startChange = 0;
    starts.forEach(function (r) { startChange = startChange + fv(r, currF); });
    finishStatus(res, R, I, excluded, 'لا توجد سجلات لها إيراد معروف في الفترتين');
    var has = I.length > 0;
    var absChange = currSum - prevSum;
    var pct = (has && prevSum > 0) ? (absChange / prevSum * 100) : null;
    res.prev_sum = has ? prevSum : null;
    res.curr_sum = has ? currSum : null;
    res.abs_change = has ? absChange : null;
    res.pct = pct;
    res.revenue_start = starts.length > 0;
    res.revenue_start_ids = ids(starts);
    res.start_change = starts.length ? startChange : null;
    res.only_curr = onlyCurr;
    res.only_prev = onlyPrev;
    res.value = pct;
    if (!has && starts.length) {
      res.status = 'not_computable';
      res.reason = REVENUE_START_ONLY_REASON;
    } else if (has && !(prevSum > 0)) {
      res.status = 'not_computable';
      res.reason = PREV_NONPOSITIVE_REASON;
    }
  }

  function kConcentration(eng, spec, R, ctx, res) {
    var field = resolveField(own(spec, 'field'), ctx);
    var topRaw = hasOwn.call(spec, 'top') ? spec.top : 3;
    var top = Math.trunc(Number(or(topRaw, 3)));
    res.resolved_fields = { field: field };
    res.top = top;
    if (!R.length) { res.status = 'empty'; return; }
    var K = R.filter(function (r) { return isKnown(r, field) && fv(r, field) > 0; });
    var excluded = [];
    var anyUnknown = false;
    R.forEach(function (r) {
      if (!isKnown(r, field)) {
        anyUnknown = true;
        excluded.push({ id: r.id, reason: statusOrAbsent(r, field) });
      } else if (!(fv(r, field) > 0)) {
        excluded.push({ id: r.id, reason: 'nonpositive' });
      }
    });
    var total = 0;
    K.forEach(function (r) { total = total + fv(r, field); }); // id order
    var ranked = K.slice().sort(function (a, b) {
      var va = fv(a, field), vb = fv(b, field);
      if (va > vb) return -1;
      if (va < vb) return 1;
      return cmpStr(a.id, b.id);
    });
    var topRecs = ranked.slice(0, top);
    var topSum = 0;
    topRecs.forEach(function (r) { topSum = topSum + fv(r, field); });
    res.n_included = K.length;
    res.included_ids = ids(K);
    res.excluded = excluded;
    res.top_ids = ids(topRecs);
    res.total = total;
    res.top_sum = topSum;
    if (total !== 0) {
      res.value = topSum / total * 100;
      var h = 0;
      ranked.forEach(function (r) { var s = fv(r, field) / total; h = h + s * s; });
      res.hhi = h * 10000;
      res.top_shares = topRecs.map(function (r) {
        return { id: r.id, value: fv(r, field), share: fv(r, field) / total * 100 };
      });
    } else {
      res.value = null;
      res.hhi = null;
      res.top_shares = [];
    }
    if (!K.length || total === 0) {
      res.status = 'not_computable';
      res.reason = 'لا توجد قيم موجبة معروفة ضمن النطاق';
      res.value = null;
    } else if (anyUnknown) {
      res.status = 'partial';
    } else {
      res.status = 'complete';
    }
  }

  function kCoverage(eng, spec, R, ctx, res) {
    var fields = or(own(spec, 'fields'), []).map(function (x) { return resolveField(x, ctx); });
    res.resolved_fields = { fields: fields };
    var n = R.length;
    var perField = fields.map(function (fld) {
      var nk = 0, nna = 0;
      R.forEach(function (r) {
        if (isKnown(r, fld)) nk += 1;
        else if (fieldStatus(r, fld) === 'not_applicable') nna += 1;
      });
      // «لا ينطبق» is not missing data: it leaves this field's denominator
      return { field: fld, n_known: nk, n: n - nna, n_not_applicable: nna };
    });
    res.per_field = perField;
    res.n_included = n;
    res.included_ids = ids(R);
    if (n === 0) {
      res.status = 'empty';
      res.value = null;
      return;
    }
    var ratios = perField.filter(function (pf) { return pf.n > 0; }).map(function (pf) { return pf.n_known / pf.n; });
    if (!ratios.length) {
      res.value = null;
      res.status = 'not_computable';
      res.reason = NO_COVERAGE_FIELD_REASON;
      return;
    }
    res.value = Math.min.apply(null, ratios) * 100;
    res.status = 'complete';
  }

  var KINDS = {
    count: kCount, sum: kSum, ratio: kRatio, diff: kDiff,
    growth: kGrowth, concentration: kConcentration, coverage: kCoverage
  };

  // -------------------------------------------------------------------------
  // Engine
  // -------------------------------------------------------------------------
  function Engine(config, records) {
    this.config = config || {};
    this.records = (records || []).slice().sort(byId);
    this.periods = or(own(this.config, 'periods'), []).slice();
    this.periodIds = this.periods.map(function (p) { return own(p, 'id'); });
    this.defaultPeriod = this.periodIds.length ? this.periodIds[0] : null;
    var listDims = {};
    this.records.forEach(function (r) {
      var dims = or(own(r, 'dims'), {});
      Object.keys(dims).forEach(function (k) { if (Array.isArray(dims[k])) listDims[k] = true; });
    });
    this.listDims = listDims;
  }

  Engine.prototype.normalizeFilters = function (filters) {
    var src = {};
    if (isObj(filters)) Object.keys(filters).forEach(function (k) { src[k] = filters[k]; });
    var period = hasOwn.call(src, 'period') ? nn(src.period) : null;
    delete src.period;
    if (period === null || period === '') period = this.defaultPeriod;
    if (this.periodIds.length && this.periodIds.indexOf(period) < 0) {
      throw new Error('الفترة غير معروفة: ' + JSON.stringify(period) + '؛ المتاح: ' + JSON.stringify(this.periodIds));
    }
    var out = { period: period };
    sortedStrings(Object.keys(src)).forEach(function (k) {
      var v = src[k];
      if (v === null || v === undefined || v === '') return;
      out[k] = v;
    });
    return out;
  };

  Engine.prototype.filterWhere = function (filters) {
    var self = this;
    var clauses = [];
    sortedStrings(Object.keys(filters)).forEach(function (k) {
      if (k === 'period') return;
      var c = {};
      c[hasOwn.call(self.listDims, k) ? 'has' : 'eq'] = [k, filters[k]];
      clauses.push(c);
    });
    return clauses.length ? { all: clauses } : null;
  };

  Engine.prototype._select = function (where) {
    return this.records.filter(function (r) { return evalWhere(where, r); });
  };

  Engine.prototype.filterRecords = function (filters) {
    var f = this.normalizeFilters(filters);
    return this._select(andWhere(own(this.config, 'base_where'), this.filterWhere(f)));
  };

  Engine.prototype.context = function (periodId) { return placeholderContext(this.config, periodId); };

  Engine.prototype._period = function (periodId) {
    for (var i = 0; i < this.periods.length; i++) {
      if (own(this.periods[i], 'id') === periodId) return this.periods[i];
    }
    return {};
  };

  Engine.prototype.resolveField = function (name, filters) {
    var f = this.normalizeFilters(filters);
    try {
      return resolveField(name, this.context(f.period));
    } catch (e) {
      if (e instanceof Unresolved) return null;
      throw e;
    }
  };

  Engine.prototype.compute = function (spec, filters) {
    var f = this.normalizeFilters(filters);
    var ctx = this.context(f.period);
    var res = baseResult(spec, f.period);
    var fw = this.filterWhere(f);
    var base = own(this.config, 'base_where');
    var where = null, unresolved = null;
    try {
      where = resolveWhere(own(spec, 'where'), ctx);
    } catch (e) {
      if (!(e instanceof Unresolved)) throw e;
      where = null;
      unresolved = e.placeholder;
    }
    var vp = own(spec, 'valid_periods');
    if (Array.isArray(vp) && vp.indexOf(f.period) < 0) {
      var scopeW = unresolved !== null ? andWhere(base, fw) : andWhere(base, where, fw);
      res.status = 'not_comparable';
      res.n_scope = this._select(scopeW).length;
      res.reason = or(own(spec, 'invalid_period_reason'), INVALID_PERIOD_REASON);
      return res;
    }
    if (unresolved !== null) {
      var rScope = this._select(andWhere(base, fw));
      res.status = 'not_computable';
      res.n_scope = rScope.length;
      res.reason = or(res.unavailable_reason, DEFAULT_UNAVAILABLE);
      res.unresolved = unresolved;
      if (rScope.length) res.scope_ids = rScope.map(function (r) { return r.id; });
      return res;
    }
    var R = this._select(andWhere(base, where, fw));
    res.n_scope = R.length;
    var kind = own(spec, 'kind');
    if (typeof kind !== 'string' || !hasOwn.call(KINDS, kind)) {
      throw new Error('نوع مؤشر غير معروف: ' + JSON.stringify(kind) + ' (' + own(spec, 'id') + ')');
    }
    try {
      KINDS[kind](this, spec, R, ctx, res, f);
    } catch (e) {
      if (!(e instanceof Unresolved)) throw e;
      res.status = 'not_computable';
      res.value = null;
      res.n_included = 0;
      res.included_ids = [];
      res.excluded = [];
      res.reason = or(res.unavailable_reason, DEFAULT_UNAVAILABLE);
      res.unresolved = e.placeholder;
    }
    // records of the scope of a not_computable result that included and excluded none (IMPLEMENTATION_NOTES scope_ids)
    if (res.status === 'not_computable' && R.length && !res.included_ids.length && !res.excluded.length) {
      res.scope_ids = R.map(function (r) { return r.id; });
    }
    if (res.status === 'empty') {
      var cause = this._emptyCause(base, where, fw);
      res.empty_cause = cause;
      if (cause === 'filters') {
        res.reason = EMPTY_FILTERS_REASON;
      } else {
        var text = or(res.empty_text, EMPTY_SCOPE_REASON);
        var eaz = hasOwn.call(spec, 'empty_as_zero') ? truthy(spec.empty_as_zero) : kind === 'count';
        if ((kind === 'sum' || kind === 'count') && eaz) {
          res.value = 0;
          res.status = 'complete';
          res.reason = null;
          res.note = text;
        } else {
          res.reason = text;
        }
      }
    }
    return res;
  };

  /** 'filters' when the dashboard dim filters alone empty the metric's scope, else 'scope'. */
  Engine.prototype._emptyCause = function (base, where, fw) {
    if (fw !== null && this._select(andWhere(base, where)).length > 0) return 'filters';
    return 'scope';
  };

  Engine.prototype.computeBreakdown = function (spec, filters) {
    var f = this.normalizeFilters(filters);
    var ctx = this.context(f.period);
    var res = baseResult(spec, f.period);
    res.kind = 'breakdown';
    res.by = own(spec, 'by');
    res.rows = [];
    res.total = null;
    res.n_known = 0;
    res.n_unknown = 0;
    res.overlap = false;
    res.note = null;
    var base = own(this.config, 'base_where');
    var fw = this.filterWhere(f);
    var field, where;
    try {
      field = resolveField(own(spec, 'field'), ctx);
      where = resolveWhere(own(spec, 'where'), ctx);
    } catch (e) {
      if (!(e instanceof Unresolved)) throw e;
      res.status = 'not_computable';
      res.n_scope = this._select(andWhere(base, fw)).length;
      res.reason = or(res.unavailable_reason, DEFAULT_UNAVAILABLE);
      res.unresolved = e.placeholder;
      return res;
    }
    res.resolved_fields = { field: field };
    var R = this._select(andWhere(base, where, fw));
    res.n_scope = R.length;
    var by = own(spec, 'by');
    if (!R.length) {
      var cause = this._emptyCause(base, where, fw);
      res.status = 'empty';
      res.empty_cause = cause;
      res.reason = cause === 'filters' ? EMPTY_FILTERS_REASON : or(res.empty_text, EMPTY_SCOPE_REASON);
      return res;
    }
    var K = R.filter(function (r) { return isKnown(r, field); });
    var U = R.filter(function (r) { return !isKnown(r, field); });
    var groups = new Map();
    var overlap = false;

    function keysOf(rec) {
      var val = getAttr(rec, by);
      if (Array.isArray(val)) {
        overlap = true;
        var ks = val.map(function (x) { return (x === null || x === undefined) ? NONE_KEY : x; });
        return ks.length ? ks : [NONE_KEY];
      }
      return [val === null ? NONE_KEY : val];
    }

    function group(key) {
      var gk = keyStr(key);
      if (!groups.has(gk)) groups.set(gk, { key: key, n_known: 0, n_unknown: 0, ids: [], _sum: 0 });
      return groups.get(gk);
    }

    var total = 0;
    K.forEach(function (r) {
      total = total + fv(r, field);
      var seen = {};
      keysOf(r).forEach(function (k) {
        var ks = keyStr(k);
        if (hasOwn.call(seen, ks)) return;
        seen[ks] = true;
        var g = group(k);
        g._sum = g._sum + fv(r, field);
        g.n_known += 1;
        g.ids.push(r.id);
      });
    });
    U.forEach(function (r) {
      var seen = {};
      keysOf(r).forEach(function (k) {
        var ks = keyStr(k);
        if (hasOwn.call(seen, ks)) return;
        seen[ks] = true;
        var g = group(k);
        g.n_unknown += 1;
        g.ids.push(r.id);
      });
    });
    var rows = [];
    groups.forEach(function (g) {
      var value = g.n_known > 0 ? g._sum : null;
      var share = (value !== null && total !== 0) ? (value / total) : null;
      rows.push({ key: g.key, value: value, share: share, n_known: g.n_known, n_unknown: g.n_unknown,
        ids: sortedStrings(g.ids) });
    });
    rows.sort(function (a, b) {
      var na = a.value === null ? 1 : 0, nb = b.value === null ? 1 : 0;
      if (na !== nb) return na - nb;
      var va = -(a.value || 0), vb = -(b.value || 0);
      if (va < vb) return -1;
      if (va > vb) return 1;
      return cmpStr(keyStr(a.key), keyStr(b.key));
    });
    res.rows = rows;
    res.total = K.length ? total : null;
    res.n_known = K.length;
    res.n_unknown = U.length;
    res.overlap = overlap;
    res.note = overlap ? OVERLAP_NOTE : null;
    res.n_included = K.length;
    res.included_ids = ids(K);
    res.excluded = U.map(function (r) { return { id: r.id, reason: statusOrAbsent(r, field) }; });
    if (!K.length) {
      res.status = 'not_computable';
      res.reason = 'لا توجد قيم معروفة للحقل ضمن النطاق';
    } else if (U.length) {
      res.status = 'partial';
    } else {
      res.status = 'complete';
    }
    return res;
  };

  Engine.prototype.computeTrend = function (filters) {
    var f = this.normalizeFilters(filters);
    var tcfg = or(own(this.config, 'trend'), {});
    var comparable = or(own(tcfg, 'comparable'), { status: 'unverified', reason: null });
    var out = {
      id: 'trend', label: own(tcfg, 'label'), status: null, points: [], comparable: comparable,
      not_comparable: own(comparable, 'status') === 'not_comparable', n_scope: 0, reason: null,
      unit: hasOwn.call(tcfg, 'unit') ? nn(tcfg.unit) : 'SAR',
      same_records: null, n_known_varies: false, empty_cause: null
    };
    if (!truthy(tcfg) || !truthy(own(tcfg, 'enabled')) || !truthy(own(tcfg, 'fields'))) {
      out.status = 'not_computable';
      out.reason = 'التطور الزمني غير مفعّل أو لا توجد فترات';
      return out;
    }
    var base = own(this.config, 'base_where');
    var ctx = this.context(f.period);
    var where;
    try {
      where = resolveWhere(own(tcfg, 'where'), ctx);
    } catch (e) {
      if (!(e instanceof Unresolved)) throw e;
      where = null;
    }
    var fw = this.filterWhere(f);
    var R = this._select(andWhere(base, where, fw));
    out.n_scope = R.length;
    if (!R.length) {
      var cause = this._emptyCause(base, where, fw);
      out.status = 'empty';
      out.empty_cause = cause;
      out.reason = cause === 'filters' ? EMPTY_FILTERS_REASON : EMPTY_SCOPE_REASON;
      out.points = tcfg.fields.map(function (t) {
        return { period: own(t, 'period'), field: own(t, 'field'), value: null, n_known: 0, n: 0 };
      });
      out.same_records = { n: 0, ids: [], points: tcfg.fields.map(function (t) {
        return { period: own(t, 'period'), field: own(t, 'field'), value: null, n_known: 0 };
      }) };
      return out;
    }
    var anyKnown = false, anyPartial = false;
    tcfg.fields.forEach(function (t) {
      var field = own(t, 'field');
      var s = 0, nKnown = 0;
      R.forEach(function (r) {
        if (isKnown(r, field)) { s = s + fv(r, field); nKnown += 1; }
      });
      if (nKnown) anyKnown = true;
      if (nKnown < R.length) anyPartial = true;
      out.points.push({ period: own(t, 'period'), field: field, value: nKnown ? s : null, n_known: nKnown, n: R.length });
    });
    // the same records in every period: a change between points is not a change in the record set
    var same = R.filter(function (r) {
      for (var i = 0; i < tcfg.fields.length; i++) if (!isKnown(r, own(tcfg.fields[i], 'field'))) return false;
      return true;
    });
    var srPoints = tcfg.fields.map(function (t) {
      var field = own(t, 'field');
      var s = 0;
      same.forEach(function (r) { s = s + fv(r, field); });
      return { period: own(t, 'period'), field: field, value: same.length ? s : null, n_known: same.length };
    });
    out.same_records = { n: same.length, ids: ids(same), points: srPoints };
    out.n_known_varies = out.points.some(function (p) { return p.n_known !== same.length; });
    if (!anyKnown) {
      out.status = 'not_computable';
      out.reason = 'لا توجد قيم إيراد معروفة ضمن النطاق';
    } else if (anyPartial) {
      out.status = 'partial';
    } else {
      out.status = 'complete';
    }
    return out;
  };

  Engine.prototype.computeAll = function (filters) {
    var self = this;
    var f = this.normalizeFilters(filters);
    var metrics = {};
    or(own(this.config, 'metrics'), []).forEach(function (spec) { metrics[spec.id] = self.compute(spec, f); });
    var breakdowns = {};
    or(own(this.config, 'breakdowns'), []).forEach(function (spec) {
      breakdowns[spec.id] = self.computeBreakdown(spec, f);
    });
    var recIds = ids(this.filterRecords(f));
    return { filters: f, metrics: metrics, breakdowns: breakdowns, trend: this.computeTrend(f),
      records: recIds, record_ids: recIds };
  };

  /** {value, status, n_included}: the fields compared by the dashboard parity checks. */
  function compact(result) {
    return { value: nn(result.value), status: nn(result.status), n_included: nn(result.n_included) };
  }

  /**
   * createEngine(data): data = analysis.json ({records, config, metrics, ...}).
   * Returns plain functions (safe to destructure) plus the underlying engine.
   */
  function createEngine(data) {
    data = data || {};
    var eng = new Engine(data.config || {}, data.records || []);
    return {
      engine: eng,
      config: eng.config,
      periodIds: eng.periodIds.slice(),
      defaultPeriod: eng.defaultPeriod,
      listDims: Object.keys(eng.listDims).sort(cmpStr),
      compute: function (spec, filters) { return eng.compute(spec, filters); },
      computeBreakdown: function (spec, filters) { return eng.computeBreakdown(spec, filters); },
      computeTrend: function (filters) { return eng.computeTrend(filters); },
      computeAll: function (filters) { return eng.computeAll(filters); },
      filterRecords: function (filters) { return eng.filterRecords(filters); },
      normalizeFilters: function (filters) { return eng.normalizeFilters(filters); },
      resolveField: function (name, filters) { return eng.resolveField(name, filters); },
      compact: compact
    };
  }

  var api = {
    createEngine: createEngine,
    Engine: Engine,
    evalWhere: evalWhere,
    resolveWhere: resolveWhere,
    strictEq: strictEq,
    compareStrings: cmpStr,
    compact: compact,
    KNOWN_STATUSES: KNOWN_STATUSES.slice(),
    PLACEHOLDERS: PLACEHOLDERS.slice(),
    NO_EXPENSE_REASON: NO_EXPENSE_REASON,
    DEFAULT_UNAVAILABLE: DEFAULT_UNAVAILABLE,
    INVALID_PERIOD_REASON: INVALID_PERIOD_REASON,
    EMPTY_FILTERS_REASON: EMPTY_FILTERS_REASON,
    EMPTY_SCOPE_REASON: EMPTY_SCOPE_REASON,
    REVENUE_START_ONLY_REASON: REVENUE_START_ONLY_REASON,
    PREV_NONPOSITIVE_REASON: PREV_NONPOSITIVE_REASON,
    NO_COVERAGE_FIELD_REASON: NO_COVERAGE_FIELD_REASON,
    NONE_KEY: NONE_KEY,
    OVERLAP_NOTE: OVERLAP_NOTE
  };

  var hasModule = typeof module === 'object' && module !== null && typeof module.exports === 'object';
  if (hasModule) module.exports = api;
  if (typeof window !== 'undefined') {
    window.createEngine = createEngine;
    window.APAEngine = api;
  } else if (!hasModule && typeof globalThis !== 'undefined') {
    globalThis.createEngine = createEngine;
    globalThis.APAEngine = api;
  }
})();
