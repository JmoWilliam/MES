/**
 * @license
 *
 * DHTMLX Kanban v.1.1.1 Professional
 *
 * This software is covered by DHTMLX Commercial License.
 * Usage without proper license is prohibited.
 *
 * (c) XB Software.
 */
var kanban = (function (e) {
    "use strict";
    var t = function () {
        return (t =
            Object.assign ||
            function (e) {
                for (var t, n = 1, l = arguments.length; n < l; n++) for (var o in (t = arguments[n])) Object.prototype.hasOwnProperty.call(t, o) && (e[o] = t[o]);
                return e;
            }).apply(this, arguments);
    };
    function n() { }
    const l = (e) => e;
    function o(e, t) {
        for (const n in t) e[n] = t[n];
        return e;
    }
    function c(e) {
        return e();
    }
    function s() {
        return Object.create(null);
    }
    function r(e) {
        e.forEach(c);
    }
    function i(e) {
        return "function" == typeof e;
    }
    function a(e, t) {
        return e != e ? t == t : e !== t || (e && "object" == typeof e) || "function" == typeof e;
    }
    let u;
    function d(e, t) {
        return u || (u = document.createElement("a")), (u.href = t), e === u.href;
    }
    function p(e, ...t) {
        if (null == e) return n;
        const l = e.subscribe(...t);
        return l.unsubscribe ? () => l.unsubscribe() : l;
    }
    function f(e, t, n) {
        e.$$.on_destroy.push(p(t, n));
    }
    function $(e, t, n, l) {
        if (e) {
            const o = m(e, t, n, l);
            return e[0](o);
        }
    }
    function m(e, t, n, l) {
        return e[1] && l ? o(n.ctx.slice(), e[1](l(t))) : n.ctx;
    }
    function h(e, t, n, l) {
        if (e[2] && l) {
            const o = e[2](l(n));
            if (void 0 === t.dirty) return o;
            if ("object" == typeof o) {
                const e = [],
                    n = Math.max(t.dirty.length, o.length);
                for (let l = 0; l < n; l += 1) e[l] = t.dirty[l] | o[l];
                return e;
            }
            return t.dirty | o;
        }
        return t.dirty;
    }
    function g(e, t, n, l, o, c) {
        if (o) {
            const s = m(t, n, l, c);
            e.p(s, o);
        }
    }
    function v(e) {
        if (e.ctx.length > 32) {
            const t = [],
                n = e.ctx.length / 32;
            for (let e = 0; e < n; e++) t[e] = -1;
            return t;
        }
        return -1;
    }
    function y(e) {
        const t = {};
        for (const n in e) "$" !== n[0] && (t[n] = e[n]);
        return t;
    }
    function w(e, t) {
        const n = {};
        t = new Set(t);
        for (const l in e) t.has(l) || "$" === l[0] || (n[l] = e[l]);
        return n;
    }
    function b(e) {
        return null == e ? "" : e;
    }
    function x(e, t, n) {
        return e.set(n), t;
    }
    function k(e) {
        return e && i(e.destroy) ? e.destroy : n;
    }
    const S = "undefined" != typeof window;
    let M = S ? () => window.performance.now() : () => Date.now(),
        _ = S ? (e) => requestAnimationFrame(e) : n;
    const C = new Set();
    function D(e) {
        C.forEach((t) => {
            t.c(e) || (C.delete(t), t.f());
        }),
            0 !== C.size && _(D);
    }
    function I(e, t) {
        e.appendChild(t);
    }
    function A(e) {
        if (!e) return document;
        const t = e.getRootNode ? e.getRootNode() : e.ownerDocument;
        return t.host ? t : document;
    }
    function E(e) {
        const t = z("style");
        return (
            (function (e, t) {
                I(e.head || e, t);
            })(A(e), t),
            t
        );
    }
    function T(e, t, n) {
        e.insertBefore(t, n || null);
    }
    function L(e) {
        e.parentNode.removeChild(e);
    }
    function j(e, t) {
        for (let n = 0; n < e.length; n += 1) e[n] && e[n].d(t);
    }
    function z(e) {
        return document.createElement(e);
    }
    function N(e) {
        return document.createElementNS("http://www.w3.org/2000/svg", e);
    }
    function O(e) {
        return document.createTextNode(e);
    }
    function F() {
        return O(" ");
    }
    function q() {
        return O("");
    }
    function R(e, t, n, l) {
        return e.addEventListener(t, n, l), () => e.removeEventListener(t, n, l);
    }
    function P(e) {
        return function (t) {
            return t.preventDefault(), e.call(this, t);
        };
    }
    function K(e) {
        return function (t) {
            return t.stopPropagation(), e.call(this, t);
        };
    }
    function U(e, t, n) {
        null == n ? e.removeAttribute(t) : e.getAttribute(t) !== n && e.setAttribute(t, n);
    }
    function H(e) {
        return "" === e ? null : +e;
    }
    function Y(e, t) {
        (t = "" + t), e.wholeText !== t && (e.data = t);
    }
    function B(e, t) {
        e.value = null == t ? "" : t;
    }
    function G(e, t, n, l) {
        e.style.setProperty(t, n, l ? "important" : "");
    }
    function J(e, t) {
        for (let n = 0; n < e.options.length; n += 1) {
            const l = e.options[n];
            if (l.__value === t) return void (l.selected = !0);
        }
        e.selectedIndex = -1;
    }
    function V(e, t, n) {
        e.classList[n ? "add" : "remove"](t);
    }
    function X(e, t, n = !1) {
        const l = document.createEvent("CustomEvent");
        return l.initCustomEvent(e, n, !1, t), l;
    }
    class Q {
        constructor() {
            this.e = this.n = null;
        }
        c(e) {
            this.h(e);
        }
        m(e, t, n = null) {
            this.e || ((this.e = z(t.nodeName)), (this.t = t), this.c(e)), this.i(n);
        }
        h(e) {
            (this.e.innerHTML = e), (this.n = Array.from(this.e.childNodes));
        }
        i(e) {
            for (let t = 0; t < this.n.length; t += 1) T(this.t, this.n[t], e);
        }
        p(e) {
            this.d(), this.h(e), this.i(this.a);
        }
        d() {
            this.n.forEach(L);
        }
    }
    const W = new Set();
    let Z,
        ee = 0;
    function te(e, t, n, l, o, c, s, r = 0) {
        const i = 16.666 / l;
        let a = "{\n";
        for (let e = 0; e <= 1; e += i) {
            const l = t + (n - t) * c(e);
            a += 100 * e + `%{${s(l, 1 - l)}}\n`;
        }
        const u = a + `100% {${s(n, 1 - n)}}\n}`,
            d = `__svelte_${(function (e) {
                let t = 5381,
                    n = e.length;
                for (; n--;) t = ((t << 5) - t) ^ e.charCodeAt(n);
                return t >>> 0;
            })(u)}_${r}`,
            p = A(e);
        W.add(p);
        const f = p.__svelte_stylesheet || (p.__svelte_stylesheet = E(e).sheet),
            $ = p.__svelte_rules || (p.__svelte_rules = {});
        $[d] || (($[d] = !0), f.insertRule(`@keyframes ${d} ${u}`, f.cssRules.length));
        const m = e.style.animation || "";
        return (e.style.animation = `${m ? `${m}, ` : ""}${d} ${l}ms linear ${o}ms 1 both`), (ee += 1), d;
    }
    function ne(e, t) {
        const n = (e.style.animation || "").split(", "),
            l = n.filter(t ? (e) => e.indexOf(t) < 0 : (e) => -1 === e.indexOf("__svelte")),
            o = n.length - l.length;
        o &&
            ((e.style.animation = l.join(", ")),
                (ee -= o),
                ee ||
                _(() => {
                    ee ||
                        (W.forEach((e) => {
                            const t = e.__svelte_stylesheet;
                            let n = t.cssRules.length;
                            for (; n--;) t.deleteRule(n);
                            e.__svelte_rules = {};
                        }),
                            W.clear());
                }));
    }
    function le(e) {
        Z = e;
    }
    function oe() {
        if (!Z) throw new Error("Function called outside component initialization");
        return Z;
    }
    function ce(e) {
        oe().$$.on_mount.push(e);
    }
    function se(e) {
        oe().$$.after_update.push(e);
    }
    function re() {
        const e = oe();
        return (t, n) => {
            const l = e.$$.callbacks[t];
            if (l) {
                const o = X(t, n);
                l.slice().forEach((t) => {
                    t.call(e, o);
                });
            }
        };
    }
    function ie(e, t) {
        oe().$$.context.set(e, t);
    }
    function ae(e) {
        return oe().$$.context.get(e);
    }
    function ue(e, t) {
        const n = e.$$.callbacks[t.type];
        n && n.slice().forEach((e) => e.call(this, t));
    }
    const de = [],
        pe = [],
        fe = [],
        $e = [],
        me = Promise.resolve();
    let he = !1;
    function ge() {
        he || ((he = !0), me.then(xe));
    }
    function ve(e) {
        fe.push(e);
    }
    function ye(e) {
        $e.push(e);
    }
    let we = !1;
    const be = new Set();
    function xe() {
        if (!we) {
            we = !0;
            do {
                for (let e = 0; e < de.length; e += 1) {
                    const t = de[e];
                    le(t), ke(t.$$);
                }
                for (le(null), de.length = 0; pe.length;) pe.pop()();
                for (let e = 0; e < fe.length; e += 1) {
                    const t = fe[e];
                    be.has(t) || (be.add(t), t());
                }
                fe.length = 0;
            } while (de.length);
            for (; $e.length;) $e.pop()();
            (he = !1), (we = !1), be.clear();
        }
    }
    function ke(e) {
        if (null !== e.fragment) {
            e.update(), r(e.before_update);
            const t = e.dirty;
            (e.dirty = [-1]), e.fragment && e.fragment.p(e.ctx, t), e.after_update.forEach(ve);
        }
    }
    let Se;
    function Me(e, t, n) {
        e.dispatchEvent(X(`${t ? "intro" : "outro"}${n}`));
    }
    const _e = new Set();
    let Ce;
    function De() {
        Ce = { r: 0, c: [], p: Ce };
    }
    function Ie() {
        Ce.r || r(Ce.c), (Ce = Ce.p);
    }
    function Ae(e, t) {
        e && e.i && (_e.delete(e), e.i(t));
    }
    function Ee(e, t, n, l) {
        if (e && e.o) {
            if (_e.has(e)) return;
            _e.add(e),
                Ce.c.push(() => {
                    _e.delete(e), l && (n && e.d(1), l());
                }),
                e.o(t);
        }
    }
    const Te = { duration: 0 };
    function Le(e, t, o, c) {
        let s = t(e, o),
            a = c ? 0 : 1,
            u = null,
            d = null,
            p = null;
        function f() {
            p && ne(e, p);
        }
        function $(e, t) {
            const n = e.b - a;
            return (t *= Math.abs(n)), { a: a, b: e.b, d: n, duration: t, start: e.start, end: e.start + t, group: e.group };
        }
        function m(t) {
            const { delay: o = 0, duration: c = 300, easing: i = l, tick: m = n, css: h } = s || Te,
                g = { start: M() + o, b: t };
            t || ((g.group = Ce), (Ce.r += 1)),
                u || d
                    ? (d = g)
                    : (h && (f(), (p = te(e, a, t, c, o, i, h))),
                        t && m(0, 1),
                        (u = $(g, c)),
                        ve(() => Me(e, t, "start")),
                        (function (e) {
                            let t;
                            0 === C.size && _(D),
                                new Promise((n) => {
                                    C.add((t = { c: e, f: n }));
                                });
                        })((t) => {
                            if ((d && t > d.start && ((u = $(d, c)), (d = null), Me(e, u.b, "start"), h && (f(), (p = te(e, a, u.b, u.duration, 0, i, s.css)))), u))
                                if (t >= u.end) m((a = u.b), 1 - a), Me(e, u.b, "end"), d || (u.b ? f() : --u.group.r || r(u.group.c)), (u = null);
                                else if (t >= u.start) {
                                    const e = t - u.start;
                                    (a = u.a + u.d * i(e / u.duration)), m(a, 1 - a);
                                }
                            return !(!u && !d);
                        }));
        }
        return {
            run(e) {
                i(s)
                    ? (Se ||
                        ((Se = Promise.resolve()),
                            Se.then(() => {
                                Se = null;
                            })),
                        Se).then(() => {
                            (s = s()), m(e);
                        })
                    : m(e);
            },
            end() {
                f(), (u = d = null);
            },
        };
    }
    const je = "undefined" != typeof window ? window : "undefined" != typeof globalThis ? globalThis : global;
    function ze(e, t) {
        e.d(1), t.delete(e.key);
    }
    function Ne(e, t) {
        Ee(e, 1, 1, () => {
            t.delete(e.key);
        });
    }
    function Oe(e, t, n, l, o, c, s, r, i, a, u, d) {
        let p = e.length,
            f = c.length,
            $ = p;
        const m = {};
        for (; $--;) m[e[$].key] = $;
        const h = [],
            g = new Map(),
            v = new Map();
        for ($ = f; $--;) {
            const e = d(o, c, $),
                r = n(e);
            let i = s.get(r);
            i ? l && i.p(e, t) : ((i = a(r, e)), i.c()), g.set(r, (h[$] = i)), r in m && v.set(r, Math.abs($ - m[r]));
        }
        const y = new Set(),
            w = new Set();
        function b(e) {
            Ae(e, 1), e.m(r, u), s.set(e.key, e), (u = e.first), f--;
        }
        for (; p && f;) {
            const t = h[f - 1],
                n = e[p - 1],
                l = t.key,
                o = n.key;
            t === n ? ((u = t.first), p-- , f--) : g.has(o) ? (!s.has(l) || y.has(l) ? b(t) : w.has(o) ? p-- : v.get(l) > v.get(o) ? (w.add(l), b(t)) : (y.add(o), p--)) : (i(n, s), p--);
        }
        for (; p--;) {
            const t = e[p];
            g.has(t.key) || i(t, s);
        }
        for (; f;) b(h[f - 1]);
        return h;
    }
    function Fe(e, t) {
        const n = {},
            l = {},
            o = { $$scope: 1 };
        let c = e.length;
        for (; c--;) {
            const s = e[c],
                r = t[c];
            if (r) {
                for (const e in s) e in r || (l[e] = 1);
                for (const e in r) o[e] || ((n[e] = r[e]), (o[e] = 1));
                e[c] = r;
            } else for (const e in s) o[e] = 1;
        }
        for (const e in l) e in n || (n[e] = void 0);
        return n;
    }
    function qe(e) {
        return "object" == typeof e && null !== e ? e : {};
    }
    function Re(e, t, n) {
        const l = e.$$.props[t];
        void 0 !== l && ((e.$$.bound[l] = n), n(e.$$.ctx[l]));
    }
    function Pe(e) {
        e && e.c();
    }
    function Ke(e, t, n, l) {
        const { fragment: o, on_mount: s, on_destroy: a, after_update: u } = e.$$;
        o && o.m(t, n),
            l ||
            ve(() => {
                const t = s.map(c).filter(i);
                a ? a.push(...t) : r(t), (e.$$.on_mount = []);
            }),
            u.forEach(ve);
    }
    function Ue(e, t) {
        const n = e.$$;
        null !== n.fragment && (r(n.on_destroy), n.fragment && n.fragment.d(t), (n.on_destroy = n.fragment = null), (n.ctx = []));
    }
    function He(e, t, l, o, c, i, a, u = [-1]) {
        const d = Z;
        le(e);
        const p = (e.$$ = {
            fragment: null,
            ctx: null,
            props: i,
            update: n,
            not_equal: c,
            bound: s(),
            on_mount: [],
            on_destroy: [],
            on_disconnect: [],
            before_update: [],
            after_update: [],
            context: new Map(d ? d.$$.context : t.context || []),
            callbacks: s(),
            dirty: u,
            skip_bound: !1,
            root: t.target || d.$$.root,
        });
        a && a(p.root);
        let f = !1;
        if (
            ((p.ctx = l
                ? l(e, t.props || {}, (t, n, ...l) => {
                    const o = l.length ? l[0] : n;
                    return (
                        p.ctx &&
                        c(p.ctx[t], (p.ctx[t] = o)) &&
                        (!p.skip_bound && p.bound[t] && p.bound[t](o),
                            f &&
                            (function (e, t) {
                                -1 === e.$$.dirty[0] && (de.push(e), ge(), e.$$.dirty.fill(0)), (e.$$.dirty[(t / 31) | 0] |= 1 << t % 31);
                            })(e, t)),
                        n
                    );
                })
                : []),
                p.update(),
                (f = !0),
                r(p.before_update),
                (p.fragment = !!o && o(p.ctx)),
                t.target)
        ) {
            if (t.hydrate) {
                const e = (function (e) {
                    return Array.from(e.childNodes);
                })(t.target);
                p.fragment && p.fragment.l(e), e.forEach(L);
            } else p.fragment && p.fragment.c();
            t.intro && Ae(e.$$.fragment), Ke(e, t.target, t.anchor, t.customElement), xe();
        }
        le(d);
    }
    class Ye {
        $destroy() {
            Ue(this, 1), (this.$destroy = n);
        }
        $on(e, t) {
            const n = this.$$.callbacks[e] || (this.$$.callbacks[e] = []);
            return (
                n.push(t),
                () => {
                    const e = n.indexOf(t);
                    -1 !== e && n.splice(e, 1);
                }
            );
        }
        $set(e) {
            var t;
            this.$$set && ((t = e), 0 !== Object.keys(t).length) && ((this.$$.skip_bound = !0), this.$$set(e), (this.$$.skip_bound = !1));
        }
    }
    function Be(e, t = "data-id") {
        let n = !e.tagName && e.target ? e.target : e;
        for (; n;) {
            if (n.getAttribute) {
                if (n.getAttribute(t)) return n;
            }
            n = n.parentNode;
        }
        return null;
    }
    function Ge(e) {
        const t = Be(e);
        return t ? 1 * t.getAttribute("data-id") : null;
    }
    function Je(e, t) {
        let n = null;
        e.addEventListener("click", function (e) {
            const l = Be(e);
            if (!l) return;
            let o,
                c = e.target;
            for (; c != l;) {
                if (((o = c.dataset ? c.dataset.action : null), o)) {
                    t[o] && t[o](1 * l.dataset.id, e), (n = new Date());
                    break;
                }
                c = c.parentNode;
            }
            t.click && !o && t.click(1 * l.dataset.id, e);
        }),
            e.addEventListener("dblclick", function (e) {
                if (n && new Date().valueOf() - n < 200) return;
                const l = Ge(e);
                l && t.dblclick && t.dblclick(l, e);
            });
    }
    function Ve(e) {
        return Xe(e.getMonth() + 1) + "/" + Xe(e.getDate()) + "/" + e.getFullYear();
    }
    function Xe(e, t) {
        let n = e.toString();
        return e < 10 && (n = "0" + n), t && e < 100 && (n = "0" + n), n;
    }
    function Qe(e, t) {
        return e.getMonth() === t.getMonth();
    }
    function We(e) {
        const t = 11 * Math.floor(e / 11);
        return { start: t, end: t + 11 };
    }
    let Ze = new Date().valueOf();
    function et() {
        return (Ze += 1), Ze;
    }
    function tt(e) {
        let t, l, o;
        return {
            c() {
                (t = z("textarea")), U(t, "id", e[1]), U(t, "placeholder", e[2]), U(t, "class", "svelte-smb02x");
            },
            m(n, c) {
                T(n, t, c), B(t, e[0]), l || ((o = R(t, "input", e[3])), (l = !0));
            },
            p(e, [n]) {
                2 & n && U(t, "id", e[1]), 4 & n && U(t, "placeholder", e[2]), 1 & n && B(t, e[0]);
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), (l = !1), o();
            },
        };
    }
    function nt(e, t, n) {
        let { value: l = "" } = t,
            { id: o = et() } = t,
            { placeholder: c = null } = t;
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (l = e.value)), "id" in e && n(1, (o = e.id)), "placeholder" in e && n(2, (c = e.placeholder));
            }),
            [
                l,
                o,
                c,
                function () {
                    (l = this.value), n(0, l);
                },
            ]
        );
    }
    function lt(e) {
        let t, n, l, o, c;
        const s = e[5].default,
            r = $(s, e, e[4], null);
        return {
            c() {
                (t = z("button")), r && r.c(), U(t, "class", (n = "btn " + e[2] + " " + e[0] + " svelte-1dai6kx")), V(t, "block", e[3]);
            },
            m(n, s) {
                T(n, t, s),
                    r && r.m(t, null),
                    (l = !0),
                    o ||
                    ((c = R(t, "click", function () {
                        i(e[1]) && e[1].apply(this, arguments);
                    })),
                        (o = !0));
            },
            p(o, [c]) {
                (e = o),
                    r && r.p && (!l || 16 & c) && g(r, s, e, e[4], l ? h(s, e[4], c, null) : v(e[4]), null),
                    (!l || (5 & c && n !== (n = "btn " + e[2] + " " + e[0] + " svelte-1dai6kx"))) && U(t, "class", n),
                    13 & c && V(t, "block", e[3]);
            },
            i(e) {
                l || (Ae(r, e), (l = !0));
            },
            o(e) {
                Ee(r, e), (l = !1);
            },
            d(e) {
                e && L(t), r && r.d(e), (o = !1), c();
            },
        };
    }
    function ot(e, t, n) {
        let { $$slots: l = {}, $$scope: o } = t,
            { type: c = "" } = t,
            { click: s } = t,
            { shape: r = "round" } = t,
            { block: i = !1 } = t;
        return (
            (e.$$set = (e) => {
                "type" in e && n(0, (c = e.type)), "click" in e && n(1, (s = e.click)), "shape" in e && n(2, (r = e.shape)), "block" in e && n(3, (i = e.block)), "$$scope" in e && n(4, (o = e.$$scope));
            }),
            [c, s, r, i, o, l]
        );
    }
    class ct extends Ye {
        constructor(e) {
            super(), He(this, e, ot, lt, a, { type: 0, click: 1, shape: 2, block: 3 });
        }
    }
    const { document: st } = je;
    function rt(e) {
        let t, n, l, o, c, s;
        const r = e[9].default,
            i = $(r, e, e[8], null);
        return {
            c() {
                (t = F()), (n = z("div")), i && i.c(), U(n, "class", (l = "popup " + e[2] + " svelte-1ykyzo8")), G(n, "width", e[0]);
            },
            m(l, r) {
                T(l, t, r), T(l, n, r), i && i.m(n, null), e[10](n), (o = !0), c || ((s = R(st.body, "mousedown", e[3])), (c = !0));
            },
            p(e, [t]) {
                i && i.p && (!o || 256 & t) && g(i, r, e, e[8], o ? h(r, e[8], t, null) : v(e[8]), null), (!o || (4 & t && l !== (l = "popup " + e[2] + " svelte-1ykyzo8"))) && U(n, "class", l), (!o || 1 & t) && G(n, "width", e[0]);
            },
            i(e) {
                o || (Ae(i, e), (o = !0));
            },
            o(e) {
                Ee(i, e), (o = !1);
            },
            d(l) {
                l && L(t), l && L(n), i && i.d(l), e[10](null), (c = !1), s();
            },
        };
    }
    function it(e, t, n) {
        let l,
            { $$slots: o = {}, $$scope: c } = t,
            { position: s = "bottom" } = t,
            { align: r = "start" } = t,
            { autoFit: i = !0 } = t,
            { cancel: a = null } = t,
            { width: u = "100%" } = t,
            d = "";
        return (
            se(() => {
                if (i) {
                    const e = l.getBoundingClientRect(),
                        t = document.body.getBoundingClientRect();
                    e.right >= t.right && (n(4, (s = "bottom")), n(5, (r = "end"))), e.bottom >= t.bottom && (n(4, (s = "top")), n(5, (r = "end")));
                }
                n(2, (d = r ? `${s}-${r}` : `${s}`));
            }),
            (e.$$set = (e) => {
                "position" in e && n(4, (s = e.position)),
                    "align" in e && n(5, (r = e.align)),
                    "autoFit" in e && n(6, (i = e.autoFit)),
                    "cancel" in e && n(7, (a = e.cancel)),
                    "width" in e && n(0, (u = e.width)),
                    "$$scope" in e && n(8, (c = e.$$scope));
            }),
            [
                u,
                l,
                d,
                function (e) {
                    l.contains(e.target) || (a && a(e));
                },
                s,
                r,
                i,
                a,
                c,
                o,
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (l = e), n(1, l);
                    });
                },
            ]
        );
    }
    class at extends Ye {
        constructor(e) {
            super(), He(this, e, it, rt, a, { position: 4, align: 5, autoFit: 6, cancel: 7, width: 0 });
        }
    }
    function ut(e, t, n) {
        const l = e.slice();
        return (l[11] = t[n]), l;
    }
    const dt = (e) => ({ option: 8 & e }),
        pt = (e) => ({ option: e[11] });
    function ft(e) {
        let t;
        return {
            c() {
                t = O(e[2]);
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, n) {
                4 & n && Y(t, e[2]);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function $t(e) {
        let t, n;
        return (
            (t = new at({ props: { cancel: e[9], $$slots: { default: [ht] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    16 & n && (l.cancel = e[9]), 1032 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function mt(e, t) {
        let n, l, o, c;
        const s = t[7].default,
            r = $(s, t, t[10], pt);
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), r && r.c(), (l = F()), U(n, "class", "wx-list-item"), U(n, "data-id", (o = t[11].id)), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), r && r.m(n, null), I(n, l), (c = !0);
            },
            p(e, l) {
                (t = e), r && r.p && (!c || 1032 & l) && g(r, s, t, t[10], c ? h(s, t[10], l, dt) : v(t[10]), pt), (!c || (8 & l && o !== (o = t[11].id))) && U(n, "data-id", o);
            },
            i(e) {
                c || (Ae(r, e), (c = !0));
            },
            o(e) {
                Ee(r, e), (c = !1);
            },
            d(e) {
                e && L(n), r && r.d(e);
            },
        };
    }
    function ht(e) {
        let t,
            n,
            l,
            o,
            c = [],
            s = new Map(),
            r = e[3];
        const i = (e) => e[11].id;
        for (let t = 0; t < r.length; t += 1) {
            let n = ut(e, r, t),
                l = i(n);
            s.set(l, (c[t] = mt(l, n)));
        }
        return {
            c() {
                t = z("div");
                for (let e = 0; e < c.length; e += 1) c[e].c();
                U(t, "class", "wx-list list");
            },
            m(s, r) {
                T(s, t, r);
                for (let e = 0; e < c.length; e += 1) c[e].m(t, null);
                (n = !0), l || ((o = k(Je.call(null, t, e[5]))), (l = !0));
            },
            p(e, n) {
                1032 & n && ((r = e[3]), De(), (c = Oe(c, n, i, 1, e, r, s, t, Ne, mt, null, ut)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < r.length; e += 1) Ae(c[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < c.length; e += 1) Ee(c[e]);
                n = !1;
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < c.length; e += 1) c[e].d();
                (l = !1), o();
            },
        };
    }
    function gt(e) {
        let t, n, l, o;
        n = new ct({ props: { type: e[0], shape: e[1], click: e[8], $$slots: { default: [ft] }, $$scope: { ctx: e } } });
        let c = e[4] && $t(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (l = F()), c && c.c(), U(t, "class", "layout svelte-5bx2eh");
            },
            m(e, s) {
                T(e, t, s), Ke(n, t, null), I(t, l), c && c.m(t, null), (o = !0);
            },
            p(e, [l]) {
                const o = {};
                1 & l && (o.type = e[0]),
                    2 & l && (o.shape = e[1]),
                    16 & l && (o.click = e[8]),
                    1028 & l && (o.$$scope = { dirty: l, ctx: e }),
                    n.$set(o),
                    e[4]
                        ? c
                            ? (c.p(e, l), 16 & l && Ae(c, 1))
                            : ((c = $t(e)), c.c(), Ae(c, 1), c.m(t, null))
                        : c &&
                        (De(),
                            Ee(c, 1, 1, () => {
                                c = null;
                            }),
                            Ie());
            },
            i(e) {
                o || (Ae(n.$$.fragment, e), Ae(c), (o = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), Ee(c), (o = !1);
            },
            d(e) {
                e && L(t), Ue(n), c && c.d();
            },
        };
    }
    function vt(e, t, n) {
        let { $$slots: l = {}, $$scope: o } = t,
            { type: c = "" } = t,
            { shape: s = "round" } = t,
            { label: r = "" } = t,
            { click: i } = t,
            { options: a = [] } = t,
            u = null;
        const d = {
            click: (e) => {
                n(4, (u = null)), i(e);
            },
        };
        return (
            (e.$$set = (e) => {
                "type" in e && n(0, (c = e.type)),
                    "shape" in e && n(1, (s = e.shape)),
                    "label" in e && n(2, (r = e.label)),
                    "click" in e && n(6, (i = e.click)),
                    "options" in e && n(3, (a = e.options)),
                    "$$scope" in e && n(10, (o = e.$$scope));
            }),
            [c, s, r, a, u, d, i, l, () => n(4, (u = !0)), () => n(4, (u = null)), o]
        );
    }
    function yt(e) {
        let t, l, o, c, s, i, a;
        return {
            c() {
                (t = z("div")),
                    (l = z("input")),
                    (o = F()),
                    (c = z("label")),
                    (s = O(e[2])),
                    U(l, "type", "checkbox"),
                    U(l, "id", e[1]),
                    U(l, "class", "svelte-1wrup4b"),
                    U(c, "for", e[1]),
                    U(c, "class", "svelte-1wrup4b"),
                    U(t, "class", "svelte-1wrup4b");
            },
            m(n, r) {
                T(n, t, r), I(t, l), (l.checked = e[0]), I(t, o), I(t, c), I(c, s), i || ((a = [R(l, "change", e[4]), R(l, "change", e[5])]), (i = !0));
            },
            p(e, [t]) {
                2 & t && U(l, "id", e[1]), 1 & t && (l.checked = e[0]), 4 & t && Y(s, e[2]), 2 & t && U(c, "for", e[1]);
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), (i = !1), r(a);
            },
        };
    }
    function wt(e, t, n) {
        const l = re();
        let { id: o = et() } = t,
            { label: c = "" } = t,
            { value: s = "" } = t;
        return (
            (e.$$set = (e) => {
                "id" in e && n(1, (o = e.id)), "label" in e && n(2, (c = e.label)), "value" in e && n(0, (s = e.value));
            }),
            [
                s,
                o,
                c,
                l,
                function () {
                    (s = this.checked), n(0, s);
                },
                () => l("change", { value: s }),
            ]
        );
    }
    class bt extends Ye {
        constructor(e) {
            super(), He(this, e, wt, yt, a, { id: 1, label: 2, value: 0 });
        }
    }
    function xt(e, t, n) {
        const l = e.slice();
        return (l[12] = t[n]), l;
    }
    function kt(e) {
        let t;
        return {
            c() {
                (t = z("div")), U(t, "class", "empty selected svelte-y41d73");
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function St(e) {
        let t;
        return {
            c() {
                (t = z("div")), U(t, "class", "color selected svelte-y41d73"), G(t, "background-color", e[0] || "#00a037");
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, n) {
                1 & n && G(t, "background-color", e[0] || "#00a037");
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Mt(e) {
        let t, l, o;
        return {
            c() {
                (t = z("i")), U(t, "class", "clear wxi-close svelte-y41d73");
            },
            m(n, c) {
                T(n, t, c), l || ((o = R(t, "click", K(e[7]))), (l = !0));
            },
            p: n,
            d(e) {
                e && L(t), (l = !1), o();
            },
        };
    }
    function _t(e) {
        let t, n;
        return (
            (t = new at({ props: { cancel: e[10], width: "auto", $$slots: { default: [Dt] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    32 & n && (l.cancel = e[10]), 32770 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Ct(e) {
        let t, n, l;
        function o() {
            return e[9](e[12]);
        }
        return {
            c() {
                (t = z("div")), U(t, "class", "color svelte-y41d73"), G(t, "background-color", e[12]);
            },
            m(e, c) {
                T(e, t, c), n || ((l = R(t, "click", K(o))), (n = !0));
            },
            p(n, l) {
                (e = n), 2 & l && G(t, "background-color", e[12]);
            },
            d(e) {
                e && L(t), (n = !1), l();
            },
        };
    }
    function Dt(e) {
        let t,
            n,
            l,
            o,
            c,
            s = e[1],
            r = [];
        for (let t = 0; t < s.length; t += 1) r[t] = Ct(xt(e, s, t));
        return {
            c() {
                (t = z("div")), (n = z("div")), (l = F());
                for (let e = 0; e < r.length; e += 1) r[e].c();
                U(n, "class", "empty svelte-y41d73"), U(t, "class", "colors svelte-y41d73");
            },
            m(s, i) {
                T(s, t, i), I(t, n), I(t, l);
                for (let e = 0; e < r.length; e += 1) r[e].m(t, null);
                o || ((c = R(n, "click", K(e[8]))), (o = !0));
            },
            p(e, n) {
                if (66 & n) {
                    let l;
                    for (s = e[1], l = 0; l < s.length; l += 1) {
                        const o = xt(e, s, l);
                        r[l] ? r[l].p(o, n) : ((r[l] = Ct(o)), r[l].c(), r[l].m(t, null));
                    }
                    for (; l < r.length; l += 1) r[l].d(1);
                    r.length = s.length;
                }
            },
            d(e) {
                e && L(t), j(r, e), (o = !1), c();
            },
        };
    }
    function It(e) {
        let t, n, l, o, c, s, r, i;
        function a(e, t) {
            return e[0] ? St : kt;
        }
        let u = a(e),
            d = u(e),
            p = e[3] && e[0] && Mt(e),
            f = e[5] && _t(e);
        return {
            c() {
                (t = z("div")),
                    d.c(),
                    (n = F()),
                    (l = z("input")),
                    (o = F()),
                    p && p.c(),
                    (c = F()),
                    f && f.c(),
                    (l.value = e[0]),
                    (l.readOnly = !0),
                    U(l, "id", e[2]),
                    U(l, "placeholder", e[4]),
                    U(l, "class", "svelte-y41d73"),
                    U(t, "class", "layout svelte-y41d73");
            },
            m(a, u) {
                T(a, t, u), d.m(t, null), I(t, n), I(t, l), I(t, o), p && p.m(t, null), I(t, c), f && f.m(t, null), (s = !0), r || ((i = R(t, "click", e[11])), (r = !0));
            },
            p(e, [o]) {
                u === (u = a(e)) && d ? d.p(e, o) : (d.d(1), (d = u(e)), d && (d.c(), d.m(t, n))),
                    (!s || (1 & o && l.value !== e[0])) && (l.value = e[0]),
                    (!s || 4 & o) && U(l, "id", e[2]),
                    (!s || 16 & o) && U(l, "placeholder", e[4]),
                    e[3] && e[0] ? (p ? p.p(e, o) : ((p = Mt(e)), p.c(), p.m(t, c))) : p && (p.d(1), (p = null)),
                    e[5]
                        ? f
                            ? (f.p(e, o), 32 & o && Ae(f, 1))
                            : ((f = _t(e)), f.c(), Ae(f, 1), f.m(t, null))
                        : f &&
                        (De(),
                            Ee(f, 1, 1, () => {
                                f = null;
                            }),
                            Ie());
            },
            i(e) {
                s || (Ae(f), (s = !0));
            },
            o(e) {
                Ee(f), (s = !1);
            },
            d(e) {
                e && L(t), d.d(), p && p.d(), f && f.d(), (r = !1), i();
            },
        };
    }
    function At(e, t, n) {
        let l,
            { colors: o = ["#00a037", "#df282f", "#fd772c", "#6d4bce", "#b26bd3", "#c87095", "#90564d", "#eb2f89", "#ea77c0", "#777676", "#a9a8a8", "#9bb402", "#e7a90b", "#0bbed7", "#038cd9"] } = t,
            { value: c = "" } = t,
            { id: s = et() } = t,
            { clear: r = !0 } = t,
            { placeholder: i = "" } = t;
        function a(e) {
            n(0, (c = e)), n(5, (l = null));
        }
        return (
            (e.$$set = (e) => {
                "colors" in e && n(1, (o = e.colors)), "value" in e && n(0, (c = e.value)), "id" in e && n(2, (s = e.id)), "clear" in e && n(3, (r = e.clear)), "placeholder" in e && n(4, (i = e.placeholder));
            }),
            [
                c,
                o,
                s,
                r,
                i,
                l,
                a,
                function () {
                    n(0, (c = null));
                },
                () => a(""),
                (e) => a(e),
                () => n(5, (l = null)),
                () => n(5, (l = !0)),
            ]
        );
    }
    function Et(e, t, n) {
        const l = e.slice();
        return (l[24] = t[n]), (l[25] = t), (l[26] = n), l;
    }
    const Tt = (e) => ({ option: 32 & e }),
        Lt = (e) => ({ option: e[24] });
    function jt(e) {
        let t, n;
        return (
            (t = new at({ props: { cancel: e[7], $$slots: { default: [Ft] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    262189 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function zt(e) {
        let t;
        return {
            c() {
                (t = z("div")), (t.textContent = "No data"), U(t, "class", "no-data svelte-qe4iw2");
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            i: n,
            o: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function Nt(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[5];
        const s = (e) => e[24].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = Et(e, c, t),
                r = s(n);
            o.set(r, (l[t] = Ot(r, n)));
        }
        return {
            c() {
                for (let e = 0; e < l.length; e += 1) l[e].c();
                t = q();
            },
            m(e, o) {
                for (let t = 0; t < l.length; t += 1) l[t].m(e, o);
                T(e, t, o), (n = !0);
            },
            p(e, n) {
                262189 & n && ((c = e[5]), De(), (l = Oe(l, n, s, 1, e, c, o, t.parentNode, Ne, Ot, t, Et)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                for (let t = 0; t < l.length; t += 1) l[t].d(e);
                e && L(t);
            },
        };
    }
    function Ot(e, t) {
        let n,
            l,
            o,
            c,
            s = t[26];
        const r = t[15].default,
            i = $(r, t, t[18], Lt),
            a = () => t[17](n, s),
            u = () => t[17](null, s);
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), i && i.c(), (l = F()), U(n, "class", "item svelte-qe4iw2"), U(n, "data-id", (o = t[24].id)), V(n, "selected", t[0] && t[0] === t[24].id), V(n, "navigate", t[3] && t[3].id === t[24].id), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), i && i.m(n, null), I(n, l), a(), (c = !0);
            },
            p(e, l) {
                (t = e),
                    i && i.p && (!c || 262176 & l) && g(i, r, t, t[18], c ? h(r, t[18], l, Tt) : v(t[18]), Lt),
                    (!c || (32 & l && o !== (o = t[24].id))) && U(n, "data-id", o),
                    s !== t[26] && (u(), (s = t[26]), a()),
                    33 & l && V(n, "selected", t[0] && t[0] === t[24].id),
                    40 & l && V(n, "navigate", t[3] && t[3].id === t[24].id);
            },
            i(e) {
                c || (Ae(i, e), (c = !0));
            },
            o(e) {
                Ee(i, e), (c = !1);
            },
            d(e) {
                e && L(n), i && i.d(e), u();
            },
        };
    }
    function Ft(e) {
        let t, n, l, o, c, s;
        const i = [Nt, zt],
            a = [];
        function u(e, t) {
            return e[5].length ? 0 : 1;
        }
        return (
            (n = u(e)),
            (l = a[n] = i[n](e)),
            {
                c() {
                    (t = z("div")), l.c(), U(t, "class", "list svelte-qe4iw2");
                },
                m(l, r) {
                    T(l, t, r), a[n].m(t, null), (o = !0), c || ((s = [R(t, "click", e[9]), R(t, "mousemove", e[10])]), (c = !0));
                },
                p(e, o) {
                    let c = n;
                    (n = u(e)),
                        n === c
                            ? a[n].p(e, o)
                            : (De(),
                                Ee(a[c], 1, 1, () => {
                                    a[c] = null;
                                }),
                                Ie(),
                                (l = a[n]),
                                l ? l.p(e, o) : ((l = a[n] = i[n](e)), l.c()),
                                Ae(l, 1),
                                l.m(t, null));
                },
                i(e) {
                    o || (Ae(l), (o = !0));
                },
                o(e) {
                    Ee(l), (o = !1);
                },
                d(e) {
                    e && L(t), a[n].d(), (c = !1), r(s);
                },
            }
        );
    }
    function qt(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            i,
            a,
            u = e[1] && jt(e);
        return {
            c() {
                (t = z("div")),
                    (n = z("input")),
                    (l = F()),
                    (o = z("i")),
                    (c = F()),
                    u && u.c(),
                    U(n, "class", "input svelte-qe4iw2"),
                    U(n, "tabindex", "0"),
                    V(n, "active", e[1]),
                    U(o, "class", "icon wxi-angle-down svelte-qe4iw2"),
                    U(t, "class", "layout svelte-qe4iw2");
            },
            m(r, d) {
                T(r, t, d), I(t, n), B(n, e[4]), I(t, l), I(t, o), I(t, c), u && u.m(t, null), (s = !0), i || ((a = [R(n, "input", e[16]), R(n, "click", e[6]), R(n, "input", e[8]), R(n, "keydown", e[11])]), (i = !0));
            },
            p(e, [l]) {
                16 & l && n.value !== e[4] && B(n, e[4]),
                    2 & l && V(n, "active", e[1]),
                    e[1]
                        ? u
                            ? (u.p(e, l), 2 & l && Ae(u, 1))
                            : ((u = jt(e)), u.c(), Ae(u, 1), u.m(t, null))
                        : u &&
                        (De(),
                            Ee(u, 1, 1, () => {
                                u = null;
                            }),
                            Ie());
            },
            i(e) {
                s || (Ae(u), (s = !0));
            },
            o(e) {
                Ee(u), (s = !1);
            },
            d(e) {
                e && L(t), u && u.d(), (i = !1), r(a);
            },
        };
    }
    function Rt(e, t, n) {
        let l,
            { $$slots: o = {}, $$scope: c } = t,
            { value: s } = t,
            { options: r = [] } = t,
            { key: i = "label" } = t;
        const a = re();
        let u,
            d,
            p,
            f = [],
            $ = "",
            m = "";
        function h() {
            n(1, (u = !0)),
                (p = s ? l.findIndex((e) => e.id === s) : 0),
                n(3, (d = l[p])),
                (ge(), me).then(() => {
                    f[p] && f[p].scrollIntoView({ block: "nearest" });
                });
        }
        function g() {
            n(14, ($ = "")), n(1, (u = null)), (p = null), n(3, (d = null));
        }
        function v(e) {
            (p += e), p > l.length - 1 ? (p = l.length - 1) : p < 0 && (p = 0), l.length && (n(3, (d = l[p])), f[p].scrollIntoView({ block: "nearest" }));
        }
        return (
            ce(() => {
                s && n(4, (m = r.find((e) => e.id === s)[i] || ""));
            }),
            (e.$$set = (e) => {
                "value" in e && n(0, (s = e.value)), "options" in e && n(12, (r = e.options)), "key" in e && n(13, (i = e.key)), "$$scope" in e && n(18, (c = e.$$scope));
            }),
            (e.$$.update = () => {
                28672 & e.$$.dirty && n(5, (l = $ ? r.filter((e) => e[i].toLowerCase().includes($.toLowerCase())) : [].concat(r)));
            }),
            [
                s,
                u,
                f,
                d,
                m,
                l,
                h,
                function () {
                    l.length ? (n(0, (s = l[0].id)), n(4, (m = l[0][i]))) : (n(0, (s = null)), n(4, (m = ""))), g();
                },
                function () {
                    u || h(), n(14, ($ = m || ""));
                },
                function (e) {
                    const t = Ge(e);
                    t && (n(0, (s = t)), n(4, (m = r.find((e) => e.id === t)[i])), a("change", { value: s }), g());
                },
                function (e) {
                    const t = Ge(e);
                    t && ((p = l.findIndex((e) => e.id === t)), n(3, (d = l[p])));
                },
                function (e) {
                    switch (e.code) {
                        case "Space":
                            u ? g() : h();
                            break;
                        case "Tab":
                            u && g();
                            break;
                        case "Enter":
                            u
                                ? (function () {
                                    const e = d.id;
                                    l.length ? (n(0, (s = e)), n(4, (m = r.find((t) => t.id === e)[i])), a("change", { value: s })) : (n(0, (s = null)), n(4, (m = ""))), g();
                                })()
                                : h();
                            break;
                        case "ArrowDown":
                            u ? v(1) : h();
                            break;
                        case "ArrowUp":
                            u ? v(-1) : h();
                            break;
                        case "Escape":
                            g();
                    }
                },
                r,
                i,
                $,
                o,
                function () {
                    (m = this.value), n(4, m);
                },
                function (e, t) {
                    pe[e ? "unshift" : "push"](() => {
                        (f[t] = e), n(2, f), n(5, l), n(14, $), n(12, r), n(13, i);
                    });
                },
                c,
            ]
        );
    }
    function Pt(e) {
        let t, n, l;
        return {
            c() {
                (t = z("input")), U(t, "id", e[1]), (t.readOnly = e[2]), (t.disabled = e[5]), U(t, "placeholder", e[4]), U(t, "style", e[6]), U(t, "class", "svelte-66mv3w");
            },
            m(o, c) {
                T(o, t, c), B(t, e[0]), e[17](t), n || ((l = [R(t, "input", e[16]), R(t, "input", e[18])]), (n = !0));
            },
            p(e, n) {
                2 & n && U(t, "id", e[1]), 4 & n && (t.readOnly = e[2]), 32 & n && (t.disabled = e[5]), 16 & n && U(t, "placeholder", e[4]), 64 & n && U(t, "style", e[6]), 1 & n && t.value !== e[0] && B(t, e[0]);
            },
            d(o) {
                o && L(t), e[17](null), (n = !1), r(l);
            },
        };
    }
    function Kt(e) {
        let t, n, l;
        return {
            c() {
                (t = z("input")), U(t, "id", e[1]), (t.readOnly = e[2]), (t.disabled = e[5]), U(t, "placeholder", e[4]), U(t, "type", "number"), U(t, "style", e[6]), U(t, "class", "svelte-66mv3w");
            },
            m(o, c) {
                T(o, t, c), B(t, e[0]), e[14](t), n || ((l = [R(t, "input", e[13]), R(t, "input", e[15])]), (n = !0));
            },
            p(e, n) {
                2 & n && U(t, "id", e[1]), 4 & n && (t.readOnly = e[2]), 32 & n && (t.disabled = e[5]), 16 & n && U(t, "placeholder", e[4]), 64 & n && U(t, "style", e[6]), 1 & n && H(t.value) !== e[0] && B(t, e[0]);
            },
            d(o) {
                o && L(t), e[14](null), (n = !1), r(l);
            },
        };
    }
    function Ut(e) {
        let t, n, l;
        return {
            c() {
                (t = z("input")), U(t, "id", e[1]), (t.readOnly = e[2]), (t.disabled = e[5]), U(t, "placeholder", e[4]), U(t, "type", "password"), U(t, "style", e[6]), U(t, "class", "svelte-66mv3w");
            },
            m(o, c) {
                T(o, t, c), B(t, e[0]), e[11](t), n || ((l = [R(t, "input", e[10]), R(t, "input", e[12])]), (n = !0));
            },
            p(e, n) {
                2 & n && U(t, "id", e[1]), 4 & n && (t.readOnly = e[2]), 32 & n && (t.disabled = e[5]), 16 & n && U(t, "placeholder", e[4]), 64 & n && U(t, "style", e[6]), 1 & n && t.value !== e[0] && B(t, e[0]);
            },
            d(o) {
                o && L(t), e[11](null), (n = !1), r(l);
            },
        };
    }
    function Ht(e) {
        let t;
        function l(e, t) {
            return "password" == e[3] ? Ut : "number" == e[3] ? Kt : Pt;
        }
        let o = l(e),
            c = o(e);
        return {
            c() {
                c.c(), (t = q());
            },
            m(e, n) {
                c.m(e, n), T(e, t, n);
            },
            p(e, [n]) {
                o === (o = l(e)) && c ? c.p(e, n) : (c.d(1), (c = o(e)), c && (c.c(), c.m(t.parentNode, t)));
            },
            i: n,
            o: n,
            d(e) {
                c.d(e), e && L(t);
            },
        };
    }
    function Yt(e, t, n) {
        let { value: l = "" } = t,
            { id: o = et() } = t,
            { readonly: c = !1 } = t,
            { focus: s = !1 } = t,
            { type: r = "text" } = t,
            { placeholder: i = null } = t,
            { disabled: a = !1 } = t,
            { inputStyle: u = "" } = t;
        const d = re();
        let p;
        s && ce(() => p.focus());
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (l = e.value)),
                    "id" in e && n(1, (o = e.id)),
                    "readonly" in e && n(2, (c = e.readonly)),
                    "focus" in e && n(9, (s = e.focus)),
                    "type" in e && n(3, (r = e.type)),
                    "placeholder" in e && n(4, (i = e.placeholder)),
                    "disabled" in e && n(5, (a = e.disabled)),
                    "inputStyle" in e && n(6, (u = e.inputStyle));
            }),
            [
                l,
                o,
                c,
                r,
                i,
                a,
                u,
                p,
                d,
                s,
                function () {
                    (l = this.value), n(0, l);
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (p = e), n(7, p);
                    });
                },
                () => d("input", { value: l }),
                function () {
                    (l = H(this.value)), n(0, l);
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (p = e), n(7, p);
                    });
                },
                () => d("input", { value: l }),
                function () {
                    (l = this.value), n(0, l);
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (p = e), n(7, p);
                    });
                },
                () => d("input", { value: l }),
            ]
        );
    }
    class Bt extends Ye {
        constructor(e) {
            super(), He(this, e, Yt, Ht, a, { value: 0, id: 1, readonly: 2, focus: 9, type: 3, placeholder: 4, disabled: 5, inputStyle: 6 });
        }
    }
    const Gt = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
        Jt = {
            __dates: {
                months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
                monthsShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
                days: Gt,
            },
            wx: { Today: "Today", Clear: "Clear", Close: "Close" },
        };
    function Vt(e) {
        let t = e;
        return {
            _: (e) => t[e] || e,
            __(e, n) {
                const l = t[e];
                return (l && l[n]) || n;
            },
            getGroup(e) {
                return (t) => this.__(e, t);
            },
            extend(e) {
                return (t = { ...t, ...e }), this;
            },
        };
    }
    function Xt(e) {
        let t;
        return {
            c() {
                (t = z("span")), U(t, "class", "spacer svelte-1qykuws");
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function Qt(e) {
        let t, n, l;
        return {
            c() {
                (t = z("i")), U(t, "class", "pager wxi-angle-left svelte-1qykuws");
            },
            m(o, c) {
                T(o, t, c),
                    n ||
                    ((l = R(t, "click", function () {
                        i(e[0]) && e[0].apply(this, arguments);
                    })),
                        (n = !0));
            },
            p(t, n) {
                e = t;
            },
            d(e) {
                e && L(t), (n = !1), l();
            },
        };
    }
    function Wt(e) {
        let t;
        return {
            c() {
                (t = z("span")), U(t, "class", "spacer svelte-1qykuws");
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function Zt(e) {
        let t, n, l;
        return {
            c() {
                (t = z("i")), U(t, "class", "pager wxi-angle-right svelte-1qykuws");
            },
            m(o, c) {
                T(o, t, c),
                    n ||
                    ((l = R(t, "click", function () {
                        i(e[1]) && e[1].apply(this, arguments);
                    })),
                        (n = !0));
            },
            p(t, n) {
                e = t;
            },
            d(e) {
                e && L(t), (n = !1), l();
            },
        };
    }
    function en(e) {
        let t, l, o, c, s, r, i;
        function a(e, t) {
            return "right" != e[2] ? Qt : Xt;
        }
        let u = a(e),
            d = u(e);
        function p(e, t) {
            return "left" != e[2] ? Zt : Wt;
        }
        let f = p(e),
            $ = f(e);
        return {
            c() {
                (t = z("div")), d.c(), (l = F()), (o = z("span")), (c = O(e[3])), (s = F()), $.c(), U(o, "class", "header-label svelte-1qykuws"), U(t, "class", "header svelte-1qykuws");
            },
            m(n, a) {
                T(n, t, a), d.m(t, null), I(t, l), I(t, o), I(o, c), I(t, s), $.m(t, null), r || ((i = R(o, "click", e[4])), (r = !0));
            },
            p(e, [n]) {
                u === (u = a(e)) && d ? d.p(e, n) : (d.d(1), (d = u(e)), d && (d.c(), d.m(t, l))), 8 & n && Y(c, e[3]), f === (f = p(e)) && $ ? $.p(e, n) : ($.d(1), ($ = f(e)), $ && ($.c(), $.m(t, null)));
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), d.d(), $.d(), (r = !1), i();
            },
        };
    }
    function tn(e, t, n) {
        const l = (ae("wx-i18n") || Vt(Jt)).getGroup("__dates");
        let o,
            c,
            s,
            { date: r } = t,
            { type: i } = t,
            { prev: a } = t,
            { next: u } = t,
            { part: d } = t;
        const p = l("months");
        return (
            (e.$$set = (e) => {
                "date" in e && n(6, (r = e.date)), "type" in e && n(5, (i = e.type)), "prev" in e && n(0, (a = e.prev)), "next" in e && n(1, (u = e.next)), "part" in e && n(2, (d = e.part));
            }),
            (e.$$.update = () => {
                if (480 & e.$$.dirty)
                    switch ((n(7, (o = r.getMonth())), n(8, (c = r.getFullYear())), i)) {
                        case "month":
                            n(3, (s = `${p[o]} ${c}`));
                            break;
                        case "year":
                            n(3, (s = c));
                            break;
                        case "duodecade": {
                            const { start: e, end: t } = We(c);
                            n(3, (s = `${e} - ${t}`));
                            break;
                        }
                    }
            }),
            [
                a,
                u,
                d,
                s,
                function () {
                    "month" === i ? n(5, (i = "year")) : "year" === i && n(5, (i = "duodecade"));
                },
                i,
                r,
                o,
                c,
            ]
        );
    }
    class nn extends Ye {
        constructor(e) {
            super(), He(this, e, tn, en, a, { date: 6, type: 5, prev: 0, next: 1, part: 2 });
        }
    }
    function ln(e, t, n) {
        const l = e.slice();
        return (l[13] = t[n]), l;
    }
    function on(e, t, n) {
        const l = e.slice();
        return (l[13] = t[n]), l;
    }
    function cn(e) {
        let t,
            l,
            o = e[13] + "";
        return {
            c() {
                (t = z("div")), (l = O(o)), U(t, "class", "weekday svelte-oagymx");
            },
            m(e, n) {
                T(e, t, n), I(t, l);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function sn(e, t) {
        let n,
            l,
            o,
            c,
            s,
            r = t[13].day + "";
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), (l = O(r)), (o = F()), U(n, "class", (c = "day " + t[13].css + " svelte-oagymx")), U(n, "data-id", (s = t[13].date)), V(n, "out", !t[13].in), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), I(n, l), I(n, o);
            },
            p(e, o) {
                (t = e), 1 & o && r !== (r = t[13].day + "") && Y(l, r), 1 & o && c !== (c = "day " + t[13].css + " svelte-oagymx") && U(n, "class", c), 1 & o && s !== (s = t[13].date) && U(n, "data-id", s), 1 & o && V(n, "out", !t[13].in);
            },
            d(e) {
                e && L(n);
            },
        };
    }
    function rn(e) {
        let t,
            l,
            o,
            c,
            s,
            r = [],
            i = new Map(),
            a = e[1],
            u = [];
        for (let t = 0; t < a.length; t += 1) u[t] = cn(on(e, a, t));
        let d = e[0];
        const p = (e) => e[13].date;
        for (let t = 0; t < d.length; t += 1) {
            let n = ln(e, d, t),
                l = p(n);
            i.set(l, (r[t] = sn(l, n)));
        }
        return {
            c() {
                t = z("div");
                for (let e = 0; e < u.length; e += 1) u[e].c();
                (l = F()), (o = z("div"));
                for (let e = 0; e < r.length; e += 1) r[e].c();
                U(t, "class", "weekdays svelte-oagymx"), U(o, "class", "days svelte-oagymx");
            },
            m(n, i) {
                T(n, t, i);
                for (let e = 0; e < u.length; e += 1) u[e].m(t, null);
                T(n, l, i), T(n, o, i);
                for (let e = 0; e < r.length; e += 1) r[e].m(o, null);
                c || ((s = k(Je.call(null, o, e[2]))), (c = !0));
            },
            p(e, [n]) {
                if (2 & n) {
                    let l;
                    for (a = e[1], l = 0; l < a.length; l += 1) {
                        const o = on(e, a, l);
                        u[l] ? u[l].p(o, n) : ((u[l] = cn(o)), u[l].c(), u[l].m(t, null));
                    }
                    for (; l < u.length; l += 1) u[l].d(1);
                    u.length = a.length;
                }
                1 & n && ((d = e[0]), (r = Oe(r, n, p, 1, e, d, i, o, ze, sn, null, ln)));
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), j(u, e), e && L(l), e && L(o);
                for (let e = 0; e < r.length; e += 1) r[e].d();
                (c = !1), s();
            },
        };
    }
    function an(e) {
        const t = e.getDay();
        return 0 === t || 6 === t;
    }
    function un(e, t, n) {
        let { value: l } = t,
            { current: o } = t,
            { cancel: c } = t,
            { part: s } = t;
        const r = ae("wx-i18n"),
            i = r ? r.__("__dates", "days") : Gt;
        let a,
            u,
            d = "normal" !== s;
        const p = {
            click: function (e, t) {
                t.stopPropagation(),
                    e
                        ? (n(4, (o = new Date(e))),
                            "normal" === s
                                ? n(3, (l = new Date(o)))
                                : (l || n(3, (l = { start: null, end: null })),
                                    l.end || !l.start ? n(3, (l = { start: new Date(o), end: null })) : (n(3, (l.end = new Date(o)), l), l.end < l.start && n(3, ([l.start, l.end] = [l.end, l.start]), l))))
                        : (n(3, (l = null)), n(4, (o = new Date())));
                "normal" === s && c();
            },
        };
        return (
            (e.$$set = (e) => {
                "value" in e && n(3, (l = e.value)), "current" in e && n(4, (o = e.current)), "cancel" in e && n(5, (c = e.cancel)), "part" in e && n(6, (s = e.part));
            }),
            (e.$$.update = () => {
                if (217 & e.$$.dirty) {
                    n(7, (u = "normal" == s ? [l ? l.valueOf() : 0] : l ? [l.start ? l.start.valueOf() : 0, l.end ? l.end.valueOf() : 0] : [0, 0]));
                    const e = (function () {
                        const e = new Date(o);
                        e.setDate(1);
                        const t = new Date(e);
                        return t.setDate(t.getDate() - t.getDay()), t;
                    })(),
                        t = (function () {
                            const e = new Date(o);
                            if ((e.setMonth(e.getMonth() + 1), Qe(o, e))) e.setDate(0);
                            else
                                do {
                                    e.setDate(e.getDate() - 1);
                                } while (!Qe(o, e));
                            const t = new Date(e);
                            return t.setDate(t.getDate() + 6 - t.getDay()), t;
                        })(),
                        c = o.getMonth();
                    n(0, (a = []));
                    for (let n = e; n <= t; n.setDate(n.getDate() + 1)) {
                        const e = { day: n.getDate(), in: n.getMonth() === c, date: n.valueOf() };
                        let t = "";
                        if (((t += e.in ? "" : " inactive"), (t += u.indexOf(e.date) > -1 ? " selected" : ""), d)) {
                            const n = e.date == u[0],
                                l = e.date == u[1];
                            n && !l ? (t += " left") : l && !n && (t += " right"), e.date > u[0] && e.date < u[1] && (t += " inrange");
                        }
                        (t += an(n) ? " weekend" : ""), a.push({ ...e, css: t });
                    }
                }
            }),
            [a, i, p, l, o, c, s, u]
        );
    }
    class dn extends Ye {
        constructor(e) {
            super(), He(this, e, un, rn, a, { value: 3, current: 4, cancel: 5, part: 6 });
        }
    }
    function pn(e, t, n) {
        const l = e.slice();
        return (l[12] = t[n]), (l[14] = n), l;
    }
    function fn(e) {
        let t,
            n,
            l,
            o,
            c = e[12] + "";
        return {
            c() {
                (t = z("div")), (n = O(c)), (l = F()), U(t, "class", "month svelte-10az4re"), U(t, "data-id", (o = e[14])), V(t, "current", e[0] === e[14]);
            },
            m(e, o) {
                T(e, t, o), I(t, n), I(t, l);
            },
            p(e, n) {
                1 & n && V(t, "current", e[0] === e[14]);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function $n(e) {
        let t,
            l = e[1]("Close") + "";
        return {
            c() {
                t = O(l);
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function mn(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i = e[2],
            a = [];
        for (let t = 0; t < i.length; t += 1) a[t] = fn(pn(e, i, t));
        return (
            (o = new ct({ props: { type: "link", click: e[9], $$slots: { default: [$n] }, $$scope: { ctx: e } } })),
            {
                c() {
                    t = z("div");
                    for (let e = 0; e < a.length; e += 1) a[e].c();
                    (n = F()), (l = z("div")), Pe(o.$$.fragment), U(t, "class", "months svelte-10az4re"), U(l, "class", "buttons svelte-10az4re");
                },
                m(i, u) {
                    T(i, t, u);
                    for (let e = 0; e < a.length; e += 1) a[e].m(t, null);
                    T(i, n, u), T(i, l, u), Ke(o, l, null), (c = !0), s || ((r = k(Je.call(null, t, e[3]))), (s = !0));
                },
                p(e, [n]) {
                    if (5 & n) {
                        let l;
                        for (i = e[2], l = 0; l < i.length; l += 1) {
                            const o = pn(e, i, l);
                            a[l] ? a[l].p(o, n) : ((a[l] = fn(o)), a[l].c(), a[l].m(t, null));
                        }
                        for (; l < a.length; l += 1) a[l].d(1);
                        a.length = i.length;
                    }
                    const l = {};
                    32768 & n && (l.$$scope = { dirty: n, ctx: e }), o.$set(l);
                },
                i(e) {
                    c || (Ae(o.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(o.$$.fragment, e), (c = !1);
                },
                d(e) {
                    e && L(t), j(a, e), e && L(n), e && L(l), Ue(o), (s = !1), r();
                },
            }
        );
    }
    function hn(e, t, n) {
        const l = ae("wx-i18n") || Vt(Jt),
            o = l.getGroup("__dates"),
            c = l.getGroup("wx");
        let { value: s } = t,
            { current: r } = t,
            { cancel: i } = t,
            { part: a } = t;
        const u = o("monthsShort");
        let d;
        const p = { click: f };
        function f(e) {
            e && (r.setMonth(e), n(6, r)), "normal" === a && n(5, (s = new Date(r))), i();
        }
        return (
            (e.$$set = (e) => {
                "value" in e && n(5, (s = e.value)), "current" in e && n(6, (r = e.current)), "cancel" in e && n(7, (i = e.cancel)), "part" in e && n(8, (a = e.part));
            }),
            (e.$$.update = () => {
                352 & e.$$.dirty && ("normal" !== a && s ? ("left" === a && s.start ? n(0, (d = s.start.getMonth())) : "right" === a && s.end ? n(0, (d = s.end.getMonth())) : n(0, (d = r.getMonth()))) : n(0, (d = r.getMonth())));
            }),
            [d, c, u, p, f, s, r, i, a, () => f()]
        );
    }
    class gn extends Ye {
        constructor(e) {
            super(), He(this, e, hn, mn, a, { value: 5, current: 6, cancel: 7, part: 8 });
        }
    }
    function vn(e, t, n) {
        const l = e.slice();
        return (l[11] = t[n]), l;
    }
    function yn(e) {
        let t,
            n,
            l,
            o = e[11] + "";
        return {
            c() {
                (t = z("div")), (n = O(o)), U(t, "class", "year svelte-wb8cm"), U(t, "data-id", (l = e[11])), V(t, "current", e[1] == e[11]);
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, c) {
                1 & c && o !== (o = e[11] + "") && Y(n, o), 1 & c && l !== (l = e[11]) && U(t, "data-id", l), 3 & c && V(t, "current", e[1] == e[11]);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function wn(e) {
        let t,
            l = e[2]("Close") + "";
        return {
            c() {
                t = O(l);
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function bn(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i = e[0],
            a = [];
        for (let t = 0; t < i.length; t += 1) a[t] = yn(vn(e, i, t));
        return (
            (o = new ct({ props: { type: "link", click: e[9], $$slots: { default: [wn] }, $$scope: { ctx: e } } })),
            {
                c() {
                    t = z("div");
                    for (let e = 0; e < a.length; e += 1) a[e].c();
                    (n = F()), (l = z("div")), Pe(o.$$.fragment), U(t, "class", "years svelte-wb8cm"), U(l, "class", "buttons svelte-wb8cm");
                },
                m(i, u) {
                    T(i, t, u);
                    for (let e = 0; e < a.length; e += 1) a[e].m(t, null);
                    T(i, n, u), T(i, l, u), Ke(o, l, null), (c = !0), s || ((r = k(Je.call(null, t, e[3]))), (s = !0));
                },
                p(e, [n]) {
                    if (3 & n) {
                        let l;
                        for (i = e[0], l = 0; l < i.length; l += 1) {
                            const o = vn(e, i, l);
                            a[l] ? a[l].p(o, n) : ((a[l] = yn(o)), a[l].c(), a[l].m(t, null));
                        }
                        for (; l < a.length; l += 1) a[l].d(1);
                        a.length = i.length;
                    }
                    const l = {};
                    16384 & n && (l.$$scope = { dirty: n, ctx: e }), o.$set(l);
                },
                i(e) {
                    c || (Ae(o.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(o.$$.fragment, e), (c = !1);
                },
                d(e) {
                    e && L(t), j(a, e), e && L(n), e && L(l), Ue(o), (s = !1), r();
                },
            }
        );
    }
    function xn(e, t, n) {
        const l = (ae("wx-i18n") || Vt(Jt)).getGroup("wx");
        let o,
            c,
            { value: s } = t,
            { current: r } = t,
            { cancel: i } = t,
            { part: a } = t;
        const u = { click: d };
        function d(e) {
            e && (r.setFullYear(e), n(6, r)), "normal" === a && n(5, (s = new Date(r))), i();
        }
        return (
            (e.$$set = (e) => {
                "value" in e && n(5, (s = e.value)), "current" in e && n(6, (r = e.current)), "cancel" in e && n(7, (i = e.cancel)), "part" in e && n(8, (a = e.part));
            }),
            (e.$$.update = () => {
                if (353 & e.$$.dirty) {
                    "normal" !== a && s ? ("left" === a && s.start ? n(1, (c = s.start.getFullYear())) : "right" === a && s.end ? n(1, (c = s.end.getFullYear())) : n(1, (c = r.getFullYear()))) : n(1, (c = r.getFullYear()));
                    const { start: e, end: t } = We(r.getFullYear());
                    n(0, (o = []));
                    for (let n = e; n <= t; ++n) o.push(n);
                }
            }),
            [o, c, l, u, d, s, r, i, a, () => d()]
        );
    }
    class kn extends Ye {
        constructor(e) {
            super(), He(this, e, xn, bn, a, { value: 5, current: 6, cancel: 7, part: 8 });
        }
    }
    function Sn(e) {
        let t, n, l, o, c;
        return (
            (n = new ct({ props: { type: "link", click: e[11], $$slots: { default: [Mn] }, $$scope: { ctx: e } } })),
            (o = new ct({ props: { type: "link", click: e[12], $$slots: { default: [_n] }, $$scope: { ctx: e } } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), (l = F()), Pe(o.$$.fragment), U(t, "class", "buttons svelte-1ykoi1k");
                },
                m(e, s) {
                    T(e, t, s), Ke(n, t, null), I(t, l), Ke(o, t, null), (c = !0);
                },
                p(e, t) {
                    const l = {};
                    8388608 & t && (l.$$scope = { dirty: t, ctx: e }), n.$set(l);
                    const c = {};
                    8388608 & t && (c.$$scope = { dirty: t, ctx: e }), o.$set(c);
                },
                i(e) {
                    c || (Ae(n.$$.fragment, e), Ae(o.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), Ee(o.$$.fragment, e), (c = !1);
                },
                d(e) {
                    e && L(t), Ue(n), Ue(o);
                },
            }
        );
    }
    function Mn(e) {
        let t,
            l = e[4]("Clear") + "";
        return {
            c() {
                t = O(l);
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function _n(e) {
        let t,
            l = e[4]("Today") + "";
        return {
            c() {
                t = O(l);
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function Cn(e) {
        let t, n, l, o, c, s, r, i, a, u;
        function d(t) {
            e[8](t);
        }
        let p = { date: e[1], part: e[2], prev: e[5][e[3]].prev, next: e[5][e[3]].next };
        function f(t) {
            e[9](t);
        }
        function $(t) {
            e[10](t);
        }
        void 0 !== e[3] && (p.type = e[3]), (n = new nn({ props: p })), pe.push(() => Re(n, "type", d));
        var m = e[5][e[3]].component;
        function h(e) {
            let t = { part: e[2], cancel: e[5][e[3]].cancel };
            return void 0 !== e[0] && (t.value = e[0]), void 0 !== e[1] && (t.current = e[1]), { props: t };
        }
        m && ((s = new m(h(e))), pe.push(() => Re(s, "value", f)), pe.push(() => Re(s, "current", $)));
        let g = "month" === e[3] && "normal" === e[2] && Sn(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (o = F()), (c = z("div")), s && Pe(s.$$.fragment), (a = F()), g && g.c(), U(c, "class", "body svelte-1ykoi1k"), U(t, "class", "calendar svelte-1ykoi1k");
            },
            m(e, l) {
                T(e, t, l), Ke(n, t, null), I(t, o), I(t, c), s && Ke(s, c, null), I(c, a), g && g.m(c, null), (u = !0);
            },
            p(e, [t]) {
                const o = {};
                2 & t && (o.date = e[1]), 4 & t && (o.part = e[2]), 8 & t && (o.prev = e[5][e[3]].prev), 8 & t && (o.next = e[5][e[3]].next), !l && 8 & t && ((l = !0), (o.type = e[3]), ye(() => (l = !1))), n.$set(o);
                const u = {};
                if (
                    (4 & t && (u.part = e[2]),
                        8 & t && (u.cancel = e[5][e[3]].cancel),
                        !r && 1 & t && ((r = !0), (u.value = e[0]), ye(() => (r = !1))),
                        !i && 2 & t && ((i = !0), (u.current = e[1]), ye(() => (i = !1))),
                        m !== (m = e[5][e[3]].component))
                ) {
                    if (s) {
                        De();
                        const e = s;
                        Ee(e.$$.fragment, 1, 0, () => {
                            Ue(e, 1);
                        }),
                            Ie();
                    }
                    m ? ((s = new m(h(e))), pe.push(() => Re(s, "value", f)), pe.push(() => Re(s, "current", $)), Pe(s.$$.fragment), Ae(s.$$.fragment, 1), Ke(s, c, a)) : (s = null);
                } else m && s.$set(u);
                "month" === e[3] && "normal" === e[2]
                    ? g
                        ? (g.p(e, t), 12 & t && Ae(g, 1))
                        : ((g = Sn(e)), g.c(), Ae(g, 1), g.m(c, null))
                    : g &&
                    (De(),
                        Ee(g, 1, 1, () => {
                            g = null;
                        }),
                        Ie());
            },
            i(e) {
                u || (Ae(n.$$.fragment, e), s && Ae(s.$$.fragment, e), Ae(g), (u = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), s && Ee(s.$$.fragment, e), Ee(g), (u = !1);
            },
            d(e) {
                e && L(t), Ue(n), s && Ue(s), g && g.d();
            },
        };
    }
    function Dn(e, t, n) {
        const l = (ae("wx-i18n") || Vt(Jt)).getGroup("wx");
        let { value: o } = t,
            { current: c } = t,
            { cancel: s = null } = t,
            { part: r = "normal" } = t,
            i = "month";
        const a = {
            month: {
                component: dn,
                next: function () {
                    c.setMonth(c.getMonth() + 1), n(1, c);
                },
                prev: function () {
                    let e = new Date(c);
                    e.setMonth(c.getMonth() - 1);
                    for (; Qe(c, e);) e.setDate(e.getDate() - 1);
                    n(1, (c = e));
                },
                cancel: u,
            },
            year: {
                component: gn,
                next: function () {
                    c.setFullYear(c.getFullYear() + 1), n(1, c);
                },
                prev: function () {
                    c.setFullYear(c.getFullYear() - 1), n(1, c);
                },
                cancel: function () {
                    n(3, (i = "month"));
                },
            },
            duodecade: {
                component: kn,
                next: function () {
                    c.setFullYear(c.getFullYear() + 12), n(1, c);
                },
                prev: function () {
                    c.setFullYear(c.getFullYear() - 12), n(1, c);
                },
                cancel: function () {
                    n(3, (i = "year"));
                },
            },
        };
        function u() {
            o && "normal" === r && n(1, (c = new Date(o))), n(3, (i = "month")), s && s();
        }
        function d(e, t) {
            e.stopPropagation(), t ? (n(1, (c = new Date(t))), n(0, (o = new Date(c)))) : (n(0, (o = null)), n(1, (c = new Date()))), "normal" === r && u();
        }
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (o = e.value)), "current" in e && n(1, (c = e.current)), "cancel" in e && n(7, (s = e.cancel)), "part" in e && n(2, (r = e.part));
            }),
            [
                o,
                c,
                r,
                i,
                l,
                a,
                d,
                s,
                function (e) {
                    (i = e), n(3, i);
                },
                function (e) {
                    (o = e), n(0, o);
                },
                function (e) {
                    (c = e), n(1, c);
                },
                (e) => d(e),
                (e) => d(e, new Date()),
            ]
        );
    }
    class In extends Ye {
        constructor(e) {
            super(), He(this, e, Dn, Cn, a, { value: 0, current: 1, cancel: 7, part: 2 });
        }
    }
    function An(e) {
        let t, n;
        return (
            (t = new at({ props: { cancel: e[5], width: "unset", $$slots: { default: [En] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    1041 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function En(e) {
        let t, n, l, o;
        function c(t) {
            e[6](t);
        }
        function s(t) {
            e[7](t);
        }
        let r = { cancel: e[5] };
        return (
            void 0 !== e[0] && (r.value = e[0]),
            void 0 !== e[4] && (r.current = e[4]),
            (t = new In({ props: r })),
            pe.push(() => Re(t, "value", c)),
            pe.push(() => Re(t, "current", s)),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, n) {
                    Ke(t, e, n), (o = !0);
                },
                p(e, o) {
                    const c = {};
                    !n && 1 & o && ((n = !0), (c.value = e[0]), ye(() => (n = !1))), !l && 16 & o && ((l = !0), (c.current = e[4]), ye(() => (l = !1))), t.$set(c);
                },
                i(e) {
                    o || (Ae(t.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (o = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Tn(e) {
        let t, n, l, o, c, s, i, a;
        n = new Bt({ props: { value: e[0] ? Ve(e[0]) : e[0], id: e[1], readonly: !0, inputStyle: "cursor: pointer;" } });
        let u = e[3] && An(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = z("i")), (c = F()), u && u.c(), U(o, "class", "icon wxi-calendar svelte-1ere456"), U(t, "class", "layout svelte-1ere456");
            },
            m(r, d) {
                T(r, t, d), Ke(n, t, null), I(t, l), I(t, o), I(t, c), u && u.m(t, null), e[8](t), (s = !0), i || ((a = [R(window, "scroll", e[5]), R(t, "click", e[9])]), (i = !0));
            },
            p(e, [l]) {
                const o = {};
                1 & l && (o.value = e[0] ? Ve(e[0]) : e[0]),
                    2 & l && (o.id = e[1]),
                    n.$set(o),
                    e[3]
                        ? u
                            ? (u.p(e, l), 8 & l && Ae(u, 1))
                            : ((u = An(e)), u.c(), Ae(u, 1), u.m(t, null))
                        : u &&
                        (De(),
                            Ee(u, 1, 1, () => {
                                u = null;
                            }),
                            Ie());
            },
            i(e) {
                s || (Ae(n.$$.fragment, e), Ae(u), (s = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), Ee(u), (s = !1);
            },
            d(l) {
                l && L(t), Ue(n), u && u.d(), e[8](null), (i = !1), r(a);
            },
        };
    }
    function Ln(e, t, n) {
        let l,
            o,
            { value: c } = t,
            { id: s = et() } = t;
        let r = c ? new Date(c) : new Date();
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (c = e.value)), "id" in e && n(1, (s = e.id));
            }),
            [
                c,
                s,
                l,
                o,
                r,
                function () {
                    n(3, (o = null));
                },
                function (e) {
                    (c = e), n(0, c);
                },
                function (e) {
                    (r = e), n(4, r);
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (l = e), n(2, l);
                    });
                },
                () => n(3, (o = l)),
            ]
        );
    }
    const jn = [];
    function zn(e, t = n) {
        let l;
        const o = new Set();
        function c(t) {
            if (a(e, t) && ((e = t), l)) {
                const t = !jn.length;
                for (const t of o) t[1](), jn.push(t, e);
                if (t) {
                    for (let e = 0; e < jn.length; e += 2) jn[e][0](jn[e + 1]);
                    jn.length = 0;
                }
            }
        }
        return {
            set: c,
            update: function (t) {
                c(t(e));
            },
            subscribe: function (s, r = n) {
                const i = [s, r];
                return (
                    o.add(i),
                    1 === o.size && (l = t(c) || n),
                    s(e),
                    () => {
                        o.delete(i), 0 === o.size && (l(), (l = null));
                    }
                );
            },
        };
    }
    function Nn(e) {
        let t, n;
        return (
            (t = new at({ props: { cancel: e[8], width: "unset", $$slots: { default: [Rn] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    131121 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function On(e) {
        let t;
        return {
            c() {
                t = O("Done");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Fn(e) {
        let t;
        return {
            c() {
                t = O("Clear");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function qn(e) {
        let t;
        return {
            c() {
                t = O("Today");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Rn(e) {
        let t, n, l, o, c, s, r, i, a, u, d, p, f, $, m, h, g, v, y, w;
        function b(t) {
            e[10](t);
        }
        function x(t) {
            e[11](t);
        }
        let k = { part: "left" };
        function S(t) {
            e[12](t);
        }
        function M(t) {
            e[13](t);
        }
        void 0 !== e[0] && (k.value = e[0]), void 0 !== e[4] && (k.current = e[4]), (o = new In({ props: k })), pe.push(() => Re(o, "value", b)), pe.push(() => Re(o, "current", x));
        let _ = { part: "right" };
        return (
            void 0 !== e[0] && (_.value = e[0]),
            void 0 !== e[5] && (_.current = e[5]),
            (a = new In({ props: _ })),
            pe.push(() => Re(a, "value", S)),
            pe.push(() => Re(a, "current", M)),
            (m = new ct({ props: { type: "primary", click: e[8], $$slots: { default: [On] }, $$scope: { ctx: e } } })),
            (g = new ct({ props: { type: "link", click: e[14], $$slots: { default: [Fn] }, $$scope: { ctx: e } } })),
            (y = new ct({ props: { type: "link", click: e[15], $$slots: { default: [qn] }, $$scope: { ctx: e } } })),
            {
                c() {
                    (t = z("div")),
                        (n = z("div")),
                        (l = z("div")),
                        Pe(o.$$.fragment),
                        (r = F()),
                        (i = z("div")),
                        Pe(a.$$.fragment),
                        (p = F()),
                        (f = z("div")),
                        ($ = z("div")),
                        Pe(m.$$.fragment),
                        (h = F()),
                        Pe(g.$$.fragment),
                        (v = F()),
                        Pe(y.$$.fragment),
                        U(l, "class", "half svelte-yv5cw2"),
                        U(i, "class", "half svelte-yv5cw2"),
                        U(n, "class", "calendars svelte-yv5cw2"),
                        U($, "class", "done svelte-yv5cw2"),
                        U(f, "class", "buttons svelte-yv5cw2"),
                        U(t, "class", "calendar svelte-yv5cw2");
                },
                m(e, c) {
                    T(e, t, c), I(t, n), I(n, l), Ke(o, l, null), I(n, r), I(n, i), Ke(a, i, null), I(t, p), I(t, f), I(f, $), Ke(m, $, null), I(f, h), Ke(g, f, null), I(f, v), Ke(y, f, null), (w = !0);
                },
                p(e, t) {
                    const n = {};
                    !c && 1 & t && ((c = !0), (n.value = e[0]), ye(() => (c = !1))), !s && 16 & t && ((s = !0), (n.current = e[4]), ye(() => (s = !1))), o.$set(n);
                    const l = {};
                    !u && 1 & t && ((u = !0), (l.value = e[0]), ye(() => (u = !1))), !d && 32 & t && ((d = !0), (l.current = e[5]), ye(() => (d = !1))), a.$set(l);
                    const r = {};
                    131072 & t && (r.$$scope = { dirty: t, ctx: e }), m.$set(r);
                    const i = {};
                    131072 & t && (i.$$scope = { dirty: t, ctx: e }), g.$set(i);
                    const p = {};
                    131072 & t && (p.$$scope = { dirty: t, ctx: e }), y.$set(p);
                },
                i(e) {
                    w || (Ae(o.$$.fragment, e), Ae(a.$$.fragment, e), Ae(m.$$.fragment, e), Ae(g.$$.fragment, e), Ae(y.$$.fragment, e), (w = !0));
                },
                o(e) {
                    Ee(o.$$.fragment, e), Ee(a.$$.fragment, e), Ee(m.$$.fragment, e), Ee(g.$$.fragment, e), Ee(y.$$.fragment, e), (w = !1);
                },
                d(e) {
                    e && L(t), Ue(o), Ue(a), Ue(m), Ue(g), Ue(y);
                },
            }
        );
    }
    function Pn(e) {
        let t, n, l, o, c, s, i, a;
        n = new Bt({ props: { value: e[3], id: e[1], readonly: !0, inputStyle: "cursor: pointer; text-overflow: ellipsis; padding-right: 18px;" } });
        let u = e[2] && Nn(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = z("i")), (c = F()), u && u.c(), U(o, "class", "icon wxi-calendar svelte-yv5cw2"), U(t, "class", "layout svelte-yv5cw2");
            },
            m(r, d) {
                T(r, t, d), Ke(n, t, null), I(t, l), I(t, o), I(t, c), u && u.m(t, null), (s = !0), i || ((a = [R(window, "scroll", e[8]), R(t, "click", e[16])]), (i = !0));
            },
            p(e, [l]) {
                const o = {};
                8 & l && (o.value = e[3]),
                    2 & l && (o.id = e[1]),
                    n.$set(o),
                    e[2]
                        ? u
                            ? (u.p(e, l), 4 & l && Ae(u, 1))
                            : ((u = Nn(e)), u.c(), Ae(u, 1), u.m(t, null))
                        : u &&
                        (De(),
                            Ee(u, 1, 1, () => {
                                u = null;
                            }),
                            Ie());
            },
            i(e) {
                s || (Ae(n.$$.fragment, e), Ae(u), (s = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), Ee(u), (s = !1);
            },
            d(e) {
                e && L(t), Ue(n), u && u.d(), (i = !1), r(a);
            },
        };
    }
    function Kn(e, t, n) {
        let l,
            o,
            c,
            { value: s = { start: null, end: null } } = t,
            { id: r = et() } = t;
        const i = zn(s && s.start ? new Date(s.start) : new Date());
        f(e, i, (e) => n(4, (l = e)));
        const a = zn(l);
        function u(e, t) {
            t ? (i.set(new Date(t)), s || n(0, (s = { start: null, end: null })), s.end || !s.start ? n(0, (s = { start: new Date(l), end: null })) : n(0, (s.end = new Date(l)), s)) : (i.set(new Date()), n(0, (s = null)));
        }
        let d;
        f(e, a, (e) => n(5, (o = e))),
            i.subscribe((e) => {
                const t = new Date(e);
                t.setMonth(t.getMonth() + 1), t.valueOf() != o.valueOf() && a.set(t);
            }),
            a.subscribe((e) => {
                const t = new Date(e);
                t.setMonth(t.getMonth() - 1), t.valueOf() != l.valueOf() && i.set(t);
            });
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (s = e.value)), "id" in e && n(1, (r = e.id));
            }),
            (e.$$.update = () => {
                1 & e.$$.dirty && n(3, (d = s ? (s.start ? Ve(s.start) + (s.end ? ` - ${Ve(s.end)}` : "") : Ve(s)) : s));
            }),
            [
                s,
                r,
                c,
                d,
                l,
                o,
                i,
                a,
                function (e) {
                    e.stopPropagation(), s && s.start && i.set(new Date(s.start)), n(2, (c = null));
                },
                u,
                function (e) {
                    (s = e), n(0, s);
                },
                function (e) {
                    (l = e), i.set(l);
                },
                function (e) {
                    (s = e), n(0, s);
                },
                function (e) {
                    (o = e), a.set(o);
                },
                (e) => u(),
                (e) => u(0, new Date()),
                () => n(2, (c = !0)),
            ]
        );
    }
    function Un(e, t, n) {
        const l = e.slice();
        return (l[16] = t[n]), l;
    }
    function Hn(e, t, n) {
        const l = e.slice();
        return (l[16] = t[n]), l;
    }
    function Yn(e, t) {
        let n,
            l,
            o,
            c,
            s,
            i = t[16].name + "";
        function a() {
            return t[8](t[16]);
        }
        function u() {
            return t[9](t[16]);
        }
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), (l = O(i)), (o = F()), U(n, "class", "item svelte-zsg7sz"), V(n, "active", t[0][0].includes(t[16].id)), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), I(n, l), I(n, o), c || ((s = [R(n, "click", a), R(n, "dblclick", u)]), (c = !0));
            },
            p(e, o) {
                (t = e), 2 & o && i !== (i = t[16].name + "") && Y(l, i), 3 & o && V(n, "active", t[0][0].includes(t[16].id));
            },
            d(e) {
                e && L(n), (c = !1), r(s);
            },
        };
    }
    function Bn(e, t) {
        let n,
            l,
            o,
            c,
            s,
            i = t[16].name + "";
        function a() {
            return t[14](t[16]);
        }
        function u() {
            return t[15](t[16]);
        }
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), (l = O(i)), (o = F()), U(n, "class", "item svelte-zsg7sz"), V(n, "active", t[0][1].includes(t[16].id)), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), I(n, l), I(n, o), c || ((s = [R(n, "click", a), R(n, "dblclick", u)]), (c = !0));
            },
            p(e, o) {
                (t = e), 2 & o && i !== (i = t[16].name + "") && Y(l, i), 3 & o && V(n, "active", t[0][1].includes(t[16].id));
            },
            d(e) {
                e && L(n), (c = !1), r(s);
            },
        };
    }
    function Gn(e) {
        let t,
            l,
            o,
            c,
            s,
            i,
            a,
            u,
            d,
            p,
            f,
            $,
            m,
            h,
            g,
            v = [],
            y = new Map(),
            w = [],
            b = new Map(),
            x = e[1][0];
        const k = (e) => e[16].id;
        for (let t = 0; t < x.length; t += 1) {
            let n = Hn(e, x, t),
                l = k(n);
            y.set(l, (v[t] = Yn(l, n)));
        }
        let S = e[1][1];
        const M = (e) => e[16].id;
        for (let t = 0; t < S.length; t += 1) {
            let n = Un(e, S, t),
                l = M(n);
            b.set(l, (w[t] = Bn(l, n)));
        }
        return {
            c() {
                (t = z("div")), (l = z("div"));
                for (let e = 0; e < v.length; e += 1) v[e].c();
                (o = F()),
                    (c = z("div")),
                    (s = z("div")),
                    (s.innerHTML = '<i class="wxi-angle-dbl-left"></i>'),
                    (i = F()),
                    (a = z("div")),
                    (a.innerHTML = '<i class="wxi-angle-dbl-right"></i>'),
                    (u = F()),
                    (d = z("div")),
                    (d.innerHTML = '<i class="wxi-angle-left"></i>'),
                    (p = F()),
                    (f = z("div")),
                    (f.innerHTML = '<i class="wxi-angle-right"></i>'),
                    ($ = F()),
                    (m = z("div"));
                for (let e = 0; e < w.length; e += 1) w[e].c();
                U(l, "class", "list svelte-zsg7sz"),
                    U(s, "class", "icon svelte-zsg7sz"),
                    U(a, "class", "icon svelte-zsg7sz"),
                    U(d, "class", "icon svelte-zsg7sz"),
                    U(f, "class", "icon svelte-zsg7sz"),
                    U(c, "class", "controls svelte-zsg7sz"),
                    U(m, "class", "list svelte-zsg7sz"),
                    U(t, "class", "layout svelte-zsg7sz");
            },
            m(n, r) {
                T(n, t, r), I(t, l);
                for (let e = 0; e < v.length; e += 1) v[e].m(l, null);
                I(t, o), I(t, c), I(c, s), I(c, i), I(c, a), I(c, u), I(c, d), I(c, p), I(c, f), I(t, $), I(t, m);
                for (let e = 0; e < w.length; e += 1) w[e].m(m, null);
                h || ((g = [R(s, "click", e[10]), R(a, "click", e[11]), R(d, "click", e[12]), R(f, "click", e[13])]), (h = !0));
            },
            p(e, [t]) {
                15 & t && ((x = e[1][0]), (v = Oe(v, t, k, 1, e, x, y, l, ze, Yn, null, Hn))), 15 & t && ((S = e[1][1]), (w = Oe(w, t, M, 1, e, S, b, m, ze, Bn, null, Un)));
            },
            i: n,
            o: n,
            d(e) {
                e && L(t);
                for (let e = 0; e < v.length; e += 1) v[e].d();
                for (let e = 0; e < w.length; e += 1) w[e].d();
                (h = !1), r(g);
            },
        };
    }
    function Jn(e, t, n) {
        let { data: l = [] } = t,
            { values: o = [] } = t;
        ce(() => {
            n(1, (s[0] = l.filter((e) => !o.includes(e.id))), s), n(1, (s[1] = l.filter((e) => o.includes(e.id))), s);
        }),
            se(() => {
                n(6, (o = s[1].map((e) => e.id)));
            });
        const c = [[], []],
            s = [l, []];
        function r(e, t) {
            let l = c[e];
            (l = l.includes(t) ? l.filter((e) => e !== t) : [...l, t]), n(0, (c[e] = l), c);
        }
        function i(e, t) {
            const o = e ? 0 : 1,
                c = l.find((e) => e.id === t);
            n(1, (s[e] = s[e].filter((e) => e.id !== t)), s), n(1, (s[o] = [...s[o], c]), s);
        }
        function a(e) {
            const t = e ? 0 : 1,
                o = c[e];
            n(1, (s[e] = s[e].filter((e) => !o.includes(e.id))), s), n(1, (s[t] = [...s[t], ...l.filter((e) => o.includes(e.id))]), s), n(0, (c[e] = []), c);
        }
        function u(e) {
            const t = e ? 0 : 1,
                l = s[e];
            n(1, (s[t] = [...s[t], ...l]), s), n(1, (s[e] = []), s), n(0, (c[e] = []), c);
        }
        return (
            (e.$$set = (e) => {
                "data" in e && n(7, (l = e.data)), "values" in e && n(6, (o = e.values));
            }),
            [c, s, r, i, a, u, o, l, (e) => r(0, e.id), (e) => i(0, e.id), () => u(1), () => u(0), () => a(1), () => a(0), (e) => r(1, e.id), (e) => i(1, e.id)]
        );
    }
    function Vn(e, t, n) {
        const l = e.slice();
        return (l[27] = t[n]), (l[28] = t), (l[29] = n), l;
    }
    const Xn = (e) => ({ option: 8 & e[0] }),
        Qn = (e) => ({ option: e[27] });
    function Wn(e, t, n) {
        const l = e.slice();
        return (l[30] = t[n]), l;
    }
    function Zn(e, t) {
        let n,
            l,
            o,
            c,
            s,
            r,
            i,
            a = t[30][t[1]] + "";
        function u() {
            return t[19](t[30]);
        }
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), (l = z("span")), (o = O(a)), (c = F()), (s = z("i")), U(l, "class", "label"), U(s, "class", "wxi-close svelte-15wp6mr"), U(n, "class", "tag svelte-15wp6mr"), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), I(n, l), I(l, o), I(n, c), I(n, s), r || ((i = R(s, "click", K(u))), (r = !0));
            },
            p(e, n) {
                (t = e), 34 & n[0] && a !== (a = t[30][t[1]] + "") && Y(o, a);
            },
            d(e) {
                e && L(n), (r = !1), i();
            },
        };
    }
    function el(e) {
        let t, n;
        return (
            (t = new at({ props: { cancel: e[11], $$slots: { default: [ol] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    8388697 & n[0] && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function tl(e) {
        let t;
        return {
            c() {
                (t = z("div")), (t.textContent = "No data"), U(t, "class", "no-data svelte-15wp6mr");
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            i: n,
            o: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function nl(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[3];
        const s = (e) => e[27].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = Vn(e, c, t),
                r = s(n);
            o.set(r, (l[t] = ll(r, n)));
        }
        return {
            c() {
                for (let e = 0; e < l.length; e += 1) l[e].c();
                t = q();
            },
            m(e, o) {
                for (let t = 0; t < l.length; t += 1) l[t].m(e, o);
                T(e, t, o), (n = !0);
            },
            p(e, n) {
                8396889 & n[0] && ((c = e[3]), De(), (l = Oe(l, n, s, 1, e, c, o, t.parentNode, Ne, ll, t, Vn)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                for (let t = 0; t < l.length; t += 1) l[t].d(e);
                e && L(t);
            },
        };
    }
    function ll(e, t) {
        let n,
            l,
            o,
            c,
            s,
            r,
            i = t[29];
        (l = new bt({ props: { value: t[0].includes(t[27].id) } })),
            l.$on("click", function () {
                return t[21](t[27]);
            });
        const a = t[18].default,
            u = $(a, t, t[23], Qn),
            d = () => t[22](n, i),
            p = () => t[22](null, i);
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), Pe(l.$$.fragment), (o = F()), u && u.c(), (c = F()), U(n, "class", "item svelte-15wp6mr"), U(n, "data-id", (s = t[27].id)), V(n, "navigate", t[6] && t[6].id === t[27].id), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), Ke(l, n, null), I(n, o), u && u.m(n, null), I(n, c), d(), (r = !0);
            },
            p(e, o) {
                t = e;
                const c = {};
                9 & o[0] && (c.value = t[0].includes(t[27].id)),
                    l.$set(c),
                    u && u.p && (!r || 8388616 & o[0]) && g(u, a, t, t[23], r ? h(a, t[23], o, Xn) : v(t[23]), Qn),
                    (!r || (8 & o[0] && s !== (s = t[27].id))) && U(n, "data-id", s),
                    i !== t[29] && (p(), (i = t[29]), d()),
                    72 & o[0] && V(n, "navigate", t[6] && t[6].id === t[27].id);
            },
            i(e) {
                r || (Ae(l.$$.fragment, e), Ae(u, e), (r = !0));
            },
            o(e) {
                Ee(l.$$.fragment, e), Ee(u, e), (r = !1);
            },
            d(e) {
                e && L(n), Ue(l), u && u.d(e), p();
            },
        };
    }
    function ol(e) {
        let t, n, l, o, c, s;
        const i = [nl, tl],
            a = [];
        function u(e, t) {
            return e[3].length ? 0 : 1;
        }
        return (
            (n = u(e)),
            (l = a[n] = i[n](e)),
            {
                c() {
                    (t = z("div")), l.c(), U(t, "class", "list svelte-15wp6mr");
                },
                m(l, r) {
                    T(l, t, r), a[n].m(t, null), (o = !0), c || ((s = [R(t, "click", e[12]), R(t, "mousemove", e[14])]), (c = !0));
                },
                p(e, o) {
                    let c = n;
                    (n = u(e)),
                        n === c
                            ? a[n].p(e, o)
                            : (De(),
                                Ee(a[c], 1, 1, () => {
                                    a[c] = null;
                                }),
                                Ie(),
                                (l = a[n]),
                                l ? l.p(e, o) : ((l = a[n] = i[n](e)), l.c()),
                                Ae(l, 1),
                                l.m(t, null));
                },
                i(e) {
                    o || (Ae(l), (o = !0));
                },
                o(e) {
                    Ee(l), (o = !1);
                },
                d(e) {
                    e && L(t), a[n].d(), (c = !1), r(s);
                },
            }
        );
    }
    function cl(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            i,
            a,
            u,
            d,
            p,
            f = [],
            $ = new Map(),
            m = e[5];
        const h = (e) => e[30].id;
        for (let t = 0; t < m.length; t += 1) {
            let n = Wn(e, m, t),
                l = h(n);
            $.set(l, (f[t] = Zn(l, n)));
        }
        let g = e[2] && el(e);
        return {
            c() {
                (t = z("div")), (n = z("div"));
                for (let e = 0; e < f.length; e += 1) f[e].c();
                (l = F()),
                    (o = z("div")),
                    (c = z("input")),
                    (s = F()),
                    (i = z("i")),
                    (a = F()),
                    g && g.c(),
                    U(c, "type", "text"),
                    U(c, "class", "input svelte-15wp6mr"),
                    U(i, "class", "wxi-angle-down svelte-15wp6mr"),
                    U(o, "class", "select svelte-15wp6mr"),
                    U(n, "class", "wrapper svelte-15wp6mr"),
                    V(n, "active", e[2]),
                    U(t, "class", "layout svelte-15wp6mr");
            },
            m(r, $) {
                T(r, t, $), I(t, n);
                for (let e = 0; e < f.length; e += 1) f[e].m(n, null);
                I(n, l), I(n, o), I(o, c), B(c, e[7]), I(o, s), I(o, i), I(t, a), g && g.m(t, null), (u = !0), d || ((p = [R(c, "input", e[20]), R(c, "input", e[10]), R(c, "keydown", e[15]), R(n, "click", e[8])]), (d = !0));
            },
            p(e, o) {
                546 & o[0] && ((m = e[5]), (f = Oe(f, o, h, 1, e, m, $, n, ze, Zn, l, Wn))),
                    128 & o[0] && c.value !== e[7] && B(c, e[7]),
                    4 & o[0] && V(n, "active", e[2]),
                    e[2]
                        ? g
                            ? (g.p(e, o), 4 & o[0] && Ae(g, 1))
                            : ((g = el(e)), g.c(), Ae(g, 1), g.m(t, null))
                        : g &&
                        (De(),
                            Ee(g, 1, 1, () => {
                                g = null;
                            }),
                            Ie());
            },
            i(e) {
                u || (Ae(g), (u = !0));
            },
            o(e) {
                Ee(g), (u = !1);
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < f.length; e += 1) f[e].d();
                g && g.d(), (d = !1), r(p);
            },
        };
    }
    function sl(e, t, n) {
        let l,
            o,
            c,
            s,
            { $$slots: r = {}, $$scope: i } = t,
            { options: a = [] } = t,
            { values: u = [] } = t,
            { key: d = "label" } = t,
            p = {},
            f = [],
            $ = "",
            m = "";
        function h() {
            n(2, (l = !0)), (s = 0), n(6, (c = o[s]));
        }
        function g(e) {
            n(0, (u = u.filter((t) => t !== e)));
        }
        function v() {
            n(2, (l = null)), n(17, ($ = "")), n(7, (m = ""));
        }
        function y(e) {
            n(0, (u = u.includes(e) ? u.filter((t) => t !== e) : [...u, e]));
        }
        function w(e) {
            (s += e), s > o.length - 1 ? (s = o.length - 1) : s < 0 && (s = 0), n(6, (c = o[s])), p[s].scrollIntoView({ block: "nearest" });
        }
        return (
            (e.$$set = (e) => {
                "options" in e && n(16, (a = e.options)), "values" in e && n(0, (u = e.values)), "key" in e && n(1, (d = e.key)), "$$scope" in e && n(23, (i = e.$$scope));
            }),
            (e.$$.update = () => {
                196610 & e.$$.dirty[0] && n(3, (o = a.filter((e) => e[d].toLowerCase().includes($.toLowerCase())))), 65537 & e.$$.dirty[0] && n(5, (f = a.filter((e) => u.includes(e.id))));
            }),
            [
                u,
                d,
                l,
                o,
                p,
                f,
                c,
                m,
                h,
                g,
                function () {
                    l || h(), n(17, ($ = m || ""));
                },
                v,
                function (e) {
                    const t = Ge(e);
                    t && n(0, (u = u.includes(t) ? u.filter((e) => e !== t) : [...u, t]));
                },
                y,
                function (e) {
                    const t = Ge(e);
                    t && ((s = o.findIndex((e) => e.id === t)), n(6, (c = o[s])));
                },
                function (e) {
                    switch (e.code) {
                        case "Space":
                            l ? v() : h();
                            break;
                        case "Tab":
                            l && v();
                            break;
                        case "Enter":
                            l
                                ? (function () {
                                    const e = c.id;
                                    n(0, (u = u.includes(e) ? u.filter((t) => t !== e) : [...u, e]));
                                })()
                                : h();
                            break;
                        case "ArrowDown":
                            l ? w(1) : h();
                            break;
                        case "ArrowUp":
                            l ? w(-1) : h();
                            break;
                        case "Escape":
                            v();
                    }
                },
                a,
                $,
                r,
                (e) => g(e.id),
                function () {
                    (m = this.value), n(7, m);
                },
                (e) => y(e.id),
                function (e, t) {
                    pe[e ? "unshift" : "push"](() => {
                        (p[t] = e), n(4, p), n(3, o), n(16, a), n(1, d), n(17, $);
                    });
                },
                i,
            ]
        );
    }
    function rl(e, t, n) {
        const l = e.slice();
        return (l[18] = t[n]), l;
    }
    const il = (e) => ({ option: 16 & e }),
        al = (e) => ({ option: e[18] });
    function ul(e, t, n) {
        const l = e.slice();
        return (l[18] = t[n]), l;
    }
    const dl = (e) => ({ option: 8 & e }),
        pl = (e) => ({ option: e[18] });
    function fl(e) {
        let t;
        return {
            c() {
                (t = z("i")), U(t, "class", "wx-list-icon wx-hover wxi-edit"), U(t, "data-action", "edit"), U(t, "title", "Edit");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function $l(e) {
        let t;
        return {
            c() {
                (t = z("i")), U(t, "class", "wx-list-icon wx-hover wxi-delete"), U(t, "data-action", "remove"), U(t, "title", "Delete");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function ml(e, t) {
        let n, l, o, c, s, r;
        const i = t[12].default,
            a = $(i, t, t[14], pl);
        let u = t[0] && fl(),
            d = t[1] && $l();
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), a && a.c(), (l = F()), u && u.c(), (o = F()), d && d.c(), (c = F()), U(n, "class", "wx-list-item"), U(n, "data-id", (s = t[18].id)), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), a && a.m(n, null), I(n, l), u && u.m(n, null), I(n, o), d && d.m(n, null), I(n, c), (r = !0);
            },
            p(e, l) {
                (t = e),
                    a && a.p && (!r || 16392 & l) && g(a, i, t, t[14], r ? h(i, t[14], l, dl) : v(t[14]), pl),
                    t[0] ? u || ((u = fl()), u.c(), u.m(n, o)) : u && (u.d(1), (u = null)),
                    t[1] ? d || ((d = $l()), d.c(), d.m(n, c)) : d && (d.d(1), (d = null)),
                    (!r || (8 & l && s !== (s = t[18].id))) && U(n, "data-id", s);
            },
            i(e) {
                r || (Ae(a, e), (r = !0));
            },
            o(e) {
                Ee(a, e), (r = !1);
            },
            d(e) {
                e && L(n), a && a.d(e), u && u.d(), d && d.d();
            },
        };
    }
    function hl(e) {
        let t, n;
        return (
            (t = new ct({ props: { type: "link", click: e[8], $$slots: { default: [gl] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    16388 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function gl(e) {
        let t, n;
        return {
            c() {
                (t = O("Add ")), (n = O(e[2]));
            },
            m(e, l) {
                T(e, t, l), T(e, n, l);
            },
            p(e, t) {
                4 & t && Y(n, e[2]);
            },
            d(e) {
                e && L(t), e && L(n);
            },
        };
    }
    function vl(e) {
        let t, n;
        return (
            (t = new at({ props: { area: e[5], cancel: e[13], $$slots: { default: [wl] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    32 & n && (l.area = e[5]), 32 & n && (l.cancel = e[13]), 16400 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function yl(e, t) {
        let n, l, o, c, s, r;
        const i = t[12].default,
            a = $(i, t, t[14], al);
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), a && a.c(), (l = F()), U(n, "class", "wx-list-item"), U(n, "data-id", (o = t[18].id)), (this.first = n);
            },
            m(e, o) {
                T(e, n, o), a && a.m(n, null), I(n, l), (c = !0), s || ((r = k(Je.call(null, n, t[7]))), (s = !0));
            },
            p(e, l) {
                (t = e), a && a.p && (!c || 16400 & l) && g(a, i, t, t[14], c ? h(i, t[14], l, il) : v(t[14]), al), (!c || (16 & l && o !== (o = t[18].id))) && U(n, "data-id", o);
            },
            i(e) {
                c || (Ae(a, e), (c = !0));
            },
            o(e) {
                Ee(a, e), (c = !1);
            },
            d(e) {
                e && L(n), a && a.d(e), (s = !1), r();
            },
        };
    }
    function wl(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[4];
        const s = (e) => e[18].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = rl(e, c, t),
                r = s(n);
            o.set(r, (l[t] = yl(r, n)));
        }
        return {
            c() {
                t = z("div");
                for (let e = 0; e < l.length; e += 1) l[e].c();
                U(t, "class", "wx-list list svelte-131aqzh");
            },
            m(e, o) {
                T(e, t, o);
                for (let e = 0; e < l.length; e += 1) l[e].m(t, null);
                n = !0;
            },
            p(e, n) {
                16528 & n && ((c = e[4]), De(), (l = Oe(l, n, s, 1, e, c, o, t, Ne, yl, null, rl)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < l.length; e += 1) l[e].d();
            },
        };
    }
    function bl(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i = [],
            a = new Map(),
            u = e[3];
        const d = (e) => e[18].id;
        for (let t = 0; t < u.length; t += 1) {
            let n = ul(e, u, t),
                l = d(n);
            a.set(l, (i[t] = ml(l, n)));
        }
        let p = e[1] && e[4].length && hl(e),
            f = e[5] && vl(e);
        return {
            c() {
                (t = z("div")), (n = z("div"));
                for (let e = 0; e < i.length; e += 1) i[e].c();
                (l = F()), p && p.c(), (o = F()), f && f.c(), U(n, "class", "wx-list"), U(t, "class", "layout svelte-131aqzh");
            },
            m(a, u) {
                T(a, t, u), I(t, n);
                for (let e = 0; e < i.length; e += 1) i[e].m(n, null);
                I(t, l), p && p.m(t, null), I(t, o), f && f.m(t, null), (c = !0), s || ((r = k(Je.call(null, n, e[6]))), (s = !0));
            },
            p(e, [l]) {
                16395 & l && ((u = e[3]), De(), (i = Oe(i, l, d, 1, e, u, a, n, Ne, ml, null, ul)), Ie()),
                    e[1] && e[4].length
                        ? p
                            ? (p.p(e, l), 18 & l && Ae(p, 1))
                            : ((p = hl(e)), p.c(), Ae(p, 1), p.m(t, o))
                        : p &&
                        (De(),
                            Ee(p, 1, 1, () => {
                                p = null;
                            }),
                            Ie()),
                    e[5]
                        ? f
                            ? (f.p(e, l), 32 & l && Ae(f, 1))
                            : ((f = vl(e)), f.c(), Ae(f, 1), f.m(t, null))
                        : f &&
                        (De(),
                            Ee(f, 1, 1, () => {
                                f = null;
                            }),
                            Ie());
            },
            i(e) {
                if (!c) {
                    for (let e = 0; e < u.length; e += 1) Ae(i[e]);
                    Ae(p), Ae(f), (c = !0);
                }
            },
            o(e) {
                for (let e = 0; e < i.length; e += 1) Ee(i[e]);
                Ee(p), Ee(f), (c = !1);
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < i.length; e += 1) i[e].d();
                p && p.d(), f && f.d(), (s = !1), r();
            },
        };
    }
    function xl(e, t, n) {
        let { $$slots: l = {}, $$scope: o } = t,
            { options: c = [] } = t,
            { selected: s = [] } = t,
            { canEdit: r = !1 } = t,
            { canDelete: i = !0 } = t,
            { edit: a } = t,
            { title: u = "" } = t,
            d = [],
            p = [];
        let f = {
            remove: function (e) {
                return n(9, (s = s.filter((t) => t !== e))), !1;
            },
            edit: function (e) {
                return r && a(e), !1;
            },
        },
            $ = {
                click: function (e) {
                    n(9, (s = [...(s || []), e])), n(5, (m = null));
                },
            },
            m = null;
        return (
            (e.$$set = (e) => {
                "options" in e && n(10, (c = e.options)),
                    "selected" in e && n(9, (s = e.selected)),
                    "canEdit" in e && n(0, (r = e.canEdit)),
                    "canDelete" in e && n(1, (i = e.canDelete)),
                    "edit" in e && n(11, (a = e.edit)),
                    "title" in e && n(2, (u = e.title)),
                    "$$scope" in e && n(14, (o = e.$$scope));
            }),
            (e.$$.update = () => {
                1560 & e.$$.dirty &&
                    (n(3, (d = [])),
                        n(4, (p = [])),
                        c.forEach((e) => {
                            s && -1 !== s.indexOf(e.id) ? d.push(e) : p.push(e);
                        }));
            }),
            [
                r,
                i,
                u,
                d,
                p,
                m,
                f,
                $,
                function (e) {
                    n(5, (m = e.target.getBoundingClientRect()));
                },
                s,
                c,
                a,
                l,
                () => n(5, (m = null)),
                o,
            ]
        );
    }
    function kl(e) {
        let t;
        return {
            c() {
                t = O("Save");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Sl(e) {
        let t, n, l, o, c, s, i;
        return (
            (n = new ct({ props: { click: e[6], $$slots: { default: [kl] }, $$scope: { ctx: e } } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = z("input")), U(o, "type", "text"), U(o, "class", "svelte-1dqaxa5"), U(t, "class", "line svelte-1dqaxa5");
                },
                m(r, a) {
                    T(r, t, a), Ke(n, t, null), I(t, l), I(t, o), e[7](o), B(o, e[0]), (c = !0), s || ((i = [R(o, "input", e[8]), R(o, "keydown", e[3])]), (s = !0));
                },
                p(e, [t]) {
                    const l = {};
                    2048 & t && (l.$$scope = { dirty: t, ctx: e }), n.$set(l), 1 & t && o.value !== e[0] && B(o, e[0]);
                },
                i(e) {
                    c || (Ae(n.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (c = !1);
                },
                d(l) {
                    l && L(t), Ue(n), e[7](null), (s = !1), r(i);
                },
            }
        );
    }
    function Ml(e, t, n) {
        let l,
            o,
            { value: c = "" } = t,
            { save: s } = t,
            { cancel: r } = t;
        function i(e) {
            s(c, e);
        }
        ce(() => {
            (o = c), l.select(), l.focus();
        });
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (c = e.value)), "save" in e && n(4, (s = e.save)), "cancel" in e && n(5, (r = e.cancel));
            }),
            [
                c,
                l,
                i,
                function (e) {
                    "Enter" === e.key && i(), "Escape" === e.key && (n(0, (c = o)), r());
                },
                s,
                r,
                () => i(!0),
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (l = e), n(1, l);
                    });
                },
                function () {
                    (c = this.value), n(0, c);
                },
            ]
        );
    }
    class _l extends Ye {
        constructor(e) {
            super(), He(this, e, Ml, Sl, a, { value: 0, save: 4, cancel: 5 });
        }
    }
    function Cl(e, t, n) {
        const l = e.slice();
        return (l[11] = t[n]), (l[13] = n), l;
    }
    function Dl(e) {
        let t,
            l,
            o,
            c,
            s,
            r = e[11] + "",
            i = e[1] && Al(),
            a = e[2] && El();
        return {
            c() {
                (t = z("div")), (l = O(r)), (o = F()), i && i.c(), (c = F()), a && a.c(), (s = F()), U(t, "class", "wx-list-item"), U(t, "data-id", e[13] + 1);
            },
            m(e, n) {
                T(e, t, n), I(t, l), I(t, o), i && i.m(t, null), I(t, c), a && a.m(t, null), I(t, s);
            },
            p(e, n) {
                1 & n && r !== (r = e[11] + "") && Y(l, r), e[1] ? i || ((i = Al()), i.c(), i.m(t, c)) : i && (i.d(1), (i = null)), e[2] ? a || ((a = El()), a.c(), a.m(t, s)) : a && (a.d(1), (a = null));
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), i && i.d(), a && a.d();
            },
        };
    }
    function Il(e) {
        let t, n;
        return (
            (t = new _l({ props: { value: e[11], save: e[7], cancel: e[8] } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    1 & n && (l.value = e[11]), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Al(e) {
        let t;
        return {
            c() {
                (t = z("i")), U(t, "class", "wx-list-icon wx-hover wxi-edit"), U(t, "data-action", "edit"), U(t, "title", "Edit");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function El(e) {
        let t;
        return {
            c() {
                (t = z("i")), U(t, "class", "wx-list-icon wx-hover wxi-delete"), U(t, "data-action", "remove"), U(t, "title", "Delete");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Tl(e) {
        let t, n, l, o;
        const c = [Il, Dl],
            s = [];
        function r(e, t) {
            return e[4] === e[13] ? 0 : 1;
        }
        return (
            (t = r(e)),
            (n = s[t] = c[t](e)),
            {
                c() {
                    n.c(), (l = q());
                },
                m(e, n) {
                    s[t].m(e, n), T(e, l, n), (o = !0);
                },
                p(e, o) {
                    let i = t;
                    (t = r(e)),
                        t === i
                            ? s[t].p(e, o)
                            : (De(),
                                Ee(s[i], 1, 1, () => {
                                    s[i] = null;
                                }),
                                Ie(),
                                (n = s[t]),
                                n ? n.p(e, o) : ((n = s[t] = c[t](e)), n.c()),
                                Ae(n, 1),
                                n.m(l.parentNode, l));
                },
                i(e) {
                    o || (Ae(n), (o = !0));
                },
                o(e) {
                    Ee(n), (o = !1);
                },
                d(e) {
                    s[t].d(e), e && L(l);
                },
            }
        );
    }
    function Ll(e) {
        let t, n;
        return (
            (t = new ct({ props: { type: "link", click: e[6], $$slots: { default: [jl] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    16392 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function jl(e) {
        let t, n;
        return {
            c() {
                (t = O("Add ")), (n = O(e[3]));
            },
            m(e, l) {
                T(e, t, l), T(e, n, l);
            },
            p(e, t) {
                8 & t && Y(n, e[3]);
            },
            d(e) {
                e && L(t), e && L(n);
            },
        };
    }
    function zl(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r = e[0],
            i = [];
        for (let t = 0; t < r.length; t += 1) i[t] = Tl(Cl(e, r, t));
        const a = (e) =>
            Ee(i[e], 1, 1, () => {
                i[e] = null;
            });
        let u = e[2] && Ll(e);
        return {
            c() {
                t = z("div");
                for (let e = 0; e < i.length; e += 1) i[e].c();
                (n = F()), u && u.c(), (l = q()), U(t, "class", "wx-list");
            },
            m(r, a) {
                T(r, t, a);
                for (let e = 0; e < i.length; e += 1) i[e].m(t, null);
                T(r, n, a), u && u.m(r, a), T(r, l, a), (o = !0), c || ((s = k(Je.call(null, t, e[5]))), (c = !0));
            },
            p(e, [n]) {
                if (407 & n) {
                    let l;
                    for (r = e[0], l = 0; l < r.length; l += 1) {
                        const o = Cl(e, r, l);
                        i[l] ? (i[l].p(o, n), Ae(i[l], 1)) : ((i[l] = Tl(o)), i[l].c(), Ae(i[l], 1), i[l].m(t, null));
                    }
                    for (De(), l = r.length; l < i.length; l += 1) a(l);
                    Ie();
                }
                e[2]
                    ? u
                        ? (u.p(e, n), 4 & n && Ae(u, 1))
                        : ((u = Ll(e)), u.c(), Ae(u, 1), u.m(l.parentNode, l))
                    : u &&
                    (De(),
                        Ee(u, 1, 1, () => {
                            u = null;
                        }),
                        Ie());
            },
            i(e) {
                if (!o) {
                    for (let e = 0; e < r.length; e += 1) Ae(i[e]);
                    Ae(u), (o = !0);
                }
            },
            o(e) {
                i = i.filter(Boolean);
                for (let e = 0; e < i.length; e += 1) Ee(i[e]);
                Ee(u), (o = !1);
            },
            d(e) {
                e && L(t), j(i, e), e && L(n), u && u.d(e), e && L(l), (c = !1), s();
            },
        };
    }
    function Nl(e, t, n) {
        let { value: l = [] } = t,
            { canEdit: o = !0 } = t,
            { canDelete: c = !0 } = t,
            { title: s = "" } = t,
            r = null;
        let i = {
            remove: function (e) {
                return n(0, (l = l.slice(0, e - 1).append(...l.slice(e)))), !1;
            },
            edit: function (e) {
                return o && n(4, (r = e - 1)), !1;
            },
        };
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (l = e.value)), "canEdit" in e && n(1, (o = e.canEdit)), "canDelete" in e && n(2, (c = e.canDelete)), "title" in e && n(3, (s = e.title));
            }),
            [
                l,
                o,
                c,
                s,
                r,
                i,
                function () {
                    n(0, (l = [...l, "http://localhost"])), n(4, (r = l.length - 1));
                },
                function (e) {
                    n(0, (l = [...l])), n(0, (l[r] = e), l), n(4, (r = null));
                },
                function () {
                    n(4, (r = null));
                },
            ]
        );
    }
    const { document: Ol } = je;
    function Fl(e) {
        let t, n, l, o, c;
        const s = e[7].default,
            r = $(s, e, e[6], null);
        return {
            c() {
                (t = F()), (n = z("div")), r && r.c(), U(n, "class", "popup svelte-12a3mjo"), G(n, "top", e[1] + "px"), G(n, "left", e[0] + "px");
            },
            m(s, i) {
                T(s, t, i), T(s, n, i), r && r.m(n, null), e[8](n), (l = !0), o || ((c = R(Ol.body, "mousedown", e[3])), (o = !0));
            },
            p(e, [t]) {
                r && r.p && (!l || 64 & t) && g(r, s, e, e[6], l ? h(s, e[6], t, null) : v(e[6]), null), (!l || 2 & t) && G(n, "top", e[1] + "px"), (!l || 1 & t) && G(n, "left", e[0] + "px");
            },
            i(e) {
                l || (Ae(r, e), (l = !0));
            },
            o(e) {
                Ee(r, e), (l = !1);
            },
            d(l) {
                l && L(t), l && L(n), r && r.d(l), e[8](null), (o = !1), c();
            },
        };
    }
    function ql(e, t, n) {
        let l,
            { $$slots: o = {}, $$scope: c } = t,
            { left: s = 0 } = t,
            { top: r = 0 } = t,
            { area: i = null } = t,
            { cancel: a } = t;
        return (
            se(() =>
                (function () {
                    const e = document.body.getBoundingClientRect(),
                        t = l.getBoundingClientRect();
                    t.right >= e.right && n(0, (s = e.right - t.width)), t.bottom >= e.bottom && n(1, (r = e.bottom - t.height - 12));
                })()
            ),
            (e.$$set = (e) => {
                "left" in e && n(0, (s = e.left)), "top" in e && n(1, (r = e.top)), "area" in e && n(4, (i = e.area)), "cancel" in e && n(5, (a = e.cancel)), "$$scope" in e && n(6, (c = e.$$scope));
            }),
            (e.$$.update = () => {
                16 & e.$$.dirty && i && (n(1, (r = i.top + i.height)), n(0, (s = i.left)));
            }),
            [
                s,
                r,
                l,
                function (e) {
                    l.contains(e.target) || (a && a(e));
                },
                i,
                a,
                c,
                o,
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (l = e), n(2, l);
                    });
                },
            ]
        );
    }
    function Rl(e) {
        let t, l, o;
        return {
            c() {
                (t = z("input")), U(t, "type", "number"), U(t, "id", e[1]), U(t, "class", "svelte-itk9nc");
            },
            m(n, c) {
                T(n, t, c), B(t, e[0]), l || ((o = R(t, "input", e[2])), (l = !0));
            },
            p(e, [n]) {
                2 & n && U(t, "id", e[1]), 1 & n && H(t.value) !== e[0] && B(t, e[0]);
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), (l = !1), o();
            },
        };
    }
    function Pl(e, t, n) {
        let { value: l = 0 } = t,
            { id: o = et() } = t;
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (l = e.value)), "id" in e && n(1, (o = e.id));
            }),
            [
                l,
                o,
                function () {
                    (l = H(this.value)), n(0, l);
                },
            ]
        );
    }
    function Kl(e) {
        let t, l, o, c, s, i, a, u, d, p, f, $, m, h, g, v, y, w, b, x, k, S;
        return {
            c() {
                (t = z("div")),
                    (l = z("div")),
                    (o = z("span")),
                    (o.textContent = "Rows per page:"),
                    (c = F()),
                    (s = z("input")),
                    (i = F()),
                    (a = z("div")),
                    (u = z("i")),
                    (d = F()),
                    (p = z("i")),
                    (f = F()),
                    ($ = z("input")),
                    (m = F()),
                    (h = z("i")),
                    (g = F()),
                    (v = z("i")),
                    (y = F()),
                    (w = z("div")),
                    (b = O("Total pages: ")),
                    (x = O(e[2])),
                    U(s, "class", "rows-per-page svelte-73e4hn"),
                    U(s, "type", "text"),
                    U(l, "class", "left"),
                    U(u, "class", "icon wxi-angle-dbl-left svelte-73e4hn"),
                    U(p, "class", "icon wxi-angle-left svelte-73e4hn"),
                    U($, "class", "active-page svelte-73e4hn"),
                    U($, "type", "text"),
                    U(h, "class", "icon wxi-angle-right svelte-73e4hn"),
                    U(v, "class", "icon wxi-angle-dbl-right svelte-73e4hn"),
                    U(a, "class", "center svelte-73e4hn"),
                    U(w, "class", "total-pages"),
                    U(t, "class", "pagination svelte-73e4hn");
            },
            m(n, r) {
                T(n, t, r),
                    I(t, l),
                    I(l, o),
                    I(l, c),
                    I(l, s),
                    B(s, e[0]),
                    I(t, i),
                    I(t, a),
                    I(a, u),
                    I(a, d),
                    I(a, p),
                    I(a, f),
                    I(a, $),
                    B($, e[1]),
                    I(a, m),
                    I(a, h),
                    I(a, g),
                    I(a, v),
                    I(t, y),
                    I(t, w),
                    I(w, b),
                    I(w, x),
                    k || ((S = [R(s, "input", e[7]), R(u, "click", e[8]), R(p, "click", e[9]), R($, "input", e[10]), R(h, "click", e[11]), R(v, "click", e[12])]), (k = !0));
            },
            p(e, [t]) {
                1 & t && s.value !== e[0] && B(s, e[0]), 2 & t && $.value !== e[1] && B($, e[1]), 4 & t && Y(x, e[2]);
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), (k = !1), r(S);
            },
        };
    }
    function Ul(e, t, n) {
        const l = re();
        let { pageSize: o = 20 } = t,
            { total: c = 0 } = t,
            { value: s = 1 } = t,
            r = 0,
            i = 0,
            a = 0;
        function u(e) {
            switch (e) {
                case "first":
                    n(1, (s = 1));
                    break;
                case "prev":
                    n(1, (s = Math.max(1, s - 1)));
                    break;
                case "next":
                    n(1, (s = Math.min(+s + 1, r)));
                    break;
                case "last":
                    n(1, (s = r));
            }
        }
        return (
            (e.$$set = (e) => {
                "pageSize" in e && n(0, (o = e.pageSize)), "total" in e && n(4, (c = e.total)), "value" in e && n(1, (s = e.value));
            }),
            (e.$$.update = () => {
                17 & e.$$.dirty && n(2, (r = Math.ceil(c / o))),
                    115 & e.$$.dirty &&
                    (n(5, (i = (s - 1) * o)),
                        n(6, (a = Math.min(s * o, c))),
                        setTimeout(() => {
                            l("change", { value: s, from: i, to: a });
                        }, 1));
            }),
            [
                o,
                s,
                r,
                u,
                c,
                i,
                a,
                function () {
                    (o = this.value), n(0, o);
                },
                () => u("first"),
                () => u("prev"),
                function () {
                    (s = this.value), n(1, s);
                },
                () => u("next"),
                () => u("last"),
            ]
        );
    }
    function Hl(e) {
        let t, l, o;
        return {
            c() {
                (t = z("input")), U(t, "type", "password"), U(t, "id", e[1]), U(t, "class", "svelte-itk9nc");
            },
            m(n, c) {
                T(n, t, c), B(t, e[0]), e[5](t), l || ((o = R(t, "input", e[4])), (l = !0));
            },
            p(e, [n]) {
                2 & n && U(t, "id", e[1]), 1 & n && t.value !== e[0] && B(t, e[0]);
            },
            i: n,
            o: n,
            d(n) {
                n && L(t), e[5](null), (l = !1), o();
            },
        };
    }
    function Yl(e, t, n) {
        let l,
            { value: o = "" } = t,
            { id: c = et() } = t,
            { focus: s = !1 } = t;
        return (
            s && ce(() => l.focus()),
            (e.$$set = (e) => {
                "value" in e && n(0, (o = e.value)), "id" in e && n(1, (c = e.id)), "focus" in e && n(3, (s = e.focus));
            }),
            [
                o,
                c,
                l,
                s,
                function () {
                    (o = this.value), n(0, o);
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (l = e), n(2, l);
                    });
                },
            ]
        );
    }
    function Bl(e) {
        let t, l, o, c, s, r, i;
        return {
            c() {
                (t = z("div")),
                    (l = z("input")),
                    (o = F()),
                    (c = z("label")),
                    (s = O(e[1])),
                    U(l, "type", "radio"),
                    (l.__value = e[2]),
                    (l.value = l.__value),
                    U(l, "id", e[3]),
                    U(l, "class", "svelte-1yal0m3"),
                    e[5][0].push(l),
                    U(c, "for", e[3]),
                    U(c, "class", "svelte-1yal0m3"),
                    U(t, "class", "svelte-1yal0m3");
            },
            m(n, a) {
                T(n, t, a), I(t, l), (l.checked = l.__value === e[0]), I(t, o), I(t, c), I(c, s), r || ((i = R(l, "change", e[4])), (r = !0));
            },
            p(e, [t]) {
                4 & t && ((l.__value = e[2]), (l.value = l.__value)), 1 & t && (l.checked = l.__value === e[0]), 2 & t && Y(s, e[1]);
            },
            i: n,
            o: n,
            d(n) {
                n && L(t), e[5][0].splice(e[5][0].indexOf(l), 1), (r = !1), i();
            },
        };
    }
    function Gl(e, t, n) {
        const l = et();
        let { label: o = "" } = t,
            { value: c = "" } = t,
            { group: s = "" } = t;
        return (
            (e.$$set = (e) => {
                "label" in e && n(1, (o = e.label)), "value" in e && n(2, (c = e.value)), "group" in e && n(0, (s = e.group));
            }),
            [
                s,
                o,
                c,
                l,
                function () {
                    (s = this.__value), n(0, s);
                },
                [[]],
            ]
        );
    }
    class Jl extends Ye {
        constructor(e) {
            super(), He(this, e, Gl, Bl, a, { label: 1, value: 2, group: 0 });
        }
    }
    function Vl(e, t, n) {
        const l = e.slice();
        return (l[3] = t[n]), l;
    }
    function Xl(e) {
        let t, n, l;
        function o(t) {
            e[2](t);
        }
        let c = { label: e[3].label, value: e[3].value };
        return (
            void 0 !== e[0] && (c.group = e[0]),
            (t = new Jl({ props: c })),
            pe.push(() => Re(t, "group", o)),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, n) {
                    Ke(t, e, n), (l = !0);
                },
                p(e, l) {
                    const o = {};
                    2 & l && (o.label = e[3].label), 2 & l && (o.value = e[3].value), !n && 1 & l && ((n = !0), (o.group = e[0]), ye(() => (n = !1))), t.$set(o);
                },
                i(e) {
                    l || (Ae(t.$$.fragment, e), (l = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (l = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Ql(e) {
        let t,
            n,
            l = e[1],
            o = [];
        for (let t = 0; t < l.length; t += 1) o[t] = Xl(Vl(e, l, t));
        const c = (e) =>
            Ee(o[e], 1, 1, () => {
                o[e] = null;
            });
        return {
            c() {
                for (let e = 0; e < o.length; e += 1) o[e].c();
                t = q();
            },
            m(e, l) {
                for (let t = 0; t < o.length; t += 1) o[t].m(e, l);
                T(e, t, l), (n = !0);
            },
            p(e, [n]) {
                if (3 & n) {
                    let s;
                    for (l = e[1], s = 0; s < l.length; s += 1) {
                        const c = Vl(e, l, s);
                        o[s] ? (o[s].p(c, n), Ae(o[s], 1)) : ((o[s] = Xl(c)), o[s].c(), Ae(o[s], 1), o[s].m(t.parentNode, t));
                    }
                    for (De(), s = l.length; s < o.length; s += 1) c(s);
                    Ie();
                }
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < l.length; e += 1) Ae(o[e]);
                    n = !0;
                }
            },
            o(e) {
                o = o.filter(Boolean);
                for (let e = 0; e < o.length; e += 1) Ee(o[e]);
                n = !1;
            },
            d(e) {
                j(o, e), e && L(t);
            },
        };
    }
    function Wl(e, t, n) {
        let { options: l = [{}] } = t,
            { value: o } = t;
        return (
            (e.$$set = (e) => {
                "options" in e && n(1, (l = e.options)), "value" in e && n(0, (o = e.value));
            }),
            [
                o,
                l,
                function (e) {
                    (o = e), n(0, o);
                },
            ]
        );
    }
    function Zl(e, t, n) {
        const l = e.slice();
        return (l[19] = t[n]), l;
    }
    const eo = (e) => ({ option: 2 & e }),
        to = (e) => ({ option: e[19] }),
        no = (e) => ({ option: 32 & e }),
        lo = (e) => ({ option: e[5] });
    function oo(e) {
        let t;
        return {
            c() {
                t = O(" ");
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            i: n,
            o: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function co(e) {
        let t;
        const n = e[12].default,
            l = $(n, e, e[15], lo);
        return {
            c() {
                l && l.c();
            },
            m(e, n) {
                l && l.m(e, n), (t = !0);
            },
            p(e, o) {
                l && l.p && (!t || 32800 & o) && g(l, n, e, e[15], t ? h(n, e[15], o, no) : v(e[15]), lo);
            },
            i(e) {
                t || (Ae(l, e), (t = !0));
            },
            o(e) {
                Ee(l, e), (t = !1);
            },
            d(e) {
                l && l.d(e);
            },
        };
    }
    function so(e) {
        let t, n;
        return (
            (t = new at({ props: { cancel: e[10], $$slots: { default: [io] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    32843 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function ro(e) {
        let t, n, l, o;
        const c = e[12].default,
            s = $(c, e, e[15], to);
        return {
            c() {
                (t = z("div")), s && s.c(), (n = F()), U(t, "class", "item svelte-dp2sk0"), U(t, "data-id", (l = e[19].id)), V(t, "selected", e[0] && e[0] === e[19].id), V(t, "navigate", e[6] && e[6].id === e[19].id);
            },
            m(e, l) {
                T(e, t, l), s && s.m(t, null), I(t, n), (o = !0);
            },
            p(e, n) {
                s && s.p && (!o || 32770 & n) && g(s, c, e, e[15], o ? h(c, e[15], n, eo) : v(e[15]), to),
                    (!o || (2 & n && l !== (l = e[19].id))) && U(t, "data-id", l),
                    3 & n && V(t, "selected", e[0] && e[0] === e[19].id),
                    66 & n && V(t, "navigate", e[6] && e[6].id === e[19].id);
            },
            i(e) {
                o || (Ae(s, e), (o = !0));
            },
            o(e) {
                Ee(s, e), (o = !1);
            },
            d(e) {
                e && L(t), s && s.d(e);
            },
        };
    }
    function io(e) {
        let t,
            n,
            l,
            o,
            c = e[1],
            s = [];
        for (let t = 0; t < c.length; t += 1) s[t] = ro(Zl(e, c, t));
        const r = (e) =>
            Ee(s[e], 1, 1, () => {
                s[e] = null;
            });
        return {
            c() {
                t = z("div");
                for (let e = 0; e < s.length; e += 1) s[e].c();
                U(t, "class", "list svelte-dp2sk0");
            },
            m(c, r) {
                T(c, t, r);
                for (let e = 0; e < s.length; e += 1) s[e].m(t, null);
                e[14](t), (n = !0), l || ((o = R(t, "click", K(e[8]))), (l = !0));
            },
            p(e, n) {
                if (32835 & n) {
                    let l;
                    for (c = e[1], l = 0; l < c.length; l += 1) {
                        const o = Zl(e, c, l);
                        s[l] ? (s[l].p(o, n), Ae(s[l], 1)) : ((s[l] = ro(o)), s[l].c(), Ae(s[l], 1), s[l].m(t, null));
                    }
                    for (De(), l = c.length; l < s.length; l += 1) r(l);
                    Ie();
                }
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(s[e]);
                    n = !0;
                }
            },
            o(e) {
                s = s.filter(Boolean);
                for (let e = 0; e < s.length; e += 1) Ee(s[e]);
                n = !1;
            },
            d(n) {
                n && L(t), j(s, n), e[14](null), (l = !1), o();
            },
        };
    }
    function ao(e) {
        let t, n, l, o, c, s, i, a, u, d;
        const p = [co, oo],
            f = [];
        function $(e, t) {
            return e[5] ? 0 : 1;
        }
        (l = $(e)), (o = f[l] = p[l](e));
        let m = e[2] && so(e);
        return {
            c() {
                (t = z("div")),
                    (n = z("div")),
                    o.c(),
                    (c = F()),
                    (s = z("i")),
                    (i = F()),
                    m && m.c(),
                    U(n, "class", "label svelte-dp2sk0"),
                    U(s, "class", "icon wxi-angle-down svelte-dp2sk0"),
                    U(t, "class", "select svelte-dp2sk0"),
                    U(t, "tabindex", "0"),
                    V(t, "active", e[2]);
            },
            m(o, r) {
                T(o, t, r), I(t, n), f[l].m(n, null), e[13](n), I(t, c), I(t, s), I(t, i), m && m.m(t, null), (a = !0), u || ((d = [R(t, "click", e[7]), R(t, "keydown", e[9])]), (u = !0));
            },
            p(e, [c]) {
                let s = l;
                (l = $(e)),
                    l === s
                        ? f[l].p(e, c)
                        : (De(),
                            Ee(f[s], 1, 1, () => {
                                f[s] = null;
                            }),
                            Ie(),
                            (o = f[l]),
                            o ? o.p(e, c) : ((o = f[l] = p[l](e)), o.c()),
                            Ae(o, 1),
                            o.m(n, null)),
                    e[2]
                        ? m
                            ? (m.p(e, c), 4 & c && Ae(m, 1))
                            : ((m = so(e)), m.c(), Ae(m, 1), m.m(t, null))
                        : m &&
                        (De(),
                            Ee(m, 1, 1, () => {
                                m = null;
                            }),
                            Ie()),
                    4 & c && V(t, "active", e[2]);
            },
            i(e) {
                a || (Ae(o), Ae(m), (a = !0));
            },
            o(e) {
                Ee(o), Ee(m), (a = !1);
            },
            d(n) {
                n && L(t), f[l].d(), e[13](null), m && m.d(), (u = !1), r(d);
            },
        };
    }
    function uo(e, t, n) {
        let l,
            o,
            c,
            s,
            r,
            i,
            a,
            { $$slots: u = {}, $$scope: d } = t,
            { value: p = null } = t,
            { options: f = [] } = t;
        function $(e) {
            let t;
            (a = a || 0 === a ? a : r), (a += e), a === f.length ? (a = 0) : a < 0 && (a = f.length - 1), n(6, (i = f[a])), (t = c.clientHeight * (a + 1) - o.clientHeight), o.scrollTo(0, t);
        }
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (p = e.value)), "options" in e && n(1, (f = e.options)), "$$scope" in e && n(15, (d = e.$$scope));
            }),
            (e.$$.update = () => {
                2055 & e.$$.dirty && (n(5, (s = f.find((e) => e.id === p))), n(11, (r = p ? f.findIndex((e) => e.id === p) : 0)), n(6, (i = l ? f[r] : null)));
            }),
            [
                p,
                f,
                l,
                o,
                c,
                s,
                i,
                function () {
                    n(2, (l = !0));
                },
                function (e) {
                    const t = Ge(e);
                    t && n(0, (p = t)), n(2, (l = null));
                },
                function (e) {
                    switch (e.code) {
                        case "Space":
                            n(2, (l = !l));
                            break;
                        case "Enter":
                            l ? (n(0, (p = i.id)), (a = null), n(2, (l = null))) : n(2, (l = !0));
                            break;
                        case "ArrowDown":
                            l ? $(1) : n(2, (l = !0));
                            break;
                        case "ArrowUp":
                            l && $(-1);
                            break;
                        case "Escape":
                            (a = null), n(2, (l = null));
                    }
                },
                function () {
                    n(2, (l = null));
                },
                r,
                u,
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (c = e), n(4, c);
                    });
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (o = e), n(3, o);
                    });
                },
                d,
            ]
        );
    }
    function po(e, t, n) {
        const l = e.slice();
        return (l[5] = t[n]), l;
    }
    function fo(e, t) {
        let n,
            l,
            o,
            c = t[5][t[1]] + "";
        return {
            key: e,
            first: null,
            c() {
                (n = z("option")), (l = O(c)), (n.__value = o = t[5].id), (n.value = n.__value), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), I(n, l);
            },
            p(e, s) {
                (t = e), 6 & s && c !== (c = t[5][t[1]] + "") && Y(l, c), 4 & s && o !== (o = t[5].id) && ((n.__value = o), (n.value = n.__value));
            },
            d(e) {
                e && L(n);
            },
        };
    }
    function $o(e) {
        let t,
            l,
            o,
            c = [],
            s = new Map(),
            r = e[2];
        const i = (e) => e[5].id;
        for (let t = 0; t < r.length; t += 1) {
            let n = po(e, r, t),
                l = i(n);
            s.set(l, (c[t] = fo(l, n)));
        }
        return {
            c() {
                t = z("select");
                for (let e = 0; e < c.length; e += 1) c[e].c();
                U(t, "class", "select svelte-1wtqgkb"), U(t, "id", e[3]), void 0 === e[0] && ve(() => e[4].call(t));
            },
            m(n, s) {
                T(n, t, s);
                for (let e = 0; e < c.length; e += 1) c[e].m(t, null);
                J(t, e[0]), l || ((o = R(t, "change", e[4])), (l = !0));
            },
            p(e, [n]) {
                6 & n && ((r = e[2]), (c = Oe(c, n, i, 1, e, r, s, t, ze, fo, null, po))), 8 & n && U(t, "id", e[3]), 5 & n && J(t, e[0]);
            },
            i: n,
            o: n,
            d(e) {
                e && L(t);
                for (let e = 0; e < c.length; e += 1) c[e].d();
                (l = !1), o();
            },
        };
    }
    function mo(e, t, n) {
        let { label: l = "label" } = t,
            { value: o = "" } = t,
            { options: c = [] } = t,
            { id: s = et() } = t;
        return (
            (e.$$set = (e) => {
                "label" in e && n(1, (l = e.label)), "value" in e && n(0, (o = e.value)), "options" in e && n(2, (c = e.options)), "id" in e && n(3, (s = e.id));
            }),
            [
                o,
                l,
                c,
                s,
                function () {
                    (o = (function (e) {
                        const t = e.querySelector(":checked") || e.options[0];
                        return t && t.__value;
                    })(this)),
                        n(0, o),
                        n(2, c);
                },
            ]
        );
    }
    function ho(e) {
        let t, l, o, c, s, i, a;
        return {
            c() {
                (t = z("div")),
                    (l = z("label")),
                    (o = O(e[1])),
                    (c = F()),
                    (s = z("input")),
                    U(l, "class", "label svelte-fo4v47"),
                    U(l, "for", e[6]),
                    U(s, "id", e[6]),
                    U(s, "class", "range svelte-fo4v47"),
                    U(s, "type", "range"),
                    U(s, "min", e[2]),
                    U(s, "max", e[3]),
                    U(s, "step", e[4]),
                    U(s, "style", e[5]),
                    U(t, "class", "layout svelte-fo4v47");
            },
            m(n, r) {
                T(n, t, r), I(t, l), I(l, o), I(t, c), I(t, s), B(s, e[0]), i || ((a = [R(s, "change", e[8]), R(s, "input", e[8])]), (i = !0));
            },
            p(e, [t]) {
                2 & t && Y(o, e[1]), 4 & t && U(s, "min", e[2]), 8 & t && U(s, "max", e[3]), 16 & t && U(s, "step", e[4]), 32 & t && U(s, "style", e[5]), 1 & t && B(s, e[0]);
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), (i = !1), r(a);
            },
        };
    }
    function go(e, t, n) {
        let { label: l = "" } = t,
            { min: o = 1 } = t,
            { max: c = 100 } = t,
            { value: s = 0 } = t,
            { step: r = 1 } = t,
            i = 0,
            a = "";
        const u = et();
        return (
            (e.$$set = (e) => {
                "label" in e && n(1, (l = e.label)), "min" in e && n(2, (o = e.min)), "max" in e && n(3, (c = e.max)), "value" in e && n(0, (s = e.value)), "step" in e && n(4, (r = e.step));
            }),
            (e.$$.update = () => {
                141 & e.$$.dirty && (n(7, (i = ((s - o) / (c - o)) * 100 + "%")), n(5, (a = `background: linear-gradient(90deg, var(--wx-primary-color) 0% ${i}, #dbdbdb ${i} 100%);`)), ("number" != typeof s || isNaN(s)) && n(0, (s = 0)));
            }),
            [
                s,
                l,
                o,
                c,
                r,
                a,
                u,
                i,
                function () {
                    (s = H(this.value)), n(0, s), n(2, o), n(3, c), n(7, i);
                },
            ]
        );
    }
    class vo extends Ye {
        constructor(e) {
            super(), He(this, e, go, ho, a, { label: 1, min: 2, max: 3, value: 0, step: 4 });
        }
    }
    function yo(e) {
        let t, l, o, c, s, r;
        return {
            c() {
                (t = z("label")), (l = z("input")), (o = F()), (c = z("span")), U(l, "type", "checkbox"), U(l, "id", e[1]), U(l, "class", "svelte-rolpl1"), U(c, "class", "slider svelte-rolpl1"), U(t, "class", "switch svelte-rolpl1");
            },
            m(n, i) {
                T(n, t, i), I(t, l), (l.checked = e[0]), I(t, o), I(t, c), s || ((r = R(l, "change", e[2])), (s = !0));
            },
            p(e, [t]) {
                2 & t && U(l, "id", e[1]), 1 & t && (l.checked = e[0]);
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), (s = !1), r();
            },
        };
    }
    function wo(e, t, n) {
        let { id: l = et() } = t,
            { checked: o } = t;
        return (
            (e.$$set = (e) => {
                "id" in e && n(1, (l = e.id)), "checked" in e && n(0, (o = e.checked));
            }),
            [
                o,
                l,
                function () {
                    (o = this.checked), n(0, o);
                },
            ]
        );
    }
    function bo(e, t, n) {
        const l = e.slice();
        return (l[3] = t[n]), l;
    }
    function xo(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i,
            a = e[3].label + "";
        function u() {
            return e[2](e[3]);
        }
        return {
            c() {
                (t = z("button")), (n = z("i")), (o = F()), (c = O(a)), (s = F()), U(n, "class", (l = "tab-icon " + e[3].icon + " svelte-vnpvx7")), U(t, "class", "tab svelte-vnpvx7"), V(t, "active", e[3].id == e[0]);
            },
            m(e, l) {
                T(e, t, l), I(t, n), I(t, o), I(t, c), I(t, s), r || ((i = R(t, "click", u)), (r = !0));
            },
            p(o, s) {
                (e = o), 2 & s && l !== (l = "tab-icon " + e[3].icon + " svelte-vnpvx7") && U(n, "class", l), 2 & s && a !== (a = e[3].label + "") && Y(c, a), 3 & s && V(t, "active", e[3].id == e[0]);
            },
            d(e) {
                e && L(t), (r = !1), i();
            },
        };
    }
    function ko(e) {
        let t,
            l = e[1],
            o = [];
        for (let t = 0; t < l.length; t += 1) o[t] = xo(bo(e, l, t));
        return {
            c() {
                t = z("div");
                for (let e = 0; e < o.length; e += 1) o[e].c();
                U(t, "class", "line svelte-vnpvx7");
            },
            m(e, n) {
                T(e, t, n);
                for (let e = 0; e < o.length; e += 1) o[e].m(t, null);
            },
            p(e, [n]) {
                if (3 & n) {
                    let c;
                    for (l = e[1], c = 0; c < l.length; c += 1) {
                        const s = bo(e, l, c);
                        o[c] ? o[c].p(s, n) : ((o[c] = xo(s)), o[c].c(), o[c].m(t, null));
                    }
                    for (; c < o.length; c += 1) o[c].d(1);
                    o.length = l.length;
                }
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), j(o, e);
            },
        };
    }
    function So(e, t, n) {
        let { options: l } = t,
            { value: o } = t;
        return (
            (e.$$set = (e) => {
                "options" in e && n(1, (l = e.options)), "value" in e && n(0, (o = e.value));
            }),
            [o, l, (e) => n(0, (o = e.id))]
        );
    }
    function Mo(e, { delay: t = 0, duration: n = 400, easing: o = l } = {}) {
        const c = +getComputedStyle(e).opacity;
        return { delay: t, duration: n, easing: o, css: (e) => "opacity: " + e * c };
    }
    function _o(e) {
        let t;
        return {
            c() {
                t = O("Cancel");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Co(e) {
        let t;
        return {
            c() {
                t = O("OK");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Do(e) {
        let t, n, l, o, c, s, r, i, a, u, d, p, f, m, y, w, b;
        const x = e[5].default,
            k = $(x, e, e[7], null);
        return (
            (u = new ct({ props: { type: "plain", click: e[2], $$slots: { default: [_o] }, $$scope: { ctx: e } } })),
            (f = new ct({ props: { type: "primary", click: e[1], $$slots: { default: [Co] }, $$scope: { ctx: e } } })),
            {
                c() {
                    (t = z("div")),
                        (n = z("div")),
                        (l = z("div")),
                        (o = O(e[0])),
                        (c = F()),
                        (s = z("div")),
                        k && k.c(),
                        (r = F()),
                        (i = z("div")),
                        (a = z("div")),
                        Pe(u.$$.fragment),
                        (d = F()),
                        (p = z("div")),
                        Pe(f.$$.fragment),
                        U(l, "class", "header svelte-8po8gp"),
                        U(s, "class", "body svelte-8po8gp"),
                        U(a, "class", "button svelte-8po8gp"),
                        U(p, "class", "button svelte-8po8gp"),
                        U(i, "class", "buttons svelte-8po8gp"),
                        U(n, "class", "confirm svelte-8po8gp"),
                        U(t, "class", "modal svelte-8po8gp"),
                        U(t, "tabindex", "0");
                },
                m($, m) {
                    T($, t, m), I(t, n), I(n, l), I(l, o), I(n, c), I(n, s), k && k.m(s, null), I(n, r), I(n, i), I(i, a), Ke(u, a, null), I(i, d), I(i, p), Ke(f, p, null), e[6](t), (y = !0), w || ((b = R(t, "keydown", e[4])), (w = !0));
                },
                p(e, [t]) {
                    (!y || 1 & t) && Y(o, e[0]), k && k.p && (!y || 128 & t) && g(k, x, e, e[7], y ? h(x, e[7], t, null) : v(e[7]), null);
                    const n = {};
                    4 & t && (n.click = e[2]), 128 & t && (n.$$scope = { dirty: t, ctx: e }), u.$set(n);
                    const l = {};
                    2 & t && (l.click = e[1]), 128 & t && (l.$$scope = { dirty: t, ctx: e }), f.$set(l);
                },
                i(e) {
                    y ||
                        (Ae(k, e),
                            Ae(u.$$.fragment, e),
                            Ae(f.$$.fragment, e),
                            ve(() => {
                                m || (m = Le(t, Mo, {}, !0)), m.run(1);
                            }),
                            (y = !0));
                },
                o(e) {
                    Ee(k, e), Ee(u.$$.fragment, e), Ee(f.$$.fragment, e), m || (m = Le(t, Mo, {}, !1)), m.run(0), (y = !1);
                },
                d(n) {
                    n && L(t), k && k.d(n), Ue(u), Ue(f), e[6](null), n && m && m.end(), (w = !1), b();
                },
            }
        );
    }
    function Io(e, t, n) {
        let l,
            { $$slots: o = {}, $$scope: c } = t,
            { title: s } = t,
            { ok: r } = t,
            { cancel: i } = t;
        return (
            ce(() => {
                l.focus();
            }),
            (e.$$set = (e) => {
                "title" in e && n(0, (s = e.title)), "ok" in e && n(1, (r = e.ok)), "cancel" in e && n(2, (i = e.cancel)), "$$scope" in e && n(7, (c = e.$$scope));
            }),
            [
                s,
                r,
                i,
                l,
                function (e) {
                    switch (e.code) {
                        case "Enter":
                            r();
                            break;
                        case "Escape":
                            i();
                    }
                },
                o,
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (l = e), n(3, l);
                    });
                },
                c,
            ]
        );
    }
    class Ao extends Ye {
        constructor(e) {
            super(), He(this, e, Io, Do, a, { title: 0, ok: 1, cancel: 2 });
        }
    }
    function Eo(e) {
        let t, n, l;
        return {
            c() {
                (t = z("input")), U(t, "class", "svelte-fvad7w");
            },
            m(o, c) {
                T(o, t, c), e[6](t), B(t, e[0]), n || ((l = [R(t, "input", e[7]), R(t, "keydown", e[5])]), (n = !0));
            },
            p(e, n) {
                1 & n && t.value !== e[0] && B(t, e[0]);
            },
            d(o) {
                o && L(t), e[6](null), (n = !1), r(l);
            },
        };
    }
    function To(e) {
        let t, n;
        return (
            (t = new Ao({ props: { title: e[1], ok: e[8], cancel: e[3], $$slots: { default: [Eo] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, [n]) {
                    const l = {};
                    2 & n && (l.title = e[1]), 5 & n && (l.ok = e[8]), 8 & n && (l.cancel = e[3]), 529 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Lo(e, t, n) {
        let l,
            { title: o } = t,
            { value: c } = t,
            { ok: s } = t,
            { cancel: r } = t;
        ce(() => {
            l.select(), l.focus();
        });
        return (
            (e.$$set = (e) => {
                "title" in e && n(1, (o = e.title)), "value" in e && n(0, (c = e.value)), "ok" in e && n(2, (s = e.ok)), "cancel" in e && n(3, (r = e.cancel));
            }),
            [
                c,
                o,
                s,
                r,
                l,
                function (e) {
                    "Escape" == e.key ? r() : "Enter" == e.key && s(c, e);
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (l = e), n(4, l);
                    });
                },
                function () {
                    (c = this.value), n(0, c);
                },
                (e) => s(c, e),
            ]
        );
    }
    class jo extends Ye {
        constructor(e) {
            super(), He(this, e, Lo, To, a, { title: 1, value: 0, ok: 2, cancel: 3 });
        }
    }
    function zo(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i,
            a,
            u,
            d = e[0].text + "";
        return {
            c() {
                (t = z("div")),
                    (n = z("div")),
                    (l = O(d)),
                    (o = F()),
                    (c = z("button")),
                    (c.textContent = "×"),
                    U(n, "class", "text svelte-1u6kko"),
                    U(c, "class", "close svelte-1u6kko"),
                    U(t, "class", (s = "message " + (e[0].type || "info") + " svelte-1u6kko")),
                    U(t, "role", "status"),
                    U(t, "aria-live", "polite");
            },
            m(s, r) {
                T(s, t, r), I(t, n), I(n, l), I(t, o), I(t, c), (i = !0), a || ((u = R(c, "click", e[1])), (a = !0));
            },
            p(e, [n]) {
                (!i || 1 & n) && d !== (d = e[0].text + "") && Y(l, d), (!i || (1 & n && s !== (s = "message " + (e[0].type || "info") + " svelte-1u6kko"))) && U(t, "class", s);
            },
            i(e) {
                i ||
                    (ve(() => {
                        r || (r = Le(t, Mo, {}, !0)), r.run(1);
                    }),
                        (i = !0));
            },
            o(e) {
                r || (r = Le(t, Mo, {}, !1)), r.run(0), (i = !1);
            },
            d(e) {
                e && L(t), e && r && r.end(), (a = !1), u();
            },
        };
    }
    function No(e, t, n) {
        let { notice: l = {} } = t;
        return (
            (e.$$set = (e) => {
                "notice" in e && n(0, (l = e.notice));
            }),
            [
                l,
                function () {
                    l.remove && l.remove();
                },
            ]
        );
    }
    class Oo extends Ye {
        constructor(e) {
            super(), He(this, e, No, zo, a, { notice: 0 });
        }
    }
    function Fo(e, t, n) {
        const l = e.slice();
        return (l[2] = t[n]), l;
    }
    function qo(e, t) {
        let n, l, o;
        return (
            (l = new Oo({ props: { notice: t[2] } })),
            {
                key: e,
                first: null,
                c() {
                    (n = q()), Pe(l.$$.fragment), (this.first = n);
                },
                m(e, t) {
                    T(e, n, t), Ke(l, e, t), (o = !0);
                },
                p(e, n) {
                    t = e;
                    const o = {};
                    2 & n && (o.notice = t[2]), l.$set(o);
                },
                i(e) {
                    o || (Ae(l.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(l.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(n), Ue(l, e);
                },
            }
        );
    }
    function Ro(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[1];
        const s = (e) => e[2].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = Fo(e, c, t),
                r = s(n);
            o.set(r, (l[t] = qo(r, n)));
        }
        return {
            c() {
                t = z("div");
                for (let e = 0; e < l.length; e += 1) l[e].c();
                U(t, "class", "area svelte-zwdqj");
            },
            m(e, o) {
                T(e, t, o);
                for (let e = 0; e < l.length; e += 1) l[e].m(t, null);
                n = !0;
            },
            p(e, [n]) {
                2 & n && ((c = e[1]), De(), (l = Oe(l, n, s, 1, e, c, o, t, Ne, qo, null, Fo)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < l.length; e += 1) l[e].d();
            },
        };
    }
    function Po(e, t, l) {
        let o,
            c = n,
            s = () => (c(), (c = p(r, (e) => l(1, (o = e)))), r);
        e.$$.on_destroy.push(() => c());
        let { data: r } = t;
        return (
            s(),
            (e.$$set = (e) => {
                "data" in e && s(l(0, (r = e.data)));
            }),
            [r, o]
        );
    }
    class Ko extends Ye {
        constructor(e) {
            super(), He(this, e, Po, Ro, a, { data: 0 });
        }
    }
    function Uo(e) {
        let t, n;
        return (
            (t = new jo({ props: { title: e[0].title, value: e[0].value, ok: e[0].resolve, cancel: e[0].reject } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    1 & n && (l.title = e[0].title), 1 & n && (l.value = e[0].value), 1 & n && (l.ok = e[0].resolve), 1 & n && (l.cancel = e[0].reject), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Ho(e) {
        let t, n;
        return (
            (t = new Ao({ props: { title: e[1].title, ok: e[1].resolve, cancel: e[1].reject, $$slots: { default: [Yo] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    2 & n && (l.title = e[1].title), 2 & n && (l.ok = e[1].resolve), 2 & n && (l.cancel = e[1].reject), 18 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Yo(e) {
        let t,
            n = e[1].message + "";
        return {
            c() {
                t = O(n);
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, l) {
                2 & l && n !== (n = e[1].message + "") && Y(t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Bo(e) {
        let t, n, l, o, c;
        const s = e[3].default,
            r = $(s, e, e[4], null);
        let i = e[0] && Uo(e),
            a = e[1] && Ho(e);
        return (
            (o = new Ko({ props: { data: e[2] } })),
            {
                c() {
                    r && r.c(), (t = F()), i && i.c(), (n = F()), a && a.c(), (l = F()), Pe(o.$$.fragment);
                },
                m(e, s) {
                    r && r.m(e, s), T(e, t, s), i && i.m(e, s), T(e, n, s), a && a.m(e, s), T(e, l, s), Ke(o, e, s), (c = !0);
                },
                p(e, [t]) {
                    r && r.p && (!c || 16 & t) && g(r, s, e, e[4], c ? h(s, e[4], t, null) : v(e[4]), null),
                        e[0]
                            ? i
                                ? (i.p(e, t), 1 & t && Ae(i, 1))
                                : ((i = Uo(e)), i.c(), Ae(i, 1), i.m(n.parentNode, n))
                            : i &&
                            (De(),
                                Ee(i, 1, 1, () => {
                                    i = null;
                                }),
                                Ie()),
                        e[1]
                            ? a
                                ? (a.p(e, t), 2 & t && Ae(a, 1))
                                : ((a = Ho(e)), a.c(), Ae(a, 1), a.m(l.parentNode, l))
                            : a &&
                            (De(),
                                Ee(a, 1, 1, () => {
                                    a = null;
                                }),
                                Ie());
                },
                i(e) {
                    c || (Ae(r, e), Ae(i), Ae(a), Ae(o.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(r, e), Ee(i), Ee(a), Ee(o.$$.fragment, e), (c = !1);
                },
                d(e) {
                    r && r.d(e), e && L(t), i && i.d(e), e && L(n), a && a.d(e), e && L(l), Ue(o, e);
                },
            }
        );
    }
    function Go(e, t, n) {
        let { $$slots: l = {}, $$scope: o } = t,
            c = null;
        let s = null;
        let r = zn([]);
        return (
            ie("wx-helpers", {
                showPrompt: function (e) {
                    return (
                        n(0, (c = { ...e })),
                        void 0 === c.value ? n(0, (c.value = ""), c) : n(0, (c.value = c.value.toString()), c),
                        new Promise((e, t) => {
                            n(
                                0,
                                (c.resolve = (t) => {
                                    n(0, (c = null)), e(t);
                                }),
                                c
                            ),
                                n(
                                    0,
                                    (c.reject = (e) => {
                                        n(0, (c = null)), t(e);
                                    }),
                                    c
                                );
                        })
                    );
                },
                showNotice: function (e) {
                    ((e = { ...e }).id = e.id || et()), (e.remove = () => r.update((t) => t.filter((t) => t.id !== e.id))), -1 != e.expire && setTimeout(e.remove, e.expire || 5e3), r.update((t) => [...t, e]);
                },
                showConfirm: function (e) {
                    return (
                        n(1, (s = { ...e })),
                        new Promise((e, t) => {
                            n(
                                1,
                                (s.resolve = (t) => {
                                    n(1, (s = null)), e(t);
                                }),
                                s
                            ),
                                n(
                                    1,
                                    (s.reject = (e) => {
                                        n(1, (s = null)), t(e);
                                    }),
                                    s
                                );
                        })
                    );
                },
            }),
            (e.$$set = (e) => {
                "$$scope" in e && n(4, (o = e.$$scope));
            }),
            [c, s, r, l, o]
        );
    }
    const Jo = (e) => ({}),
        Vo = (e) => ({ id: e[2] });
    function Xo(e) {
        let t, n, l, o, c, s;
        const r = e[4].default,
            i = $(r, e, e[3], Vo);
        return {
            c() {
                (t = z("div")), (n = z("label")), (l = O(e[0])), (o = F()), i && i.c(), U(n, "for", e[2]), U(n, "class", "svelte-p56d46"), U(t, "class", (c = b(e[1]) + " svelte-p56d46"));
            },
            m(e, c) {
                T(e, t, c), I(t, n), I(n, l), I(t, o), i && i.m(t, null), (s = !0);
            },
            p(e, [n]) {
                (!s || 1 & n) && Y(l, e[0]), i && i.p && (!s || 8 & n) && g(i, r, e, e[3], s ? h(r, e[3], n, Jo) : v(e[3]), Vo), (!s || (2 & n && c !== (c = b(e[1]) + " svelte-p56d46"))) && U(t, "class", c);
            },
            i(e) {
                s || (Ae(i, e), (s = !0));
            },
            o(e) {
                Ee(i, e), (s = !1);
            },
            d(e) {
                e && L(t), i && i.d(e);
            },
        };
    }
    function Qo(e, t, n) {
        let { $$slots: l = {}, $$scope: o } = t,
            { label: c } = t,
            { position: s = "left" } = t,
            r = et();
        return (
            (e.$$set = (e) => {
                "label" in e && n(0, (c = e.label)), "position" in e && n(1, (s = e.position)), "$$scope" in e && n(3, (o = e.$$scope));
            }),
            [c, s, r, o, l]
        );
    }
    function Wo(e) {
        let t, n;
        return (
            (t = new at({ props: { cancel: e[7], width: "unset", $$slots: { default: [Zo] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    65542 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Zo(e) {
        let t, n, l, o, c, s, i, a, u, d, p, f, $, m, h, g;
        function v(t) {
            e[13](t);
        }
        let y = { label: "Hours", max: tc };
        function w(t) {
            e[14](t);
        }
        void 0 !== e[1] && (y.value = e[1]), (u = new vo({ props: y })), pe.push(() => Re(u, "value", v));
        let b = { label: "Minutes", max: nc };
        return (
            void 0 !== e[2] && (b.value = e[2]),
            (f = new vo({ props: b })),
            pe.push(() => Re(f, "value", w)),
            {
                c() {
                    (t = z("div")),
                        (n = z("div")),
                        (l = z("input")),
                        (o = F()),
                        (c = z("div")),
                        (c.textContent = ":"),
                        (s = F()),
                        (i = z("input")),
                        (a = F()),
                        Pe(u.$$.fragment),
                        (p = F()),
                        Pe(f.$$.fragment),
                        U(l, "class", "digit svelte-1pyog1z"),
                        U(c, "class", "separator svelte-1pyog1z"),
                        U(i, "class", "digit svelte-1pyog1z"),
                        U(n, "class", "timer svelte-1pyog1z"),
                        U(t, "class", "wrapper svelte-1pyog1z");
                },
                m(r, d) {
                    T(r, t, d),
                        I(t, n),
                        I(n, l),
                        B(l, e[1]),
                        I(n, o),
                        I(n, c),
                        I(n, s),
                        I(n, i),
                        B(i, e[2]),
                        I(t, a),
                        Ke(u, t, null),
                        I(t, p),
                        Ke(f, t, null),
                        (m = !0),
                        h || ((g = [R(l, "input", e[9]), R(l, "blur", e[10]), R(i, "input", e[11]), R(i, "blur", e[12])]), (h = !0));
                },
                p(e, t) {
                    2 & t && l.value !== e[1] && B(l, e[1]), 4 & t && i.value !== e[2] && B(i, e[2]);
                    const n = {};
                    !d && 2 & t && ((d = !0), (n.value = e[1]), ye(() => (d = !1))), u.$set(n);
                    const o = {};
                    !$ && 4 & t && (($ = !0), (o.value = e[2]), ye(() => ($ = !1))), f.$set(o);
                },
                i(e) {
                    m || (Ae(u.$$.fragment, e), Ae(f.$$.fragment, e), (m = !0));
                },
                o(e) {
                    Ee(u.$$.fragment, e), Ee(f.$$.fragment, e), (m = !1);
                },
                d(e) {
                    e && L(t), Ue(u), Ue(f), (h = !1), r(g);
                },
            }
        );
    }
    function ec(e) {
        let t, n, l, o, c, s, r, i, a;
        function u(t) {
            e[8](t);
        }
        let d = { id: e[4], readonly: !0 };
        void 0 !== e[0] && (d.value = e[0]), (n = new Bt({ props: d })), pe.push(() => Re(n, "value", u));
        let p = e[3] && Wo(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (o = F()), (c = z("i")), (s = F()), p && p.c(), U(c, "class", "icon wxi-clock svelte-1pyog1z"), U(t, "class", "input svelte-1pyog1z");
            },
            m(l, u) {
                T(l, t, u), Ke(n, t, null), I(t, o), I(t, c), I(t, s), p && p.m(t, null), (r = !0), i || ((a = R(t, "click", e[6])), (i = !0));
            },
            p(e, [o]) {
                const c = {};
                !l && 1 & o && ((l = !0), (c.value = e[0]), ye(() => (l = !1))),
                    n.$set(c),
                    e[3]
                        ? p
                            ? (p.p(e, o), 8 & o && Ae(p, 1))
                            : ((p = Wo(e)), p.c(), Ae(p, 1), p.m(t, null))
                        : p &&
                        (De(),
                            Ee(p, 1, 1, () => {
                                p = null;
                            }),
                            Ie());
            },
            i(e) {
                r || (Ae(n.$$.fragment, e), Ae(p), (r = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), Ee(p), (r = !1);
            },
            d(e) {
                e && L(t), Ue(n), p && p.d(), (i = !1), a();
            },
        };
    }
    const tc = 23,
        nc = 59;
    function lc(e) {
        return (e = ((e = `${e}`.replace(/[^\d]/g, "")) < 10 ? `0${e}` : e).slice(-2));
    }
    function oc(e, t, n) {
        let { value: l = "12:20" } = t;
        const o = et(),
            c = (e, t) => (e < t ? e : t);
        let s;
        const r = l.split(":");
        let i = c(r[0], tc),
            a = c(r[1], nc);
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (l = e.value));
            }),
            (e.$$.update = () => {
                6 & e.$$.dirty && (n(1, (i = lc(i))), n(2, (a = lc(a))), n(0, (l = `${i}:${a}`)));
            }),
            [
                l,
                i,
                a,
                s,
                o,
                c,
                function () {
                    n(3, (s = !0));
                },
                function () {
                    n(3, (s = null));
                },
                function (e) {
                    (l = e), n(0, l), n(1, i), n(2, a);
                },
                function () {
                    (i = this.value), n(1, i), n(2, a);
                },
                () => n(1, (i = c(i, tc))),
                function () {
                    (a = this.value), n(2, a), n(1, i);
                },
                () => n(2, (a = c(a, nc))),
                function (e) {
                    (i = e), n(1, i), n(2, a);
                },
                function (e) {
                    (a = e), n(2, a), n(1, i);
                },
            ]
        );
    }
    function cc(e, t, n) {
        const l = e.slice();
        return (l[5] = t[n]), l;
    }
    const sc = (e) => ({ obj: 1 & e }),
        rc = (e) => ({ obj: e[5] });
    function ic(e, t) {
        let n, l, o, c;
        const s = t[4].default,
            r = $(s, t, t[3], rc),
            i =
                r ||
                (function (e) {
                    let t,
                        n,
                        l = e[5].label + "";
                    return {
                        c() {
                            (t = z("div")), (n = O(l)), U(t, "class", "content svelte-65zxki");
                        },
                        m(e, l) {
                            T(e, t, l), I(t, n);
                        },
                        p(e, t) {
                            1 & t && l !== (l = e[5].label + "") && Y(n, l);
                        },
                        d(e) {
                            e && L(t);
                        },
                    };
                })(t);
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")), i && i.c(), (l = F()), U(n, "class", "item svelte-65zxki"), U(n, "data-id", (o = t[5].id)), (this.first = n);
            },
            m(e, t) {
                T(e, n, t), i && i.m(n, null), I(n, l), (c = !0);
            },
            p(e, l) {
                (t = e), r ? r.p && (!c || 9 & l) && g(r, s, t, t[3], c ? h(s, t[3], l, sc) : v(t[3]), rc) : i && i.p && (!c || 1 & l) && i.p(t, c ? l : -1), (!c || (1 & l && o !== (o = t[5].id))) && U(n, "data-id", o);
            },
            i(e) {
                c || (Ae(i, e), (c = !0));
            },
            o(e) {
                Ee(i, e), (c = !1);
            },
            d(e) {
                e && L(n), i && i.d(e);
            },
        };
    }
    function ac(e) {
        let t,
            n,
            l,
            o,
            c = [],
            s = new Map(),
            r = e[0];
        const i = (e) => e[5].id;
        for (let t = 0; t < r.length; t += 1) {
            let n = cc(e, r, t),
                l = i(n);
            s.set(l, (c[t] = ic(l, n)));
        }
        return {
            c() {
                t = z("div");
                for (let e = 0; e < c.length; e += 1) c[e].c();
                U(t, "class", "items svelte-65zxki");
            },
            m(s, r) {
                T(s, t, r);
                for (let e = 0; e < c.length; e += 1) c[e].m(t, null);
                (n = !0), l || ((o = R(t, "click", e[1])), (l = !0));
            },
            p(e, [n]) {
                9 & n && ((r = e[0]), De(), (c = Oe(c, n, i, 1, e, r, s, t, Ne, ic, null, cc)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < r.length; e += 1) Ae(c[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < c.length; e += 1) Ee(c[e]);
                n = !1;
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < c.length; e += 1) c[e].d();
                (l = !1), o();
            },
        };
    }
    function uc(e, t, n) {
        let { $$slots: l = {}, $$scope: o } = t,
            { data: c = [] } = t,
            { click: s } = t;
        return (
            (e.$$set = (e) => {
                "data" in e && n(0, (c = e.data)), "click" in e && n(2, (s = e.click)), "$$scope" in e && n(3, (o = e.$$scope));
            }),
            [
                c,
                function (e) {
                    const t = Ge(e);
                    t && s && s(t, e);
                },
                s,
                o,
                l,
            ]
        );
    }
    function dc(e, t, n) {
        const l = e.slice();
        return (l[3] = t[n]), l;
    }
    function pc(e) {
        let t,
            n,
            l = e[3].label + "";
        return {
            c() {
                (t = O(l)), (n = F());
            },
            m(e, l) {
                T(e, t, l), T(e, n, l);
            },
            p(e, n) {
                2 & n && l !== (l = e[3].label + "") && Y(t, l);
            },
            d(e) {
                e && L(t), e && L(n);
            },
        };
    }
    function fc(e, t) {
        let n, l, o;
        function c() {
            return t[2](t[3]);
        }
        return (
            (l = new ct({ props: { shape: "square", click: c, type: t[3].id === t[0] ? "selected" : "", $$slots: { default: [pc] }, $$scope: { ctx: t } } })),
            {
                key: e,
                first: null,
                c() {
                    (n = q()), Pe(l.$$.fragment), (this.first = n);
                },
                m(e, t) {
                    T(e, n, t), Ke(l, e, t), (o = !0);
                },
                p(e, n) {
                    t = e;
                    const o = {};
                    3 & n && (o.click = c), 3 & n && (o.type = t[3].id === t[0] ? "selected" : ""), 66 & n && (o.$$scope = { dirty: n, ctx: t }), l.$set(o);
                },
                i(e) {
                    o || (Ae(l.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(l.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(n), Ue(l, e);
                },
            }
        );
    }
    function $c(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[1];
        const s = (e) => e[3].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = dc(e, c, t),
                r = s(n);
            o.set(r, (l[t] = fc(r, n)));
        }
        return {
            c() {
                t = z("div");
                for (let e = 0; e < l.length; e += 1) l[e].c();
                U(t, "class", "toggle svelte-1kwva48");
            },
            m(e, o) {
                T(e, t, o);
                for (let e = 0; e < l.length; e += 1) l[e].m(t, null);
                n = !0;
            },
            p(e, [n]) {
                3 & n && ((c = e[1]), De(), (l = Oe(l, n, s, 1, e, c, o, t, Ne, fc, null, dc)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < l.length; e += 1) l[e].d();
            },
        };
    }
    function mc(e, t, n) {
        let { options: l } = t,
            { value: o = l[0].id } = t;
        return (
            (e.$$set = (e) => {
                "options" in e && n(1, (l = e.options)), "value" in e && n(0, (o = e.value));
            }),
            [o, l, (e) => n(0, (o = e.id))]
        );
    }
    function hc(e) {
        let t;
        return {
            c() {
                (t = z("i")), U(t, "class", e[1]);
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, n) {
                2 & n && U(t, "class", e[1]);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function gc(e) {
        let t,
            l,
            o,
            c = e[1] && hc(e);
        return {
            c() {
                (t = z("button")), c && c.c(), U(t, "class", "svelte-1vaevne"), V(t, "pressed", e[0]);
            },
            m(n, s) {
                T(n, t, s), c && c.m(t, null), l || ((o = R(t, "click", e[2])), (l = !0));
            },
            p(e, [n]) {
                e[1] ? (c ? c.p(e, n) : ((c = hc(e)), c.c(), c.m(t, null))) : c && (c.d(1), (c = null)), 1 & n && V(t, "pressed", e[0]);
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), c && c.d(), (l = !1), o();
            },
        };
    }
    function vc(e, t, n) {
        let { icon: l = "" } = t,
            { checked: o = !1 } = t;
        return (
            (e.$$set = (e) => {
                "icon" in e && n(1, (l = e.icon)), "checked" in e && n(0, (o = e.checked));
            }),
            [o, l, () => n(0, (o = !o))]
        );
    }
    const yc = (e) => ({}),
        wc = (e) => ({ open: e[10] });
    function bc(e) {
        let t, l, o, c, s, i;
        const a = e[15].default,
            u = $(a, e, e[14], wc),
            d =
                u ||
                (function (e) {
                    let t, l, o, c, s, r;
                    return {
                        c() {
                            (t = z("div")), (l = z("span")), (o = O("Drop files here or\n\t\t\t\t")), (c = z("span")), (c.textContent = "select files"), U(c, "class", "action svelte-1e30chc"), U(t, "class", "dropzone svelte-1e30chc");
                        },
                        m(n, i) {
                            T(n, t, i), I(t, l), I(l, o), I(l, c), s || ((r = R(c, "click", e[10])), (s = !0));
                        },
                        p: n,
                        d(e) {
                            e && L(t), (s = !1), r();
                        },
                    };
                })(e);
        return {
            c() {
                (t = z("div")),
                    (l = z("input")),
                    (o = F()),
                    d && d.c(),
                    U(l, "type", "file"),
                    U(l, "class", "input svelte-1e30chc"),
                    U(l, "accept", e[1]),
                    (l.disabled = e[2]),
                    (l.multiple = e[3]),
                    U(t, "class", "label svelte-1e30chc"),
                    V(t, "active", e[5]);
            },
            m(n, r) {
                T(n, t, r), I(t, l), e[17](l), I(t, o), d && d.m(t, null), (c = !0), s || ((i = [R(l, "change", e[6]), R(t, "dragenter", e[8]), R(t, "dragleave", e[9]), R(t, "dragover", P(e[16])), R(t, "drop", P(e[7]))]), (s = !0));
            },
            p(e, [n]) {
                (!c || 2 & n) && U(l, "accept", e[1]),
                    (!c || 4 & n) && (l.disabled = e[2]),
                    (!c || 8 & n) && (l.multiple = e[3]),
                    u && u.p && (!c || 16384 & n) && g(u, a, e, e[14], c ? h(a, e[14], n, yc) : v(e[14]), wc),
                    32 & n && V(t, "active", e[5]);
            },
            i(e) {
                c || (Ae(d, e), (c = !0));
            },
            o(e) {
                Ee(d, e), (c = !1);
            },
            d(n) {
                n && L(t), e[17](null), d && d.d(n), (s = !1), r(i);
            },
        };
    }
    function xc(e, t, l) {
        let o,
            c = n,
            s = () => (c(), (c = p(a, (e) => l(19, (o = e)))), a);
        e.$$.on_destroy.push(() => c());
        let { $$slots: r = {}, $$scope: i } = t,
            { data: a } = t;
        s();
        let u,
            d,
            { accept: f = "" } = t,
            { disabled: $ = !1 } = t,
            { multiple: m = !0 } = t,
            { folder: h = !1 } = t,
            { uploadURL: g = "" } = t,
            { ready: v = new Promise(() => { }) } = t,
            y = 0;
        function w(e, t) {
            if (((t = t || ""), e.isFile)) e.file((e) => b(e));
            else if (e.isDirectory) {
                e.createReader().readEntries((e) =>
                    e.forEach((e) => {
                        w(e, t + e.name + "/");
                    })
                );
            }
        }
        function b(e) {
            const t = { id: et(), status: "client", name: e.name, file: e };
            m ? a.update((e) => [...e, t]) : a.set([t]);
        }
        function x() {
            const e = o.filter((e) => "client" === e.status);
            if (!e.length) return;
            const t = [];
            e.forEach((e) => {
                const n = new FormData();
                n.append("upload", e.file);
                const l = fetch(g, { method: "POST", body: n })
                    .then((e) => e.json())
                    .then(
                        (t) => {
                            (t.status = t.status || "server"), k(e.id, t);
                        },
                        () => k(e.id, { status: "error" })
                    )
                    .catch((e) => console.log(e));
                t.push(l);
            }),
                l(
                    11,
                    (v = Promise.all(t)
                        .then(() => o.filter((e) => "server" === e.status).map((e) => e.file))
                        .catch((e) => console.log(e)))
                );
        }
        function k(e, t) {
            a.update((n) => {
                const l = n.findIndex((t) => t.id === e);
                return (n[l] = { ...n[l], ...t }), n;
            });
        }
        return (
            ce(() => {
                l(4, (u.webkitdirectory = h), u);
            }),
            (e.$$set = (e) => {
                "data" in e && s(l(0, (a = e.data))),
                    "accept" in e && l(1, (f = e.accept)),
                    "disabled" in e && l(2, ($ = e.disabled)),
                    "multiple" in e && l(3, (m = e.multiple)),
                    "folder" in e && l(12, (h = e.folder)),
                    "uploadURL" in e && l(13, (g = e.uploadURL)),
                    "ready" in e && l(11, (v = e.ready)),
                    "$$scope" in e && l(14, (i = e.$$scope));
            }),
            [
                a,
                f,
                $,
                m,
                u,
                d,
                function (e) {
                    Array.from(e.target.files).forEach((e) => b(e)), x();
                },
                function (e) {
                    Array.from(e.dataTransfer.items).forEach((e) => {
                        const t = e.webkitGetAsEntry();
                        t && w(t);
                    }),
                        l(5, (d = !1)),
                        (y = 0),
                        x();
                },
                function () {
                    0 === y && l(5, (d = !0)), y++;
                },
                function () {
                    y-- , 0 === y && l(5, (d = !1));
                },
                function () {
                    u.click();
                },
                v,
                h,
                g,
                i,
                r,
                function (t) {
                    ue.call(this, e, t);
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (u = e), l(4, u);
                    });
                },
            ]
        );
    }
    function kc(e, t, n) {
        const l = e.slice();
        return (l[8] = t[n]), l;
    }
    function Sc(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i = [],
            a = new Map(),
            u = e[1];
        const d = (e) => e[8].id;
        for (let t = 0; t < u.length; t += 1) {
            let n = kc(e, u, t),
                l = d(n);
            a.set(l, (i[t] = Ic(l, n)));
        }
        return {
            c() {
                (t = z("div")), (n = z("div")), (l = z("i")), (o = F()), (c = z("div"));
                for (let e = 0; e < i.length; e += 1) i[e].c();
                U(l, "class", "wxi-close svelte-l7oe93"), U(n, "class", "header svelte-l7oe93"), U(c, "class", "list svelte-l7oe93"), U(t, "class", "layout svelte-l7oe93");
            },
            m(a, u) {
                T(a, t, u), I(t, n), I(n, l), I(t, o), I(t, c);
                for (let e = 0; e < i.length; e += 1) i[e].m(c, null);
                s || ((r = R(l, "click", e[2])), (s = !0));
            },
            p(e, t) {
                26 & t && ((u = e[1]), (i = Oe(i, t, d, 1, e, u, a, c, ze, Ic, null, kc)));
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < i.length; e += 1) i[e].d();
                (s = !1), r();
            },
        };
    }
    function Mc(e) {
        let t,
            n,
            l = e[4](e[8].file.size) + "";
        return {
            c() {
                (t = z("div")), (n = O(l)), U(t, "class", "size");
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, t) {
                2 & t && l !== (l = e[4](e[8].file.size) + "") && Y(n, l);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function _c(e) {
        let t, n, l, o, c;
        function s() {
            return e[6](e[8]);
        }
        return {
            c() {
                (t = z("i")), (n = F()), (l = z("i")), U(t, "class", "icon wxi-check svelte-l7oe93"), U(l, "class", "icon wxi-close svelte-l7oe93");
            },
            m(e, r) {
                T(e, t, r), T(e, n, r), T(e, l, r), o || ((c = R(l, "click", s)), (o = !0));
            },
            p(t, n) {
                e = t;
            },
            d(e) {
                e && L(t), e && L(n), e && L(l), (o = !1), c();
            },
        };
    }
    function Cc(e) {
        let t, n, l, o, c;
        function s() {
            return e[5](e[8]);
        }
        return {
            c() {
                (t = z("i")), (n = F()), (l = z("i")), U(t, "class", "icon wxi-alert svelte-l7oe93"), U(l, "class", "icon wxi-close svelte-l7oe93");
            },
            m(e, r) {
                T(e, t, r), T(e, n, r), T(e, l, r), o || ((c = R(l, "click", s)), (o = !0));
            },
            p(t, n) {
                e = t;
            },
            d(e) {
                e && L(t), e && L(n), e && L(l), (o = !1), c();
            },
        };
    }
    function Dc(e) {
        let t;
        return {
            c() {
                (t = z("i")), U(t, "class", "icon wxi-spin wxi-loading svelte-l7oe93");
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function Ic(e, t) {
        let n,
            l,
            o,
            c,
            s,
            r,
            i,
            a,
            u,
            d = t[8].name + "",
            p = t[8].file && Mc(t);
        function f(e, t) {
            return "client" === e[8].status ? Dc : "error" === e[8].status ? Cc : e[8].status && "server" !== e[8].status ? void 0 : _c;
        }
        let $ = f(t),
            m = $ && $(t);
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")),
                    (l = z("div")),
                    (o = F()),
                    (c = z("div")),
                    (s = O(d)),
                    (r = F()),
                    p && p.c(),
                    (i = F()),
                    (a = z("div")),
                    m && m.c(),
                    (u = F()),
                    U(l, "class", "file-icon"),
                    U(c, "class", "name svelte-l7oe93"),
                    U(a, "class", "controls svelte-l7oe93"),
                    U(n, "class", "row svelte-l7oe93"),
                    (this.first = n);
            },
            m(e, t) {
                T(e, n, t), I(n, l), I(n, o), I(n, c), I(c, s), I(n, r), p && p.m(n, null), I(n, i), I(n, a), m && m.m(a, null), I(n, u);
            },
            p(e, l) {
                (t = e),
                    2 & l && d !== (d = t[8].name + "") && Y(s, d),
                    t[8].file ? (p ? p.p(t, l) : ((p = Mc(t)), p.c(), p.m(n, i))) : p && (p.d(1), (p = null)),
                    $ === ($ = f(t)) && m ? m.p(t, l) : (m && m.d(1), (m = $ && $(t)), m && (m.c(), m.m(a, null)));
            },
            d(e) {
                e && L(n), p && p.d(), m && m.d();
            },
        };
    }
    function Ac(e) {
        let t,
            l = e[1].length && Sc(e);
        return {
            c() {
                l && l.c(), (t = q());
            },
            m(e, n) {
                l && l.m(e, n), T(e, t, n);
            },
            p(e, [n]) {
                e[1].length ? (l ? l.p(e, n) : ((l = Sc(e)), l.c(), l.m(t.parentNode, t))) : l && (l.d(1), (l = null));
            },
            i: n,
            o: n,
            d(e) {
                l && l.d(e), e && L(t);
            },
        };
    }
    function Ec(e, t, l) {
        let o,
            c = n,
            s = () => (c(), (c = p(r, (e) => l(1, (o = e)))), r);
        e.$$.on_destroy.push(() => c());
        let { data: r } = t;
        s();
        const i = ["b", "Kb", "Mb", "Gb", "Tb", "Pb", "Eb"];
        function a(e) {
            r.update((t) => t.filter((t) => t.id !== e));
        }
        return (
            (e.$$set = (e) => {
                "data" in e && s(l(0, (r = e.data)));
            }),
            [
                r,
                o,
                function () {
                    r.set([]);
                },
                a,
                function (e) {
                    let t = 0;
                    for (; e > 1024;) t++ , (e /= 1024);
                    return Math.round(100 * e) / 100 + " " + i[t];
                },
                (e) => a(e.id),
                (e) => a(e.id),
            ]
        );
    }
    function Tc(e) {
        let t, n;
        const l = e[2].default,
            o = $(l, e, e[1], null);
        return {
            c() {
                (t = z("div")), o && o.c(), U(t, "class", "bar svelte-uf9259"), U(t, "style", e[0]);
            },
            m(e, l) {
                T(e, t, l), o && o.m(t, null), (n = !0);
            },
            p(e, [c]) {
                o && o.p && (!n || 2 & c) && g(o, l, e, e[1], n ? h(l, e[1], c, null) : v(e[1]), null), (!n || 1 & c) && U(t, "style", e[0]);
            },
            i(e) {
                n || (Ae(o, e), (n = !0));
            },
            o(e) {
                Ee(o, e), (n = !1);
            },
            d(e) {
                e && L(t), o && o.d(e);
            },
        };
    }
    function Lc(e, t, n) {
        let { $$slots: l = {}, $$scope: o } = t,
            { style: c = "" } = t;
        return (
            (e.$$set = (e) => {
                "style" in e && n(0, (c = e.style)), "$$scope" in e && n(1, (o = e.$$scope));
            }),
            [c, o, l]
        );
    }
    function jc(e) {
        let t;
        return {
            c() {
                (t = z("div")), (t.innerHTML = '<h2 class="svelte-4j092g">Feature is not implemented yet</h2>');
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            i: n,
            o: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function zc(e) {
        let t,
            n,
            l,
            o =
                e[0] &&
                e[0].default &&
                (function (e) {
                    let t, n;
                    const l = e[2].default,
                        o = $(l, e, e[1], null);
                    return {
                        c() {
                            (t = z("div")), o && o.c(), U(t, "class", "wx-meta-theme svelte-1n2gzpn"), G(t, "height", "100%");
                        },
                        m(e, l) {
                            T(e, t, l), o && o.m(t, null), (n = !0);
                        },
                        p(e, t) {
                            o && o.p && (!n || 2 & t) && g(o, l, e, e[1], n ? h(l, e[1], t, null) : v(e[1]), null);
                        },
                        i(e) {
                            n || (Ae(o, e), (n = !0));
                        },
                        o(e) {
                            Ee(o, e), (n = !1);
                        },
                        d(e) {
                            e && L(t), o && o.d(e);
                        },
                    };
                })(e);
        return {
            c() {
                o && o.c(), (t = F()), (n = z("link")), U(n, "rel", "stylesheet"), U(n, "href", "https://cdn.dhtmlx.com/fonts/wxi/wx-icons.css"), U(n, "class", "svelte-1n2gzpn");
            },
            m(e, c) {
                o && o.m(e, c), T(e, t, c), I(document.head, n), (l = !0);
            },
            p(e, [t]) {
                e[0] && e[0].default && o.p(e, t);
            },
            i(e) {
                l || (Ae(o), (l = !0));
            },
            o(e) {
                Ee(o), (l = !1);
            },
            d(e) {
                o && o.d(e), e && L(t), L(n);
            },
        };
    }
    function Nc(e, t, n) {
        let { $$slots: l = {}, $$scope: c } = t;
        const s = t.$$slots;
        return (
            (e.$$set = (e) => {
                n(3, (t = o(o({}, t), y(e)))), "$$scope" in e && n(1, (c = e.$$scope));
            }),
            (t = y(t)),
            [s, c, l]
        );
    }
    const Oc = {
        Area: class extends Ye {
            constructor(e) {
                super(), He(this, e, nt, tt, a, { value: 0, id: 1, placeholder: 2 });
            }
        },
        Button: ct,
        ButtonSelect: class extends Ye {
            constructor(e) {
                super(), He(this, e, vt, gt, a, { type: 0, shape: 1, label: 2, click: 6, options: 3 });
            }
        },
        Checkbox: bt,
        ColorPicker: class extends Ye {
            constructor(e) {
                super(), He(this, e, At, It, a, { colors: 1, value: 0, id: 2, clear: 3, placeholder: 4 });
            }
        },
        Combo: class extends Ye {
            constructor(e) {
                super(), He(this, e, Rt, qt, a, { value: 0, options: 12, key: 13 });
            }
        },
        Calendar: In,
        DatePicker: class extends Ye {
            constructor(e) {
                super(), He(this, e, Ln, Tn, a, { value: 0, id: 1 });
            }
        },
        DateRangePicker: class extends Ye {
            constructor(e) {
                super(), He(this, e, Kn, Pn, a, { value: 0, id: 1 });
            }
        },
        DoubleList: class extends Ye {
            constructor(e) {
                super(), He(this, e, Jn, Gn, a, { data: 7, values: 6 });
            }
        },
        Dropdown: at,
        MultiCombo: class extends Ye {
            constructor(e) {
                super(), He(this, e, sl, cl, a, { options: 16, values: 0, key: 1 }, null, [-1, -1]);
            }
        },
        MultiSelect: class extends Ye {
            constructor(e) {
                super(), He(this, e, xl, bl, a, { options: 10, selected: 9, canEdit: 0, canDelete: 1, edit: 11, title: 2 });
            }
        },
        MultiText: class extends Ye {
            constructor(e) {
                super(), He(this, e, Nl, zl, a, { value: 0, canEdit: 1, canDelete: 2, title: 3 });
            }
        },
        Popup: class extends Ye {
            constructor(e) {
                super(), He(this, e, ql, Fl, a, { left: 0, top: 1, area: 4, cancel: 5 });
            }
        },
        Number: class extends Ye {
            constructor(e) {
                super(), He(this, e, Pl, Rl, a, { value: 0, id: 1 });
            }
        },
        Pager: class extends Ye {
            constructor(e) {
                super(), He(this, e, Ul, Kl, a, { pageSize: 0, total: 4, value: 1 });
            }
        },
        Password: class extends Ye {
            constructor(e) {
                super(), He(this, e, Yl, Hl, a, { value: 0, id: 1, focus: 3 });
            }
        },
        RadioButton: Jl,
        RadioButtonGroup: class extends Ye {
            constructor(e) {
                super(), He(this, e, Wl, Ql, a, { options: 1, value: 0 });
            }
        },
        RichSelect: class extends Ye {
            constructor(e) {
                super(), He(this, e, uo, ao, a, { value: 0, options: 1 });
            }
        },
        Select: class extends Ye {
            constructor(e) {
                super(), He(this, e, mo, $o, a, { label: 1, value: 0, options: 2, id: 3 });
            }
        },
        Slider: vo,
        Switch: class extends Ye {
            constructor(e) {
                super(), He(this, e, wo, yo, a, { id: 1, checked: 0 });
            }
        },
        Tabs: class extends Ye {
            constructor(e) {
                super(), He(this, e, So, ko, a, { options: 1, value: 0 });
            }
        },
        Text: Bt,
        Toggle: class extends Ye {
            constructor(e) {
                super(), He(this, e, mc, $c, a, { options: 1, value: 0 });
            }
        },
        TwoState: class extends Ye {
            constructor(e) {
                super(), He(this, e, vc, gc, a, { icon: 1, checked: 0 });
            }
        },
        Field: class extends Ye {
            constructor(e) {
                super(), He(this, e, Qo, Xo, a, { label: 0, position: 1 });
            }
        },
        Globals: class extends Ye {
            constructor(e) {
                super(), He(this, e, Go, Bo, a, {});
            }
        },
        Uploader: class extends Ye {
            constructor(e) {
                super(), He(this, e, xc, bc, a, { data: 0, accept: 1, disabled: 2, multiple: 3, folder: 12, uploadURL: 13, ready: 11 });
            }
        },
        UploaderList: class extends Ye {
            constructor(e) {
                super(), He(this, e, Ec, Ac, a, { data: 0 });
            }
        },
        Timepicker: class extends Ye {
            constructor(e) {
                super(), He(this, e, oc, ec, a, { value: 0 });
            }
        },
        List: class extends Ye {
            constructor(e) {
                super(), He(this, e, uc, ac, a, { data: 0, click: 2 });
            }
        },
        Toolbar: class extends Ye {
            constructor(e) {
                super(), He(this, e, Lc, Tc, a, { style: 0 });
            }
        },
        Editor: _l,
        InProgress: class extends Ye {
            constructor(e) {
                super(), He(this, e, null, jc, a, {});
            }
        },
    },
        Fc = {
            Meta: class extends Ye {
                constructor(e) {
                    super(), He(this, e, Nc, zc, a, {});
                }
            },
        };
    function qc(e) {
        let t, n;
        const l = e[2].default,
            o = $(l, e, e[3], null);
        return {
            c() {
                (t = z("div")), o && o.c(), U(t, "class", "wx-default"), G(t, "height", "100%"), G(t, "width", "100%");
            },
            m(e, l) {
                T(e, t, l), o && o.m(t, null), (n = !0);
            },
            p(e, t) {
                o && o.p && (!n || 8 & t) && g(o, l, e, e[3], n ? h(l, e[3], t, null) : v(e[3]), null);
            },
            i(e) {
                n || (Ae(o, e), (n = !0));
            },
            o(e) {
                Ee(o, e), (n = !1);
            },
            d(e) {
                e && L(t), o && o.d(e);
            },
        };
    }
    function Rc(e) {
        let t,
            n,
            l =
                e[1] &&
                e[1].default &&
                (function (e) {
                    let t, n;
                    return (
                        (t = new e[0]({ props: { $$slots: { default: [qc] }, $$scope: { ctx: e } } })),
                        {
                            c() {
                                Pe(t.$$.fragment);
                            },
                            m(e, l) {
                                Ke(t, e, l), (n = !0);
                            },
                            p(e, n) {
                                const l = {};
                                8 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                            },
                            i(e) {
                                n || (Ae(t.$$.fragment, e), (n = !0));
                            },
                            o(e) {
                                Ee(t.$$.fragment, e), (n = !1);
                            },
                            d(e) {
                                Ue(t, e);
                            },
                        }
                    );
                })(e);
        return {
            c() {
                l && l.c(), (t = q());
            },
            m(e, o) {
                l && l.m(e, o), T(e, t, o), (n = !0);
            },
            p(e, [t]) {
                e[1] && e[1].default && l.p(e, t);
            },
            i(e) {
                n || (Ae(l), (n = !0));
            },
            o(e) {
                Ee(l), (n = !1);
            },
            d(e) {
                l && l.d(e), e && L(t);
            },
        };
    }
    function Pc(e, t, n) {
        let { $$slots: l = {}, $$scope: c } = t;
        const { Meta: s } = Fc,
            r = t.$$slots;
        return (
            (e.$$set = (e) => {
                n(4, (t = o(o({}, t), y(e)))), "$$scope" in e && n(3, (c = e.$$scope));
            }),
            (t = y(t)),
            [s, r, l, c]
        );
    }
    class Kc extends Ye {
        constructor(e) {
            super(), He(this, e, Pc, Rc, a, {});
        }
    }
    class Uc {
        constructor() {
            (this._nextHandler = () => null), (this._handlers = {}), (this.exec = this.exec.bind(this));
        }
        on(e, t) {
            const n = e,
                l = this._handlers[n];
            this._handlers[n] = l
                ? function (e) {
                    Hc(l, e)(t(e));
                }
                : (e) => {
                    Hc(this._nextHandler, e, n)(t(e));
                };
        }
        exec(e, t) {
            const n = e,
                l = this._handlers[n];
            l ? l(t) : this._nextHandler(n, t);
        }
        setNext(e) {
            this._nextHandler = e;
        }
    }
    function Hc(e, t, n) {
        return (l) => {
            !1 !== l && (l && l.then ? l.then(Hc(e, t, n)) : n ? e(n, t) : e(t));
        };
    }
    class Yc {
        constructor(e) {
            (this._nextHandler = () => null), (this._dispatch = e), (this.exec = this.exec.bind(this));
        }
        exec(e, t) {
            this._dispatch(e, t), this._nextHandler && this._nextHandler(e, t);
        }
        setNext(e) {
            this._nextHandler = e;
        }
    }
    let Bc = new Date().valueOf();
    function Gc() {
        return "temp://" + Bc++;
    }
    function Jc(e, t) {
        return !(!e || !t) && e == t;
    }
    function Vc(e, t) {
        return !!e ?.find((e) => Jc(e, t));
    }
    function Xc(e, t) {
        return `${e}` + (t ? `:${t}` : "");
    }
    function Qc(e, t, n) {
        return n ? e[t] + ":" + e[n] : e[t];
    }
    const Wc = {
        label: { show: !0 },
        description: { show: !1 },
        progress: { show: !1 },
        start_date: { show: !1 },
        end_date: { show: !1 },
        users: { show: !1 },
        priority: {
            show: !1,
            values: [
                { id: 1, color: "#FF5252", label: "high" },
                { id: 2, color: "#FFC975", label: "medium" },
                { id: 3, color: "#0AB169", label: "low" },
            ],
        },
        color: { show: !1, values: ["#65D3B3", "#FFC975", "#58C3FE"] },
        cover: { show: !1 },
        attached: { show: !1 },
        menu: { show: !0 },
    },
        Zc = [
            { key: "label", type: "text", label: "Label" },
            { key: "description", type: "textarea", label: "Description" },
            { type: "combo", label: "Priority", key: "priority" },
            { type: "color", label: "Color", key: "color" },
            { type: "progress", key: "progress", label: "Progress" },
            { type: "date", key: "start_date", label: "Start date" },
            { type: "date", key: "end_date", label: "End date" },
        ],
        es = "wx-kanban-editor",
        ts = "wx-kanban-content";
    function ns(e, { id: t, before: n, columnId: l, rowId: o }) {
        const c = e.getState(),
            s = c.columnKey,
            r = c.rowKey,
            i = c.cards.findIndex((e) => Jc(e.id, t));
        if (i < 0) return;
        if (!c.cardsMap[Xc(l, o)]) return;
        if (Jc(t, n)) return;
        const a = c.cards[i],
            u = Xc(l, o) === Qc(a, s, r),
            d = Jc(l, a[s]);
        if (e.getState().areasMeta[Xc(l, o)].noFreeSpace && !d && !u) return;
        const p = c.cards.splice(i, 1)[0];
        if (((p[s] = l), r && o && (p[r] = o), n)) {
            const e = c.cards.findIndex((e) => Jc(e.id, n));
            c.cards.splice(e, 0, p);
        } else c.cards.push(p);
        const { cardsMap: f, areasMeta: $ } = e.getInnerState(c);
        e.setState({ cards: c.cards, cardsMap: f, areasMeta: $ });
    }
    function ls(e, t) {
        const n = t.card || {},
            l = t.id || n.id || Gc(),
            o = e.getState(),
            c = o.columnKey,
            s = o.rowKey,
            r = t.rowId || (s && n[s]) || o.rows[0].id,
            i = t.columnId || n[c] || o.columns[0].id;
        if (o.areasMeta[Xc(i, r)].noFreeSpace) return !1;
        const a = { [c]: i, label: "Untitled", id: l, ...n };
        s && (a[s] = r), o.cards.push(a);
        const { cardsMap: u, areasMeta: d } = e.getInnerState(o);
        e.setState({ ...o, cardsMap: u, areasMeta: d }), t.before && ns(e, { ...t, id: l }), e.in.exec("select-card", { id: l }), (t.card = a), (t.id = l);
    }
    function os(e, { id: t, card: n }) {
        const l = e.getState(),
            o = l.cards.map((e) => {
                if (Jc(e.id, t)) {
                    return { ...e, ...n };
                }
                return e;
            }),
            { cardsMap: c, areasMeta: s } = e.getInnerState({ ...l, cards: o });
        e.setState({ cards: o, cardsMap: c, areasMeta: s });
    }
    function cs(e, { id: t }) {
        e.in.exec("unselect-card", { id: t });
        const n = e.getState(),
            l = n.cards.filter((e) => !Jc(e.id, t)),
            { cardsMap: o, areasMeta: c } = e.getInnerState({ ...n, cards: l });
        e.setState({ cards: l, cardsMap: o, areasMeta: c });
    }
    function ss(e, t) {
        const n = t.id || t.column ?.id || Gc(),
            l = e.getState(),
            o = l.columns,
            c = { id: n, label: "Untitled", ...(t.column || {}) };
        o.push(c);
        const { cardsMap: s, areasMeta: r } = e.getInnerState({ ...l, columns: o });
        e.setState({ columns: o, cardsMap: s, areasMeta: r }), (t.id = n), (t.column = c);
    }
    function rs(e, { id: t, column: n }) {
        const l = e.getState().columns.map((e) => (Jc(e.id, t) ? { ...e, ...n } : e));
        e.setState({ columns: l });
    }
    function is(e, { id: t, before: n }) {
        const { columns: l } = e.getState(),
            o = l.findIndex((e) => Jc(e.id, t)),
            c = l.splice(o, 1)[0];
        if (n) {
            const e = l.findIndex((e) => Jc(e.id, n));
            l.splice(e, 0, c);
        } else l.push(c);
        e.setState({ columns: l });
    }
    function as(e, { id: t }) {
        if (t) {
            const n = e.getState(),
                l = n.columns.filter((e) => !Jc(e.id, t)),
                { cardsMap: o, areasMeta: c } = e.getInnerState({ ...n, columns: l });
            e.setState({ columns: l, cardsMap: o, areasMeta: c });
        }
    }
    function us(e, t) {
        const n = e.getState(),
            l = n.rows,
            o = t.id || t.row ?.id || Gc(),
            c = { id: o, label: "Untitled", collapsed: !1, ...(t.row || {}) };
        if ((l.push(c), !n.rowKey)) {
            const e = (n.rowKey = "rowId");
            (n.rows[0] = { id: "default", label: "Untitled" }),
                n.cards.map((t) => {
                    t[e] = "default";
                });
        }
        const { cardsMap: s, areasMeta: r } = e.getInnerState({ ...n, rows: l });
        e.setState({ rows: l, cardsMap: s, areasMeta: r }), (t.id = o), (t.row = c);
    }
    function ds(e, { id: t, row: n }) {
        const l = e.getState().rows.map((e) => (Jc(e.id, t) ? { ...e, ...n } : e));
        e.setState({ rows: l });
    }
    function ps(e, { id: t, before: n }) {
        const { rows: l, rowKey: o } = e.getState();
        if (!o) return;
        const c = l.findIndex((e) => Jc(e.id, t)),
            s = l.splice(c, 1)[0];
        if (n) {
            const e = l.findIndex((e) => Jc(e.id, n));
            l.splice(e, 0, s);
        } else l.push(s);
        e.setState({ rows: l });
    }
    function fs(e, { id: t }) {
        if (t) {
            const n = e.getState(),
                l = n.rows.filter((e) => !Jc(e.id, t)),
                { cardsMap: o, areasMeta: c } = e.getInnerState({ ...n, rows: l });
            e.setState({ rows: l, cardsMap: o, areasMeta: c });
        }
    }
    function $s(e, { dragItemsCoords: t, dropAreasCoords: n }) {
        e.setState({ dragItemsCoords: t, dropAreasCoords: n });
    }
    function ms(e, { dragItemId: t, overAreaId: n, before: l, overAreaMeta: o }) {
        e.setState({ dragItemId: t, fromAreaMeta: o, overAreaId: n, before: l, overAreaMeta: o });
    }
    function hs(e, { overAreaId: t, before: n, overAreaMeta: l }) {
        const { fromAreaMeta: o } = e.getState(),
            c = Jc(l ?.columnId, o ?.columnId),
            s = Jc(t, Xc(o ?.columnId || "", o ?.rowId));
        (!l ?.noFreeSpace || c || s) && e.setState({ overAreaId: t, before: n, overAreaMeta: l });
    }
    function gs(e) {
        const { before: t, overAreaId: n, overAreaMeta: l, dragItemId: o, selected: c } = e.getState();
        if (n && o && l) {
            const { columnId: n, rowId: s, collapsed: r } = l;
            r ||
                (c && c.length > 1
                    ? c.map((l) => {
                        e.in.exec("move-card", { id: l, columnId: n, rowId: s, before: t });
                    })
                    : e.in.exec("move-card", { id: o, columnId: n, rowId: s, before: t }));
        }
        e.setState({ dropAreaItemsCoords: null, dropAreasCoords: null, dragItemsCoords: null, dragItemId: null, fromAreaMeta: null, before: null, overAreaId: null, overAreaMeta: null });
    }
    function vs(e, { id: t, groupMode: n }) {
        const { selected: l, search: o } = e.getState();
        if (t) {
            let c = null;
            if (n) {
                if (((c = l ? [...l] : []), c.includes(t))) return void e.in.exec("unselect-card", { id: t });
                c.push(t);
            } else c = [t];
            o && e.in.exec("set-search", { value: null }), e.setState({ selected: c });
        }
    }
    function ys(e, { id: t }) {
        const n = e.getState().selected;
        if (n) {
            if (!t) return void e.setState({ selected: null });
            const l = n.filter((e) => !Jc(e, t));
            e.setState({ selected: l });
        }
    }
    function ws(e, t) {
        return e.toLowerCase().includes(t.toLowerCase());
    }
    function bs(e, { value: t, by: n }) {
        const l = e.getState(),
            o = t ?.trim(),
            c = l.cardsMeta || {};
        let s = { value: t, by: n };
        if (o) {
            (function (e) {
                return Object.keys(e.cardsMap).reduce((t, n) => t.concat(e.cardsMap[n]), []);
            })(l).map((e) => {
                const t = (c[e.id] = c[e.id] || {});
                !(function (e, t, n) {
                    return n ? ws(e[n] || "", t) : ws(e.label || "", t) || ws(e.description || "", t);
                })(e, o, n)
                    ? ((t.found = !1), (t.dimmed = !0))
                    : ((t.found = !0), (t.dimmed = !1));
            });
        } else
            Object.keys(c).forEach((e) => {
                const t = c[e];
                t && (delete t.dimmed, delete t.found);
            }),
                (s = null);
        e.setState({ cardsMeta: c, search: s });
    }
    class xs extends class {
        constructor(e) {
            (this._writable = e), (this._values = {}), (this._state = {});
        }
        setState(e) {
            const t = this._state;
            for (const n in e) t[n] ? t[n].set(e[n]) : (this._state[n] = this._wrapWritable(n, e[n]));
        }
        getState() {
            return this._values;
        }
        getReactive() {
            return this._state;
        }
        _wrapWritable(e, t) {
            const n = this._writable(t, e);
            return (
                n.subscribe((t) => {
                    this._values[e] = t;
                }),
                n
            );
        }
    } {
        constructor(e) {
            super(e), (this.in = new Uc()), (this.out = new Uc()), this.in.setNext(this.out.exec), this._initStructure();
            const t = {
                "add-card": ls,
                "update-card": os,
                "move-card": ns,
                "delete-card": cs,
                "add-column": ss,
                "update-column": rs,
                "move-column": is,
                "delete-column": as,
                "add-row": us,
                "update-row": ds,
                "move-row": ps,
                "delete-row": fs,
                "before-drag": $s,
                "drag-start": ms,
                "drag-move": hs,
                "drag-end": gs,
                "set-search": bs,
                "select-card": vs,
                "unselect-card": ys,
            };
            this._setHandlers(t);
        }
        init(e) {
            const { cards: t = [], columns: n = [], rows: l, columnKey: o = "column", rowKey: c = "", ...s } = e,
                r = this._normalizeCards(t),
                i = n.map((e) => ({ ...e })),
                a = (l || [{ id: "" }]).map((e) => ({ ...e })),
                { cardsMap: u, areasMeta: d } = this.getInnerState({ cards: r, columns: i, rows: a, columnKey: o, rowKey: c }),
                { cardShape: p, editorShape: f } = this._normalizeShapes({ ...e, cards: r }),
                $ = { ...s, cards: r, columns: i, columnKey: o, rowKey: c, rows: a, cardsMap: u, areasMeta: d, cardShape: p, editorShape: f };
            this.setState($);
        }
        getInnerState(e) {
            const { cards: t, rows: n, columns: l, columnKey: o, rowKey: c } = e,
                s = {},
                r = {};
            if (!o) return { cardsMap: r, areasMeta: s };
            l.map((e) => {
                c &&
                    n.map((t) => {
                        const n = Xc(e.id, t.id);
                        (s[n] = { columnId: e.id, rowId: t.id, limit: ("object" == typeof e.limit && e.limit[t.id]) || 0, cardsCount: 0, strictLimit: !!e.strictLimit, collapsed: e.collapsed }), (r[n] = []);
                    });
                let t = 0;
                e.limit && (t = "object" == typeof e.limit ? Object.keys(e.limit).reduce((t, n) => t + e.limit[n], 0) : e.limit),
                    (s[e.id] = { columnId: e.id, limit: t, cardsCount: 0, strictLimit: !!e.strictLimit, collapsed: e.collapsed }),
                    (r[e.id] = []);
            }),
                t.map((e) => {
                    const t = Qc(e, o, c);
                    r[t] ?.push(e), c && r[e[o]] ?.push(e);
                });
            for (const e in s) {
                const t = s[e],
                    n = r[e];
                (t.cardsCount = n.length), (t.isOverLimit = !!t.limit && n.length > t.limit), (t.noFreeSpace = t.strictLimit && !!t.limit && n.length >= t.limit);
            }
            return (
                o &&
                l.forEach((e) => {
                    s[e.id].noFreeSpace &&
                        c &&
                        n.forEach((t) => {
                            s[Xc(e.id, t.id)].noFreeSpace = !0;
                        });
                }),
                { cardsMap: r, areasMeta: s }
            );
        }
        _setHandlers(e) {
            Object.keys(e).forEach((t) => {
                this.in.on(t, (n) => e[t](this, n));
            });
        }
        _initStructure() {
            this.setState({
                columnKey: "column",
                rowKey: "",
                columns: [],
                rows: [],
                cards: [],
                cardsMap: {},
                cardsMeta: {},
                cardShape: Wc,
                editorShape: Zc,
                areasMeta: {},
                dropAreaItemsCoords: null,
                dropAreasCoords: null,
                dragItemsCoords: null,
                fromAreaMeta: null,
                overAreaMeta: null,
                dragItemId: null,
                before: null,
                overAreaId: null,
                selected: null,
                search: null,
            });
        }
        _normalizeCards(e) {
            return e.map((e) => {
                const t = e.id || Gc();
                return { ...e, id: t };
            });
        }
        _normalizeShapes(e) {
            const { cardShape: t = Wc } = e;
            for (const e in t) !0 === t[e] && (t[e] = { show: !0 });
            const n = Object.keys(t).reduce((e, n) => {
                const l = Wc[n];
                return (e[n] = l ? { ...l, ...t[n] } : t[n]), e;
            }, {}),
                l = (
                    e.editorShape ||
                    Zc.filter((e) => {
                        const t = n[e.key];
                        return t && t ?.show;
                    })
                ).map((e) => {
                    const t = n[e.key];
                    return (
                        t && "string" == typeof e.key && (t.values && !e.values && (e.values = t.values), e.config && (t.config = e.config), ("options" in e || "colors" in e) && (e.values = e.options || e.colors)), (e.id = e.id || Gc()), e
                    );
                });
            return { cardShape: n, editorShape: l };
        }
    }
    function ks(e, t) {
        return e >= t[0] && e <= t[1];
    }
    function Ss(e, t, n) {
        const l = t.x - n.x,
            o = t.y - n.y;
        return { x: e.x - l, y: e.y - o };
    }
    function Ms(e) {
        const t = {};
        if (((t.target = e.target), "touches" in e)) {
            const n = e.touches[0];
            (t.touches = e.touches), (t.clientX = n.clientX), (t.clientY = n.clientY);
        } else (t.clientX = e.clientX), (t.clientY = e.clientY);
        return t;
    }
    function _s(e, t = "data-id") {
        let n = !e.tagName && e.target ? e.target : e;
        for (; n;) {
            if (n.getAttribute) {
                if (n.getAttribute(t)) return n;
            }
            n = n.parentNode;
        }
        return null;
    }
    function Cs(e, t = "data-id") {
        const n = _s(e, t);
        return n
            ? (function (e) {
                if ("string" == typeof e) {
                    const t = 1 * e;
                    if (!isNaN(t)) return t;
                }
                return e;
            })(n.getAttribute(t))
            : null;
    }
    function Ds(e, t) {
        if (t.readonly) return;
        let n, l;
        const o = e;
        let c, s, r, i, a, u, d;
        const p = t.onAction,
            { data: f } = t.api.getStores();
        function $(e) {
            if ((s && clearTimeout(s), c)) {
                const t = c.getBoundingClientRect(),
                    n = { x: c.scrollLeft, y: c.scrollTop },
                    l = 50;
                e.clientX > t.width + t.left - l && c.scrollTo(n.x + l, n.y),
                    e.clientX < t.left + l && c.scrollTo(n.x - l, n.y),
                    e.clientY > t.height + t.top - l && c.scrollTo(n.x, n.y + l),
                    e.clientY < t.top + l && c.scrollTo(n.x, n.y - l),
                    (s = setTimeout(() => {
                        $(e);
                    }, 100));
            }
        }
        function m(e) {
            const t = {},
                n = d.find((t) =>
                    (function (e, t) {
                        const { x: n, y: l } = e,
                            o = ks(n, [t.x, t.right]),
                            c = ks(l, [t.y, t.bottom]);
                        return o && c;
                    })(e, t)
                ) ?.id;
            if (n) {
                const e = f.getState().areasMeta;
                (t.overAreaMeta = e[n]), (t.overAreaId = n);
            }
            if (n) {
                const l = u[n];
                t.before = l.find((t) => ks(e.y, [t.y, t.bottom])) ?.id;
            }
            return t;
        }
        function h(e) {
            e.preventDefault(), e.stopPropagation();
            const t = Ms(e);
            $(t);
            const { dragItemId: l, overAreaId: o, before: s, overAreaMeta: u, selected: h } = f.getState();
            if (!i) return;
            const g = i.scroll,
                v = c.scrollLeft,
                y = c.scrollTop,
                w = { x: t.clientX + (v - g.x), y: t.clientY + (y - g.y) },
                b = { x: t.clientX, y: t.clientY };
            if (
                !l &&
                (function (e, t, n = 5) {
                    return Math.abs(t.x - e.x) > n || Math.abs(t.y - e.y) > n;
                })(i, b)
            ) {
                if (t.touches && t.touches.length > 1) return;
                p("before-drag", { dragItemsCoords: a, dropAreasCoords: d });
                const e = i.itemId,
                    l = a[e];
                (r = Ss(b, i, l)), n.classList.add("dragged-card");
                const o = document.querySelector(".wx-portal");
                h && h.length > 1 && o.style.setProperty("--wx-dragged-cards-count", JSON.stringify(`+${h.length}`)),
                    o.appendChild(n),
                    (n.style.position = "fixed"),
                    (n.style.left = r.x + "px"),
                    (n.style.top = r.y + "px"),
                    document.body.classList.add("wx-ondrag");
                const c = m(w);
                p("drag-start", { dragItemId: e, ...c });
            }
            if (l) {
                const e = a[l];
                (r = Ss(b, i, e)), (n.style.left = r.x + "px"), (n.style.top = r.y + "px");
                const t = m(w),
                    c = { overAreaId: o, before: s, overAreaMeta: u };
                o !== t.overAreaId && ((c.overAreaId = t.overAreaId), (c.overAreaMeta = t.overAreaMeta)), s !== t.before && (c.before = t.before), p("drag-move", c);
            }
        }
        function g() {
            if ((document.removeEventListener(l.move, h), document.removeEventListener(l.up, g), n.parentNode)) {
                document.querySelector(".wx-portal").removeChild(n);
            }
            document.body.classList.remove("wx-ondrag");
            document.querySelector(".wx-portal").style.removeProperty("--wx-dragged-cards-count"), (n = null), s && clearTimeout(s);
            const { dragItemId: e } = f.getState();
            e && p("drag-end", null);
        }
        function v(e) {
            const t = Ms(e);
            if ((t.touches && t.touches.length > 1) || ("button" in e && 0 !== e.button)) return;
            const s = _s(t.target, "data-drag-item");
            if (((c = o.querySelector(`[data-kanban-id="${ts}"]`)), s)) {
                const r = Cs(s, "data-drag-item"),
                    p = Cs(t.target, "data-drag-item"),
                    $ = f.getState().selected,
                    m = $ && $.length > 1 ? [...$, r] : r,
                    v = (function (e, t) {
                        const n = Array.from(e.querySelectorAll("[data-drop-area]")),
                            l = Array.isArray(t) ? t : [t],
                            o = e.querySelector(`[data-drag-item='${l[l.length - 1]}']`) ?.offsetHeight || 300,
                            c = {},
                            s = [],
                            r = n.reduce((e, t) => {
                                const n = JSON.parse(JSON.stringify(t.getBoundingClientRect())),
                                    r = t.getAttribute("data-drop-area"),
                                    i = Array.from(t.querySelectorAll("[data-drag-item]")),
                                    a = [],
                                    u = i.reduce((e, t) => {
                                        const n = JSON.parse(JSON.stringify(t.getBoundingClientRect())),
                                            o = t.getAttribute("data-drag-item"),
                                            s = e[e.length - 1] ?.bottom ?? n.y,
                                            r = { ...n, y: s, id: o };
                                        return (c[o] = r), e.push(r), Vc(l, o) || a.push(o), e;
                                    }, []),
                                    d = a.map((e, t) => ({ ...u[t], id: e })),
                                    p = t.offsetParent;
                                if (t.offsetTop + t.offsetHeight >= p.scrollHeight) {
                                    const e = 30;
                                    (n.bottom += o + e), (n.height += o + e);
                                }
                                return s.push({ ...n, id: r }), (e[r] = d), e;
                            }, {});
                        return { dragItemsCoords: c, dropAreasCoords: s, dropAreaItemsCoords: r };
                    })(o, m);
                (a = v.dragItemsCoords),
                    (d = v.dropAreasCoords),
                    (u = v.dropAreaItemsCoords),
                    (n = s.cloneNode(!0)),
                    (l = "touches" in e ? { up: "touchend", move: "touchmove" } : { up: "mouseup", move: "mousemove" }),
                    document.addEventListener(l.move, h),
                    document.addEventListener(l.up, g),
                    (i = { x: t.clientX, y: t.clientY, itemId: r, areaId: p, scroll: { x: c.scrollLeft, y: c.scrollTop } });
            }
        }
        return (
            o.addEventListener("mousedown", v),
            o.addEventListener("touchstart", v),
            {
                destroy() {
                    o.removeEventListener("mousedown", v), o.removeEventListener("touchstart", v);
                },
            }
        );
    }
    function Is(e, t) {
        if (t.readonly) return;
        const { api: n } = t;
        let l;
        const o = (e) => {
            l = e.target;
        },
            c = (e) => {
                const t = Cs(l, "data-drag-item"),
                    o = Cs(l, "data-kanban-id"),
                    c = e.metaKey || e.ctrlKey,
                    s = e.shiftKey;
                l === e.target &&
                    o !== es &&
                    (function (e) {
                        const { itemId: t, groupMode: n, rangeMode: l, api: o } = e,
                            { cardsMap: c, columnKey: s } = o.getState(),
                            { selected: r } = o.getState();
                        if (!t && r ?.length) return void o.exec("unselect-card", { id: null });
                        if (l && r ?.length) {
                            const e = o.getCard(t),
                                n = o.getCard(r[r.length - 1]);
                            if (
                                (function (e, t, n) {
                                    if (!e || !t || !n) return !1;
                                    return Jc(e[n], t[n]);
                                })(e, n, s)
                            ) {
                                const l = Object.keys(c)
                                    .filter((t) => t.split(":")[0] === e[s])
                                    .reduce((e, t) => {
                                        const n = c[t];
                                        return e.concat(n);
                                    }, []),
                                    i = l.findIndex((e) => Jc(e.id, t)),
                                    a = l.findIndex((e) => Jc(e.id, n ?.id)),
                                    u = Math.min(i, a),
                                    d = Math.max(i, a);
                                return void l.slice(u, d + 1).forEach((e) => {
                                    -1 === r.indexOf(e.id) && o.exec("select-card", { id: e.id, groupMode: !0 });
                                });
                            }
                        }
                        o.exec("select-card", { id: t, groupMode: n });
                    })({ itemId: t, groupMode: c, rangeMode: s, api: n });
            };
        return (
            e.addEventListener("mousedown", o),
            e.addEventListener("mouseup", c),
            {
                destroy() {
                    e.removeEventListener("mousedown", o), e.removeEventListener("mouseup", c);
                },
            }
        );
    }
    function As(e, t) {
        if (t ?.readonly) return;
        const n = t ?.onAction;
        let l = t ?.inFocus || !1;
        function o(e) {
            if (l) {
                const t = e.ctrlKey || e.metaKey,
                    l = e.code.replace("Key", "").toLowerCase();
                n("keydown", { hotkey: `${t ? "ctrl+" : ""}${l}` });
            }
        }
        function c(t) {
            (l = (function (e, t) {
                let n = e;
                for (; n;) {
                    if (n === t) return !0;
                    n = n ?.parentNode;
                }
                return null;
            })(t.target, e)),
                n("set-focus", { inFocus: l });
        }
        return (
            document.addEventListener("keydown", o),
            document.addEventListener("mousedown", c),
            {
                destroy: () => {
                    document.removeEventListener("keydown", o), document.removeEventListener("mousedown", c);
                },
            }
        );
    }
    const Es = (e) => e.getFullYear() + "." + Ls(e.getMonth() + 1) + "." + Ls(e.getDate()),
        Ts = (e, t = "en") => {
            const n = (function (e = "en", t = "long") {
                const n = new Intl.DateTimeFormat(e, { month: t, timeZone: "UTC" });
                return [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map((e) => new Date(`1970-${e < 10 ? `0${e}` : e}-01T00:00:00+00:00`)).map((e) => n.format(e));
            })(t, "short"),
                l = e.getMonth(),
                o = e.getDate();
            return `${n[l]} ${o}`;
        };
    function Ls(e) {
        const t = e.toString();
        return 1 === t.length ? "0" + t : t;
    }
    function js(e) {
        switch (e) {
            case "jpg":
            case "jpeg":
            case "gif":
            case "png":
            case "bmp":
            case "tiff":
            case "pcx":
            case "svg":
            case "ico":
                return !0;
            default:
                return !1;
        }
    }
    function zs(e, t) {
        const { container: n, at: l, position: o = "top", align: c = "start" } = t,
            s =
                (function (e) {
                    if ("string" == typeof e) return document.querySelector(e);
                    return e;
                })(n) || document.body;
        if (l) {
            e.style.position = "absolute";
            const t = l.getBoundingClientRect();
            (e.style.top = `${t[o]}px`), (e.style.left = `${t["start" === c ? "left" : "right"]}px`);
        }
        return (
            s.appendChild(e),
            {
                destroy() {
                    e.remove();
                },
            }
        );
    }
    function Ns(e) {
        let t = e;
        return {
            data: t,
            _: (e) => t[e] || e,
            __(e, n) {
                const l = t[e];
                return (l && l[n]) || n;
            },
            getGroup(e) {
                return (t) => this.__(e, t);
            },
            extend(e) {
                return (this.data = t = { ...t, ...e }), this;
            },
        };
    }
    const Os = {
        lang: "en",
        __dates: {
            months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
            monthsShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            days: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
        },
        wx: { Today: "Today", Clear: "Clear", Close: "Close" },
        kanban: {
            Save: "Save",
            Close: "Close",
            Delete: "Delete",
            Name: "Name",
            Description: "Description",
            Type: "Type",
            "Start date": "Start date",
            "End date": "End date",
            Result: "Result",
            "No results": "No results",
            Search: "Search",
            "Search in": "Search in",
            "Add new row": "Add new row",
            "Add new column": "Add new column",
            "Add new card": "Add new card",
            "Edit card": "Edit card",
            Edit: "Edit",
            Everywhere: "Everywhere",
            Label: "Label",
            Status: "Status",
            Color: "Color",
            Date: "Date",
            Untitled: "Untitled",
            Rename: "Rename",
            "Move up": "Move up",
            "Move down": "Move down",
            "Move left": "Move left",
            "Move right": "Move right",
        },
    },
        Fs = ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月"],
        qs = {
            lang: "cn",
            __dates: { months: Fs, monthsShort: Fs, days: ["周日", "周一", "周二", "周三", "周四", "周五", "周六"] },
            wx: { Today: "今天", Clear: "清除", Close: "關閉" },
            kanban: {
                Save: "保存",
                Close: "關閉",
                Delete: "删除",
                Name: "名称",
                Description: "描述",
                Type: "類型",
                "Start Date": "開始日期",
                "End Date": "結束日期",
                Result: "結果",
                "No results": "没有結果",
                Search: "搜索",
                "Search in": "搜索",
                "Add new row": "添加新行",
                "Add new column": "添加新列",
                "Add new card": "添加新卡",
                "Edit card": "編輯卡片",
                Edit: "編輯",
                Everywhere: "所有範圍",
                Label: "標籤",
                Status: "狀態",
                Color: "颜色",
                Date: "日期",
                Untitled: "無標題",
                Rename: "重命名",
                "Move up": "上移",
                "Move down": "下移",
                "Move left": "左移",
                "Move right": "右移",
            },
        },
        Rs = {
            lang: "ru",
            __dates: {
                months: ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"],
                monthsShort: ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"],
                days: ["Вск", "Пон", "Втр", "Срд", "Чет", "Птн", "Суб"],
            },
            wx: { Today: "Сегодня", Clear: "Очистить", Close: "Закрыть" },
            kanban: {
                Save: "Сохранить",
                Close: "Закрыть",
                Delete: "Удалить",
                Name: "Задача",
                Description: "Описание",
                Type: "Тип",
                "Start Date": "Дата выполнения",
                "End Date": "Дата окончания",
                Result: "Результат",
                "No results": "Ничего не найдено",
                Search: "Поиск",
                "Search in": "Поиск",
                "Add new row": "Новая строка",
                "Add new column": "Новая колонка",
                "Add new card": "Добавить новую карточку",
                "Edit card": "Редактировать карточку",
                Edit: "Редактировать",
                Everywhere: "Всюду",
                Label: "Заголовок",
                Status: "Статус",
                Color: "Цвет",
                Date: "Дата",
                Untitled: "Без названия",
                Rename: "Переименовать",
                "Move up": "Переместить вверх",
                "Move down": "Переместить вниз",
                "Move left": "Переместить влево",
                "Move right": "Переместить вправо",
            },
        };
    function Ps(e) {
        let t, n, l, o;
        return {
            c() {
                (t = z("i")), U(t, "class", (n = "icon wxi wxi-" + e[0] + " svelte-qe130a")), G(t, "font-size", e[1] + "px"), V(t, "wxi-spin", e[2]), V(t, "clickable", e[4]);
            },
            m(n, c) {
                T(n, t, c),
                    l ||
                    ((o = R(t, "click", function () {
                        i(e[3]) && e[3].apply(this, arguments);
                    })),
                        (l = !0));
            },
            p(l, o) {
                (e = l), 1 & o && n !== (n = "icon wxi wxi-" + e[0] + " svelte-qe130a") && U(t, "class", n), 2 & o && G(t, "font-size", e[1] + "px"), 5 & o && V(t, "wxi-spin", e[2]), 17 & o && V(t, "clickable", e[4]);
            },
            d(e) {
                e && L(t), (l = !1), o();
            },
        };
    }
    function Ks(e) {
        let t, n, l;
        return {
            c() {
                (t = N("svg")),
                    (n = N("path")),
                    U(n, "d", (l = e[5][e[0]].path)),
                    U(t, "xmlns", "http://www.w3.org/2000/svg"),
                    U(t, "viewBox", "0 0 24 24"),
                    U(t, "width", e[1]),
                    U(t, "height", e[1]),
                    U(t, "class", "icon svelte-qe130a"),
                    U(t, "fill", "currentColor"),
                    V(t, "clickable", e[4]),
                    V(t, "wxi-spin", e[2]);
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, o) {
                1 & o && l !== (l = e[5][e[0]].path) && U(n, "d", l), 2 & o && U(t, "width", e[1]), 2 & o && U(t, "height", e[1]), 16 & o && V(t, "clickable", e[4]), 4 & o && V(t, "wxi-spin", e[2]);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Us(e) {
        let t;
        function l(e, t) {
            return e[5][e[0]] ? Ks : Ps;
        }
        let o = l(e),
            c = o(e);
        return {
            c() {
                c.c(), (t = q());
            },
            m(e, n) {
                c.m(e, n), T(e, t, n);
            },
            p(e, [n]) {
                o === (o = l(e)) && c ? c.p(e, n) : (c.d(1), (c = o(e)), c && (c.c(), c.m(t.parentNode, t)));
            },
            i: n,
            o: n,
            d(e) {
                c.d(e), e && L(t);
            },
        };
    }
    function Hs(e, t, n) {
        let { name: l } = t,
            { size: o = 20 } = t,
            { spin: c = !1 } = t,
            { click: s = null } = t,
            { clickable: r = !!s } = t;
        return (
            (e.$$set = (e) => {
                "name" in e && n(0, (l = e.name)), "size" in e && n(1, (o = e.size)), "spin" in e && n(2, (c = e.spin)), "click" in e && n(3, (s = e.click)), "clickable" in e && n(4, (r = e.clickable));
            }),
            [
                l,
                o,
                c,
                s,
                r,
                {
                    "arrow-right": { path: "M4,11V13H16L10.5,18.5L11.92,19.92L19.84,12L11.92,4.08L10.5,5.5L16,11H4Z" },
                    "arrow-left": { path: "M20,11V13H8L13.5,18.5L12.08,19.92L4.16,12L12.08,4.08L13.5,5.5L8,11H20Z" },
                    "arrow-up": { path: "M13,20H11V8L5.5,13.5L4.08,12.08L12,4.16L19.92,12.08L18.5,13.5L13,8V20Z" },
                    "arrow-down": { path: "M11,4H13V16L18.5,10.5L19.92,11.92L12,19.84L4.08,11.92L5.5,10.5L11,16V4Z" },
                },
            ]
        );
    }
    class Ys extends Ye {
        constructor(e) {
            super(), He(this, e, Hs, Us, a, { name: 0, size: 1, spin: 2, click: 3, clickable: 4 });
        }
    }
    function Bs(e) {
        let t, n, l, o, c, s, r, i, a;
        function u(e, t) {
            return e[5] ? Js : Gs;
        }
        n = new Ys({ props: { name: e[0].collapsed ? "angle-right" : "angle-down", clickable: !0, click: e[13] } });
        let d = u(e),
            p = d(e),
            f = !e[5] && e[3] && Vs(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = z("div")), p.c(), (c = F()), f && f.c(), (s = q()), U(t, "class", "label-icon svelte-8m1jct"), U(o, "class", "label-text svelte-8m1jct");
            },
            m(u, d) {
                T(u, t, d), Ke(n, t, null), T(u, l, d), T(u, o, d), p.m(o, null), T(u, c, d), f && f.m(u, d), T(u, s, d), (r = !0), i || ((a = R(o, "dblclick", e[15])), (i = !0));
            },
            p(e, t) {
                const l = {};
                1 & t && (l.name = e[0].collapsed ? "angle-right" : "angle-down"),
                    n.$set(l),
                    d === (d = u(e)) && p ? p.p(e, t) : (p.d(1), (p = d(e)), p && (p.c(), p.m(o, null))),
                    !e[5] && e[3]
                        ? f
                            ? (f.p(e, t), 40 & t && Ae(f, 1))
                            : ((f = Vs(e)), f.c(), Ae(f, 1), f.m(s.parentNode, s))
                        : f &&
                        (De(),
                            Ee(f, 1, 1, () => {
                                f = null;
                            }),
                            Ie());
            },
            i(e) {
                r || (Ae(n.$$.fragment, e), Ae(f), (r = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), Ee(f), (r = !1);
            },
            d(e) {
                e && L(t), Ue(n), e && L(l), e && L(o), p.d(), e && L(c), f && f.d(e), e && L(s), (i = !1), a();
            },
        };
    }
    function Gs(e) {
        let t,
            n = e[0].label + "";
        return {
            c() {
                t = O(n);
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, l) {
                1 & l && n !== (n = e[0].label + "") && Y(t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Js(e) {
        let t, n, l, o;
        return {
            c() {
                (t = z("input")), U(t, "type", "text"), U(t, "class", "input svelte-8m1jct"), (t.value = n = e[0].label);
            },
            m(n, c) {
                T(n, t, c), l || ((o = [R(t, "input", e[16]), R(t, "keypress", e[17]), R(t, "blur", e[14]), k(tr.call(null, t))]), (l = !0));
            },
            p(e, l) {
                1 & l && n !== (n = e[0].label) && t.value !== n && (t.value = n);
            },
            d(e) {
                e && L(t), (l = !1), r(o);
            },
        };
    }
    function Vs(e) {
        let t, n, l, o;
        n = new Ys({ props: { name: "dots-h", click: e[18] } });
        let c = e[6] && Xs(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (l = F()), c && c.c(), U(t, "class", "menu svelte-8m1jct");
            },
            m(s, r) {
                T(s, t, r), Ke(n, t, null), I(t, l), c && c.m(t, null), e[22](t), (o = !0);
            },
            p(e, n) {
                e[6]
                    ? c
                        ? (c.p(e, n), 64 & n && Ae(c, 1))
                        : ((c = Xs(e)), c.c(), Ae(c, 1), c.m(t, null))
                    : c &&
                    (De(),
                        Ee(c, 1, 1, () => {
                            c = null;
                        }),
                        Ie());
            },
            i(e) {
                o || (Ae(n.$$.fragment, e), Ae(c), (o = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), Ee(c), (o = !1);
            },
            d(l) {
                l && L(t), Ue(n), c && c.d(), e[22](null);
            },
        };
    }
    function Xs(e) {
        let t, n, l, o, c, s;
        return (
            (n = new e[10]({ props: { cancel: e[21], width: "auto", $$slots: { default: [Ws] }, $$scope: { ctx: e } } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), U(t, "class", "menu-wrap svelte-8m1jct"), G(t, "left", e[8].offsetLeft + "px"), G(t, "top", e[7].offsetTop + e[7].offsetHeight + "px");
                },
                m(r, i) {
                    T(r, t, i), Ke(n, t, null), (o = !0), c || ((s = k((l = zs.call(null, t, { container: e[4] })))), (c = !0));
                },
                p(e, c) {
                    const s = {};
                    64 & c && (s.cancel = e[21]),
                        16777730 & c && (s.$$scope = { dirty: c, ctx: e }),
                        n.$set(s),
                        (!o || 256 & c) && G(t, "left", e[8].offsetLeft + "px"),
                        (!o || 128 & c) && G(t, "top", e[7].offsetTop + e[7].offsetHeight + "px"),
                        l && i(l.update) && 16 & c && l.update.call(null, { container: e[4] });
                },
                i(e) {
                    o || (Ae(n.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(t), Ue(n), (c = !1), s();
                },
            }
        );
    }
    function Qs(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r = e[28].label + "";
        return (
            (n = new Ys({ props: { name: e[28].icon } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = z("span")), (c = O(r)), U(o, "class", "svelte-8m1jct"), U(t, "class", "menu-item svelte-8m1jct");
                },
                m(e, r) {
                    T(e, t, r), Ke(n, t, null), I(t, l), I(t, o), I(o, c), (s = !0);
                },
                p(e, t) {
                    const l = {};
                    268435456 & t && (l.name = e[28].icon), n.$set(l), (!s || 268435456 & t) && r !== (r = e[28].label + "") && Y(c, r);
                },
                i(e) {
                    s || (Ae(n.$$.fragment, e), (s = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (s = !1);
                },
                d(e) {
                    e && L(t), Ue(n);
                },
            }
        );
    }
    function Ws(e) {
        let t, n;
        return (
            (t = new e[11]({
                props: {
                    click: e[19],
                    data: [
                        { icon: "edit", label: e[12]("Rename"), id: 1 },
                        { icon: "arrow-up", label: e[12]("Move up"), id: e[9] > 0 ? 3 : "wx-list-disabled" },
                        { icon: "arrow-down", label: e[12]("Move down"), id: e[9] < e[1].length - 1 ? 4 : "wx-list-disabled" },
                        { icon: "delete", label: e[12]("Delete"), id: 2 },
                    ],
                    $$slots: { default: [Qs, ({ obj: e }) => ({ 28: e }), ({ obj: e }) => (e ? 268435456 : 0)] },
                    $$scope: { ctx: e },
                },
            })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    514 & n &&
                        (l.data = [
                            { icon: "edit", label: e[12]("Rename"), id: 1 },
                            { icon: "arrow-up", label: e[12]("Move up"), id: e[9] > 0 ? 3 : "wx-list-disabled" },
                            { icon: "arrow-down", label: e[12]("Move down"), id: e[9] < e[1].length - 1 ? 4 : "wx-list-disabled" },
                            { icon: "delete", label: e[12]("Delete"), id: 2 },
                        ]),
                        285212672 & n && (l.$$scope = { dirty: n, ctx: e }),
                        t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Zs(e) {
        let t;
        const n = e[20].default,
            l = $(n, e, e[24], null);
        return {
            c() {
                l && l.c();
            },
            m(e, n) {
                l && l.m(e, n), (t = !0);
            },
            p(e, o) {
                l && l.p && (!t || 16777216 & o) && g(l, n, e, e[24], t ? h(n, e[24], o, null) : v(e[24]), null);
            },
            i(e) {
                t || (Ae(l, e), (t = !0));
            },
            o(e) {
                Ee(l, e), (t = !1);
            },
            d(e) {
                l && l.d(e);
            },
        };
    }
    function er(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i = e[2] && Bs(e),
            a = !e[0].collapsed && Zs(e);
        return {
            c() {
                (t = z("div")),
                    (n = z("div")),
                    i && i.c(),
                    (l = F()),
                    (o = z("div")),
                    (c = F()),
                    (s = z("div")),
                    a && a.c(),
                    U(o, "class", "label-line svelte-8m1jct"),
                    U(n, "class", "label svelte-8m1jct"),
                    U(s, "class", "content svelte-8m1jct"),
                    U(t, "class", "row svelte-8m1jct"),
                    V(t, "collapsed", e[0].collapsed);
            },
            m(u, d) {
                T(u, t, d), I(t, n), i && i.m(n, null), I(n, l), I(n, o), e[23](n), I(t, c), I(t, s), a && a.m(s, null), (r = !0);
            },
            p(e, [o]) {
                e[2]
                    ? i
                        ? (i.p(e, o), 4 & o && Ae(i, 1))
                        : ((i = Bs(e)), i.c(), Ae(i, 1), i.m(n, l))
                    : i &&
                    (De(),
                        Ee(i, 1, 1, () => {
                            i = null;
                        }),
                        Ie()),
                    e[0].collapsed
                        ? a &&
                        (De(),
                            Ee(a, 1, 1, () => {
                                a = null;
                            }),
                            Ie())
                        : a
                            ? (a.p(e, o), 1 & o && Ae(a, 1))
                            : ((a = Zs(e)), a.c(), Ae(a, 1), a.m(s, null)),
                    1 & o && V(t, "collapsed", e[0].collapsed);
            },
            i(e) {
                r || (Ae(i), Ae(a), (r = !0));
            },
            o(e) {
                Ee(i), Ee(a), (r = !1);
            },
            d(n) {
                n && L(t), i && i.d(), e[23](null), a && a.d();
            },
        };
    }
    function tr(e) {
        e.focus();
    }
    function nr(e, t, n) {
        let l,
            { $$slots: o = {}, $$scope: c } = t;
        const { Dropdown: s, List: r } = Oc;
        let { row: i = { id: "default", label: "default", collapsed: !1 } } = t,
            { rows: a = [] } = t,
            { collapsable: u = !0 } = t,
            { edit: d = !0 } = t,
            { contentEl: p } = t;
        const f = ae("wx-i18n").getGroup("kanban"),
            $ = re();
        let m,
            h,
            g,
            v = !1,
            y = null;
        function w() {
            v && (null == y ? void 0 : y.trim()) && $("action", { action: "update-row", data: { id: i.id, row: { label: y } } }), n(5, (v = !1)), (y = null);
        }
        function b() {
            d && n(5, (v = !0));
        }
        function x(e) {
            var t;
            const n = null === (t = a["up" === e ? l - 1 : l + 2]) || void 0 === t ? void 0 : t.id;
            $("action", { action: "move-row", data: { id: i.id, before: n } });
        }
        return (
            (e.$$set = (e) => {
                "row" in e && n(0, (i = e.row)),
                    "rows" in e && n(1, (a = e.rows)),
                    "collapsable" in e && n(2, (u = e.collapsable)),
                    "edit" in e && n(3, (d = e.edit)),
                    "contentEl" in e && n(4, (p = e.contentEl)),
                    "$$scope" in e && n(24, (c = e.$$scope));
            }),
            (e.$$.update = () => {
                3 & e.$$.dirty && n(9, (l = a.findIndex((e) => e.id === i.id)));
            }),
            [
                i,
                a,
                u,
                d,
                p,
                v,
                m,
                h,
                g,
                l,
                s,
                r,
                f,
                function () {
                    n(0, (i.collapsed = !(null == i ? void 0 : i.collapsed)), i);
                },
                w,
                b,
                function (e) {
                    y = e.target.value;
                },
                function (e) {
                    13 === e.charCode && w();
                },
                function () {
                    n(6, (m = !0));
                },
                function (e) {
                    1 === e && b(), 2 === e && $("action", { action: "delete-row", data: { id: i.id } }), 3 === e && x("up"), 4 === e && x("down"), n(6, (m = null));
                },
                o,
                () => n(6, (m = !1)),
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (g = e), n(8, g);
                    });
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (h = e), n(7, h);
                    });
                },
                c,
            ]
        );
    }
    class lr extends Ye {
        constructor(e) {
            super(), He(this, e, nr, er, a, { row: 0, rows: 1, collapsable: 2, edit: 3, contentEl: 4 });
        }
    }
    function or(e) {
        let t;
        return {
            c() {
                t = O(e[2]);
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, n) {
                4 & n && Y(t, e[2]);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function cr(e) {
        let t,
            n = e[0].label + "";
        return {
            c() {
                t = O(n);
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, l) {
                1 & l && n !== (n = e[0].label + "") && Y(t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function sr(e) {
        let t, n, l;
        return {
            c() {
                (t = z("img")), d(t.src, (n = e[0].path)) || U(t, "src", n), U(t, "alt", (l = e[0].label)), U(t, "class", "svelte-1oh2jxi");
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, o) {
                1 & o && !d(t.src, (n = e[0].path)) && U(t, "src", n), 1 & o && l !== (l = e[0].label) && U(t, "alt", l);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function rr(e) {
        let t;
        function l(e, t) {
            return e[0].path ? sr : e[1] ? cr : or;
        }
        let o = l(e),
            c = o(e);
        return {
            c() {
                (t = z("div")), c.c(), U(t, "class", "user svelte-1oh2jxi");
            },
            m(e, n) {
                T(e, t, n), c.m(t, null);
            },
            p(e, [n]) {
                o === (o = l(e)) && c ? c.p(e, n) : (c.d(1), (c = o(e)), c && (c.c(), c.m(t, null)));
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), c.d();
            },
        };
    }
    function ir(e, t, n) {
        let l,
            { data: o = { label: "", path: "" } } = t,
            { noTransform: c = !1 } = t;
        return (
            (e.$$set = (e) => {
                "data" in e && n(0, (o = e.data)), "noTransform" in e && n(1, (c = e.noTransform));
            }),
            (e.$$.update = () => {
                1 & e.$$.dirty &&
                    n(
                        2,
                        (l = o.label
                            .split(" ")
                            .map((e) => e[0])
                            .join(""))
                    );
            }),
            [o, c, l]
        );
    }
    class ar extends Ye {
        constructor(e) {
            super(), He(this, e, ir, rr, a, { data: 0, noTransform: 1 });
        }
    }
    function ur(e, t, n) {
        const l = e.slice();
        return (l[5] = t[n]), l;
    }
    function dr(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[0].users;
        const s = (e) => e[5].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = ur(e, c, t),
                r = s(n);
            o.set(r, (l[t] = pr(r, n)));
        }
        return {
            c() {
                for (let e = 0; e < l.length; e += 1) l[e].c();
                t = q();
            },
            m(e, o) {
                for (let t = 0; t < l.length; t += 1) l[t].m(e, o);
                T(e, t, o), (n = !0);
            },
            p(e, n) {
                1 & n && ((c = e[0].users), De(), (l = Oe(l, n, s, 1, e, c, o, t.parentNode, Ne, pr, t, ur)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                for (let t = 0; t < l.length; t += 1) l[t].d(e);
                e && L(t);
            },
        };
    }
    function pr(e, t) {
        let n, l, o;
        return (
            (l = new ar({ props: { data: t[5], noTransform: "$total" === t[5].id } })),
            {
                key: e,
                first: null,
                c() {
                    (n = q()), Pe(l.$$.fragment), (this.first = n);
                },
                m(e, t) {
                    T(e, n, t), Ke(l, e, t), (o = !0);
                },
                p(e, n) {
                    t = e;
                    const o = {};
                    1 & n && (o.data = t[5]), 1 & n && (o.noTransform = "$total" === t[5].id), l.$set(o);
                },
                i(e) {
                    o || (Ae(l.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(l.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(n), Ue(l, e);
                },
            }
        );
    }
    function fr(e) {
        let t,
            n,
            l,
            o,
            c,
            s = e[0].attached + "";
        return (
            (n = new Ys({ props: { name: "paperclip" } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = O(s)), U(t, "class", "attached svelte-xhy9ls");
                },
                m(e, s) {
                    T(e, t, s), Ke(n, t, null), I(t, l), I(t, o), (c = !0);
                },
                p(e, t) {
                    (!c || 1 & t) && s !== (s = e[0].attached + "") && Y(o, s);
                },
                i(e) {
                    c || (Ae(n.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (c = !1);
                },
                d(e) {
                    e && L(t), Ue(n);
                },
            }
        );
    }
    function $r(e) {
        let t, n, l, o, c, s;
        n = new Ys({ props: { name: "calendar" } });
        let r = e[0].startDate && mr(e),
            i = e[0].endDate && e[0].startDate && hr(),
            a = e[0].endDate && gr(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (l = F()), r && r.c(), (o = F()), i && i.c(), (c = F()), a && a.c(), U(t, "class", "date svelte-xhy9ls");
            },
            m(e, u) {
                T(e, t, u), Ke(n, t, null), I(t, l), r && r.m(t, null), I(t, o), i && i.m(t, null), I(t, c), a && a.m(t, null), (s = !0);
            },
            p(e, n) {
                e[0].startDate ? (r ? r.p(e, n) : ((r = mr(e)), r.c(), r.m(t, o))) : r && (r.d(1), (r = null)),
                    e[0].endDate && e[0].startDate ? i || ((i = hr()), i.c(), i.m(t, c)) : i && (i.d(1), (i = null)),
                    e[0].endDate ? (a ? a.p(e, n) : ((a = gr(e)), a.c(), a.m(t, null))) : a && (a.d(1), (a = null));
            },
            i(e) {
                s || (Ae(n.$$.fragment, e), (s = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), (s = !1);
            },
            d(e) {
                e && L(t), Ue(n), r && r.d(), i && i.d(), a && a.d();
            },
        };
    }
    function mr(e) {
        let t,
            n,
            l = e[0].startDate + "";
        return {
            c() {
                (t = z("span")), (n = O(l)), U(t, "class", "date-value svelte-xhy9ls");
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, t) {
                1 & t && l !== (l = e[0].startDate + "") && Y(n, l);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function hr(e) {
        let t;
        return {
            c() {
                t = O("-");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function gr(e) {
        let t,
            n,
            l = e[0].endDate + "";
        return {
            c() {
                (t = z("span")), (n = O(l)), U(t, "class", "date-value svelte-xhy9ls");
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, t) {
                1 & t && l !== (l = e[0].endDate + "") && Y(n, l);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function vr(e) {
        let t,
            n,
            l,
            o,
            c,
            s = e[0].users && dr(e),
            r = e[0].attached && fr(e),
            i = (e[0].endDate || e[0].startDate) && $r(e);
        return {
            c() {
                (t = z("div")), (n = z("div")), s && s.c(), (l = F()), r && r.c(), (o = F()), i && i.c(), U(n, "class", "users svelte-xhy9ls"), U(t, "class", "footer svelte-xhy9ls"), V(t, "with-content", !!Object.keys(e[0]).length);
            },
            m(e, a) {
                T(e, t, a), I(t, n), s && s.m(n, null), I(t, l), r && r.m(t, null), I(t, o), i && i.m(t, null), (c = !0);
            },
            p(e, [l]) {
                e[0].users
                    ? s
                        ? (s.p(e, l), 1 & l && Ae(s, 1))
                        : ((s = dr(e)), s.c(), Ae(s, 1), s.m(n, null))
                    : s &&
                    (De(),
                        Ee(s, 1, 1, () => {
                            s = null;
                        }),
                        Ie()),
                    e[0].attached
                        ? r
                            ? (r.p(e, l), 1 & l && Ae(r, 1))
                            : ((r = fr(e)), r.c(), Ae(r, 1), r.m(t, o))
                        : r &&
                        (De(),
                            Ee(r, 1, 1, () => {
                                r = null;
                            }),
                            Ie()),
                    e[0].endDate || e[0].startDate
                        ? i
                            ? (i.p(e, l), 1 & l && Ae(i, 1))
                            : ((i = $r(e)), i.c(), Ae(i, 1), i.m(t, null))
                        : i &&
                        (De(),
                            Ee(i, 1, 1, () => {
                                i = null;
                            }),
                            Ie()),
                    1 & l && V(t, "with-content", !!Object.keys(e[0]).length);
            },
            i(e) {
                c || (Ae(s), Ae(r), Ae(i), (c = !0));
            },
            o(e) {
                Ee(s), Ee(r), Ee(i), (c = !1);
            },
            d(e) {
                e && L(t), s && s.d(), r && r.d(), i && i.d();
            },
        };
    }
    function yr(e, t, n) {
        let l,
            { cardFields: o } = t,
            { cardShape: c } = t;
        const s = ae("wx-i18n")._;
        return (
            (e.$$set = (e) => {
                "cardFields" in e && n(1, (o = e.cardFields)), "cardShape" in e && n(2, (c = e.cardShape));
            }),
            (e.$$.update = () => {
                6 & e.$$.dirty &&
                    n(
                        0,
                        (l = (function (e, t) {
                            var n, l;
                            let o = {};
                            const { show: c } = null == t ? void 0 : t.users,
                                r = e.users;
                            if (c && r) {
                                const e = r.reduce((e, n) => {
                                    const l = t.users.values.find((e) => e.id === n);
                                    return l && e.push(l), e;
                                }, []);
                                let n = e.map((e) => {
                                    let t = e.label || "";
                                    return Object.assign(Object.assign({}, e), { label: t, id: e.id || Gc() });
                                });
                                const l = 2;
                                e.length > l && ((n = n.splice(0, l)), n.push({ label: "+" + (e.length - n.length), id: "$total" })), (null == n ? void 0 : n.length) && (o.users = n);
                            }
                            const { show: i } = t.start_date || {},
                                { show: a } = t.end_date || {};
                            let { end_date: u, start_date: d } = e;
                            return (
                                (i || a) && (d && (o.startDate = Ts(d, s("lang"))), u && (o.endDate = Ts(u, s("lang")))),
                                (null === (n = null == t ? void 0 : t.attached) || void 0 === n ? void 0 : n.show) && (null === (l = e.attached) || void 0 === l ? void 0 : l.length) && (o.attached = e.attached.length),
                                o
                            );
                        })(o, c))
                    );
            }),
            [l, o, c]
        );
    }
    class wr extends Ye {
        constructor(e) {
            super(), He(this, e, yr, vr, a, { cardFields: 1, cardShape: 2 });
        }
    }
    function br(e, t, n) {
        const l = e.slice();
        return (l[3] = t[n]), l;
    }
    function xr(e) {
        let t;
        function n(e, t) {
            return "priority" === e[3].type ? Sr : kr;
        }
        let l = n(e),
            o = l(e);
        return {
            c() {
                o.c(), (t = q());
            },
            m(e, n) {
                o.m(e, n), T(e, t, n);
            },
            p(e, c) {
                l === (l = n(e)) && o ? o.p(e, c) : (o.d(1), (o = l(e)), o && (o.c(), o.m(t.parentNode, t)));
            },
            d(e) {
                o.d(e), e && L(t);
            },
        };
    }
    function kr(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r = e[3].value + "",
            i = e[3] ?.label && Mr(e);
        return {
            c() {
                (t = z("div")), i && i.c(), (n = F()), (l = z("span")), (o = O(r)), (c = F()), U(l, "class", "value"), U(t, "class", (s = "field " + e[3].css + " svelte-x922rc"));
            },
            m(e, s) {
                T(e, t, s), i && i.m(t, null), I(t, n), I(t, l), I(l, o), I(t, c);
            },
            p(e, l) {
                e[3] ?.label ? (i ? i.p(e, l) : ((i = Mr(e)), i.c(), i.m(t, n))) : i && (i.d(1), (i = null)), 1 & l && r !== (r = e[3].value + "") && Y(o, r), 1 & l && s !== (s = "field " + e[3].css + " svelte-x922rc") && U(t, "class", s);
            },
            d(e) {
                e && L(t), i && i.d();
            },
        };
    }
    function Sr(e) {
        let t,
            n,
            l,
            o,
            c,
            s,

            r,
            i = e[3].value + "";
        return {
            c() {
                (t = z("div")),
                    (n = z("div")),
                    (l = F()),
                    (o = z("span")),
                    (c = O(i)),
                    (s = F()),
                    U(n, "class", "priority-background svelte-x922rc"),
                    G(n, "background", e[3].color),
                    U(o, "class", "priority-label svelte-x922rc"),
                    U(t, "class", (r = "field " + e[3].type + " svelte-x922rc")),
                    G(t, "color", e[3].color);
            },
            m(e, r) {
                T(e, t, r), I(t, n), I(t, l), I(t, o), I(o, c), I(t, s);
            },
            p(e, l) {
                1 & l && G(n, "background", e[3].color), 1 & l && i !== (i = e[3].value + "") && Y(c, i), 1 & l && r !== (r = "field " + e[3].type + " svelte-x922rc") && U(t, "class", r), 1 & l && G(t, "color", e[3].color);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Mr(e) {
        let t,
            n,
            l,
            o = e[3].label + "";
        return {
            c() {
                (t = z("span")), (n = O(o)), (l = O(":")), U(t, "class", "label");
            },
            m(e, o) {
                T(e, t, o), I(t, n), I(t, l);
            },
            p(e, t) {
                1 & t && o !== (o = e[3].label + "") && Y(n, o);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function _r(e) {
        let t,
            n = e[3].value && xr(e);
        return {
            c() {
                n && n.c(), (t = q());
            },
            m(e, l) {
                n && n.m(e, l), T(e, t, l);
            },
            p(e, l) {
                e[3].value ? (n ? n.p(e, l) : ((n = xr(e)), n.c(), n.m(t.parentNode, t))) : n && (n.d(1), (n = null));
            },
            d(e) {
                n && n.d(e), e && L(t);
            },
        };
    }
    function Cr(e) {
        let t,
            l = e[0],
            o = [];
        for (let t = 0; t < l.length; t += 1) o[t] = _r(br(e, l, t));
        return {
            c() {
                t = z("div");
                for (let e = 0; e < o.length; e += 1) o[e].c();
                U(t, "class", "header svelte-x922rc");
            },
            m(e, n) {
                T(e, t, n);
                for (let e = 0; e < o.length; e += 1) o[e].m(t, null);
            },
            p(e, [n]) {
                if (1 & n) {
                    let c;
                    for (l = e[0], c = 0; c < l.length; c += 1) {
                        const s = br(e, l, c);
                        o[c] ? o[c].p(s, n) : ((o[c] = _r(s)), o[c].c(), o[c].m(t, null));
                    }
                    for (; c < o.length; c += 1) o[c].d(1);
                    o.length = l.length;
                }
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), j(o, e);
            },
        };
    }
    function Dr(e, t, n) {
        let l,
            { cardFields: o } = t,
            { cardShape: c } = t;
        return (
            (e.$$set = (e) => {
                "cardFields" in e && n(1, (o = e.cardFields)), "cardShape" in e && n(2, (c = e.cardShape));
            }),
            (e.$$.update = () => {
                6 & e.$$.dirty &&
                    n(
                        0,
                        (l = (function (e, t) {
                            var n, l;
                            let o = [];
                            if (null == t ? void 0 : t.priority.show) {
                                const c = null === (l = null === (n = null == t ? void 0 : t.priority) || void 0 === n ? void 0 : n.values) || void 0 === l ? void 0 : l.find((t) => t.id === e.priority);
                                c && o.push({ type: "priority", value: c.label, color: c.color });
                            }
                            const c = t.headerFields;
                            if (c) {
                                const t = c.reduce((t, n) => (e[n.key] && t.push({ value: e[n.key], label: n.label, css: n.css }), t), []);
                                t && o.push(...t);
                            }
                            return o;
                        })(o, c))
                    );
            }),
            [l, o, c]
        );
    }
    class Ir extends Ye {
        constructor(e) {
            super(), He(this, e, Dr, Cr, a, { cardFields: 1, cardShape: 2 });
        }
    }
    function Ar(e) {
        let t, n;
        return {
            c() {
                (t = z("div")), (n = O(e[2])), U(t, "class", "value svelte-g4699o");
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, t) {
                4 & t && Y(n, e[2]);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Er(e) {
        let t,
            l,
            o,
            c,
            s,
            r,
            i,
            a = e[1] && Ar(e);
        return {
            c() {
                (t = z("div")),
                    (l = z("div")),
                    (o = O(e[0])),
                    (c = F()),
                    (s = z("div")),
                    (r = z("div")),
                    (i = F()),
                    a && a.c(),
                    U(l, "class", "label svelte-g4699o"),
                    U(r, "class", "progress svelte-g4699o"),
                    U(r, "style", e[3]),
                    U(s, "class", "wrap svelte-g4699o"),
                    U(t, "class", "layout svelte-g4699o");
            },
            m(e, n) {
                T(e, t, n), I(t, l), I(l, o), I(t, c), I(t, s), I(s, r), I(s, i), a && a.m(s, null);
            },
            p(e, [t]) {
                1 & t && Y(o, e[0]), 8 & t && U(r, "style", e[3]), e[1] ? (a ? a.p(e, t) : ((a = Ar(e)), a.c(), a.m(s, null))) : a && (a.d(1), (a = null));
            },
            i: n,
            o: n,
            d(e) {
                e && L(t), a && a.d();
            },
        };
    }
    function Tr(e, t, n) {
        let { label: l = "" } = t,
            { min: o = 0 } = t,
            { max: c = 100 } = t,
            { value: s = 0 } = t,
            { showValue: r = !0 } = t,
            i = "0",
            a = "";
        return (
            (e.$$set = (e) => {
                "label" in e && n(0, (l = e.label)), "min" in e && n(4, (o = e.min)), "max" in e && n(5, (c = e.max)), "value" in e && n(6, (s = e.value)), "showValue" in e && n(1, (r = e.showValue));
            }),
            (e.$$.update = () => {
                116 & e.$$.dirty && (n(2, (i = Math.floor(((s - o) / (c - o)) * 100) + "%")), n(3, (a = `background: linear-gradient(90deg, var(--wx-primary-color) 0% ${i}, #dbdbdb ${i} 100%);`)));
            }),
            [l, r, i, a, o, c, s]
        );
    }
    class Lr extends Ye {
        constructor(e) {
            super(), He(this, e, Tr, Er, a, { label: 0, min: 4, max: 5, value: 6, showValue: 1 });
        }
    }
    function jr(e) {
        let t;
        return {
            c() {
                (t = z("div")), U(t, "class", "color rounded svelte-1ju4xal"), G(t, "background", e[0].color);
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, n) {
                1 & n && G(t, "background", e[0].color);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function zr(e) {
        let t, n, l;
        return {
            c() {
                (t = z("div")), (n = z("img")), d(n.src, (l = e[2])) || U(n, "src", l), U(n, "alt", ""), U(n, "class", "svelte-1ju4xal"), U(t, "class", "field image svelte-1ju4xal"), V(t, "rounded", !(e[1].color ?.show && e[0].color));
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, o) {
                4 & o && !d(n.src, (l = e[2])) && U(n, "src", l), 3 & o && V(t, "rounded", !(e[1].color ?.show && e[0].color));
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Nr(e) {
        let t,
            n,
            l = e[0].label + "";
        return {
            c() {
                (t = z("span")), (n = O(l));
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, t) {
                1 & t && l !== (l = e[0].label + "") && Y(n, l);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Or(e) {
        let t, n, l, o;
        n = new Ys({ props: { name: "dots-v", click: e[7] } });
        let c = e[3] && Fr(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (l = F()), c && c.c(), U(t, "class", "menu svelte-1ju4xal");
            },
            m(e, s) {
                T(e, t, s), Ke(n, t, null), I(t, l), c && c.m(t, null), (o = !0);
            },
            p(e, n) {
                e[3]
                    ? c
                        ? (c.p(e, n), 8 & n && Ae(c, 1))
                        : ((c = Fr(e)), c.c(), Ae(c, 1), c.m(t, null))
                    : c &&
                    (De(),
                        Ee(c, 1, 1, () => {
                            c = null;
                        }),
                        Ie());
            },
            i(e) {
                o || (Ae(n.$$.fragment, e), Ae(c), (o = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), Ee(c), (o = !1);
            },
            d(e) {
                e && L(t), Ue(n), c && c.d();
            },
        };
    }
    function Fr(e) {
        let t, n;
        return (
            (t = new e[5]({ props: { cancel: e[11], width: "auto", $$slots: { default: [Rr] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    8 & n && (l.cancel = e[11]), 16384 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function qr(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r = e[13].label + "";
        return (
            (n = new Ys({ props: { name: e[13].icon } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = z("span")), (c = O(r)), U(o, "class", "svelte-1ju4xal"), U(t, "class", "menu-item svelte-1ju4xal");
                },
                m(e, r) {
                    T(e, t, r), Ke(n, t, null), I(t, l), I(t, o), I(o, c), (s = !0);
                },
                p(e, t) {
                    const l = {};
                    8192 & t && (l.name = e[13].icon), n.$set(l), (!s || 8192 & t) && r !== (r = e[13].label + "") && Y(c, r);
                },
                i(e) {
                    s || (Ae(n.$$.fragment, e), (s = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (s = !1);
                },
                d(e) {
                    e && L(t), Ue(n);
                },
            }
        );
    }
    function Rr(e) {
        let t, n;
        return (
            (t = new e[4]({ props: { click: e[8], data: [{ icon: "delete", label: e[6]("Delete"), id: 1 }], $$slots: { default: [qr, ({ obj: e }) => ({ 13: e }), ({ obj: e }) => (e ? 8192 : 0)] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    24576 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Pr(e) {
        let t,
            n,
            l = e[0].description + "";
        return {
            c() {
                (t = z("div")), (n = O(l)), U(t, "class", "field description svelte-1ju4xal");
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, t) {
                1 & t && l !== (l = e[0].description + "") && Y(n, l);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Kr(e) {
        let t, n, l;
        return (
            (n = new Lr({ props: { min: e[1].progress ?.config ?.min || 0, max: e[1].progress ?.config ?.max || 100, value: e[0].progress } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), U(t, "class", "field svelte-1ju4xal");
                },
                m(e, o) {
                    T(e, t, o), Ke(n, t, null), (l = !0);
                },
                p(e, t) {
                    const l = {};
                    2 & t && (l.min = e[1].progress ?.config ?.min || 0), 2 & t && (l.max = e[1].progress ?.config ?.max || 100), 1 & t && (l.value = e[0].progress), n.$set(l);
                },
                i(e) {
                    l || (Ae(n.$$.fragment, e), (l = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (l = !1);
                },
                d(e) {
                    e && L(t), Ue(n);
                },
            }
        );
    }
    function Ur(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i,
            a,
            u,
            d,
            p,
            f,
            $ = e[1].color ?.show && e[0].color && jr(e),
            m = e[1] ?.cover ?.show && e[2] && zr(e);
        o = new Ir({ props: { cardFields: e[0], cardShape: e[1] } });
        let h = e[1] ?.label ?.show && e[0].label && Nr(e),
            g = e[1] ?.menu ?.show && Or(e),
            v = e[1] ?.description ?.show && e[0].description && Pr(e),
            y = e[1] ?.progress ?.show && e[0].progress && Kr(e);
        return (
            (p = new wr({ props: { cardFields: e[0], cardShape: e[1] } })),
            {
                c() {
                    $ && $.c(),
                        (t = F()),
                        m && m.c(),
                        (n = F()),
                        (l = z("div")),
                        Pe(o.$$.fragment),
                        (c = F()),
                        (s = z("div")),
                        (r = z("div")),
                        h && h.c(),
                        (i = F()),
                        g && g.c(),
                        (a = F()),
                        v && v.c(),
                        (u = F()),
                        y && y.c(),
                        (d = F()),
                        Pe(p.$$.fragment),
                        U(r, "class", "field label svelte-1ju4xal"),
                        U(s, "class", "body svelte-1ju4xal"),
                        U(l, "class", "content svelte-1ju4xal");
                },
                m(e, w) {
                    $ && $.m(e, w),
                        T(e, t, w),
                        m && m.m(e, w),
                        T(e, n, w),
                        T(e, l, w),
                        Ke(o, l, null),
                        I(l, c),
                        I(l, s),
                        I(s, r),
                        h && h.m(r, null),
                        I(r, i),
                        g && g.m(r, null),
                        I(s, a),
                        v && v.m(s, null),
                        I(s, u),
                        y && y.m(s, null),
                        I(l, d),
                        Ke(p, l, null),
                        (f = !0);
                },
                p(e, [l]) {
                    e[1].color ?.show && e[0].color ? ($ ? $.p(e, l) : (($ = jr(e)), $.c(), $.m(t.parentNode, t))) : $ && ($.d(1), ($ = null)),
                        e[1] ?.cover ?.show && e[2] ? (m ? m.p(e, l) : ((m = zr(e)), m.c(), m.m(n.parentNode, n))) : m && (m.d(1), (m = null));
                    const c = {};
                    1 & l && (c.cardFields = e[0]),
                        2 & l && (c.cardShape = e[1]),
                        o.$set(c),
                        e[1] ?.label ?.show && e[0].label ? (h ? h.p(e, l) : ((h = Nr(e)), h.c(), h.m(r, i))) : h && (h.d(1), (h = null)),
                        e[1] ?.menu ?.show
                            ? g
                                ? (g.p(e, l), 2 & l && Ae(g, 1))
                                : ((g = Or(e)), g.c(), Ae(g, 1), g.m(r, null))
                            : g &&
                            (De(),
                                Ee(g, 1, 1, () => {
                                    g = null;
                                }),
                                Ie()),
                        e[1] ?.description ?.show && e[0].description ? (v ? v.p(e, l) : ((v = Pr(e)), v.c(), v.m(s, u))) : v && (v.d(1), (v = null)),
                        e[1] ?.progress ?.show && e[0].progress
                            ? y
                                ? (y.p(e, l), 3 & l && Ae(y, 1))
                                : ((y = Kr(e)), y.c(), Ae(y, 1), y.m(s, null))
                            : y &&
                            (De(),
                                Ee(y, 1, 1, () => {
                                    y = null;
                                }),
                                Ie());
                    const a = {};
                    1 & l && (a.cardFields = e[0]), 2 & l && (a.cardShape = e[1]), p.$set(a);
                },
                i(e) {
                    f || (Ae(o.$$.fragment, e), Ae(g), Ae(y), Ae(p.$$.fragment, e), (f = !0));
                },
                o(e) {
                    Ee(o.$$.fragment, e), Ee(g), Ee(y), Ee(p.$$.fragment, e), (f = !1);
                },
                d(e) {
                    $ && $.d(e), e && L(t), m && m.d(e), e && L(n), e && L(l), Ue(o), h && h.d(), g && g.d(), v && v.d(), y && y.d(), Ue(p);
                },
            }
        );
    }
    function Hr(e, t, n) {
        let l, o, c;
        var s;
        const { List: r, Dropdown: i } = Oc;
        let { cardFields: a } = t,
            { cardShape: u } = t;
        const d = ae("wx-i18n").getGroup("kanban"),
            p = re();
        return (
            (e.$$set = (e) => {
                "cardFields" in e && n(0, (a = e.cardFields)), "cardShape" in e && n(1, (u = e.cardShape));
            }),
            (e.$$.update = () => {
                513 & e.$$.dirty && n(10, (o = null === n(9, (s = null == a ? void 0 : a.attached)) || void 0 === s ? void 0 : s.find((e) => e.isCover))), 1024 & e.$$.dirty && n(2, (c = o ? o.coverURL || o.url : null));
            }),
            n(3, (l = !1)),
            [
                a,
                u,
                c,
                l,
                r,
                i,
                d,
                function (e) {
                    (e.cancelBubble = !0), n(3, (l = !l));
                },
                function (e) {
                    1 === e && p("action", { action: "delete-card", data: { id: a.id } });
                },
                s,
                o,
                () => n(3, (l = !1)),
            ]
        );
    }
    class Yr extends Ye {
        constructor(e) {
            super(), He(this, e, Hr, Ur, a, { cardFields: 0, cardShape: 1 });
        }
    }
    function Br(e) {
        let t, n, l, o;
        var c = e[4];
        function s(e) {
            return { props: { cardFields: e[0], dragging: e[1], selected: e[2], cardShape: e[3] } };
        }
        return (
            c && ((n = new c(s(e))), n.$on("action", e[6])),
            {
                c() {
                    (t = z("div")), n && Pe(n.$$.fragment), U(t, "class", "card svelte-9295rp"), U(t, "data-drag-item", (l = e[0].id)), V(t, "hidden", e[1]), V(t, "selected", e[2]), V(t, "dimmed", e[5] ?.dimmed);
                },
                m(e, l) {
                    T(e, t, l), n && Ke(n, t, null), (o = !0);
                },
                p(e, [r]) {
                    const i = {};
                    if ((1 & r && (i.cardFields = e[0]), 2 & r && (i.dragging = e[1]), 4 & r && (i.selected = e[2]), 8 & r && (i.cardShape = e[3]), c !== (c = e[4]))) {
                        if (n) {
                            De();
                            const e = n;
                            Ee(e.$$.fragment, 1, 0, () => {
                                Ue(e, 1);
                            }),
                                Ie();
                        }
                        c ? ((n = new c(s(e))), n.$on("action", e[6]), Pe(n.$$.fragment), Ae(n.$$.fragment, 1), Ke(n, t, null)) : (n = null);
                    } else c && n.$set(i);
                    (!o || (1 & r && l !== (l = e[0].id))) && U(t, "data-drag-item", l), 2 & r && V(t, "hidden", e[1]), 4 & r && V(t, "selected", e[2]), 32 & r && V(t, "dimmed", e[5] ?.dimmed);
                },
                i(e) {
                    o || (n && Ae(n.$$.fragment, e), (o = !0));
                },
                o(e) {
                    n && Ee(n.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(t), n && Ue(n);
                },
            }
        );
    }
    function Gr(e, t, n) {
        let { cardFields: l } = t,
            { dragging: o = !1 } = t,
            { selected: c = !1 } = t,
            { cardShape: s } = t,
            { cardTemplate: r } = t,
            { meta: i = null } = t;
        return (
            (e.$$set = (e) => {
                "cardFields" in e && n(0, (l = e.cardFields)),
                    "dragging" in e && n(1, (o = e.dragging)),
                    "selected" in e && n(2, (c = e.selected)),
                    "cardShape" in e && n(3, (s = e.cardShape)),
                    "cardTemplate" in e && n(4, (r = e.cardTemplate)),
                    "meta" in e && n(5, (i = e.meta));
            }),
            [
                l,
                o,
                c,
                s,
                r,
                i,
                function (t) {
                    ue.call(this, e, t);
                },
            ]
        );
    }
    class Jr extends Ye {
        constructor(e) {
            super(), He(this, e, Gr, Br, a, { cardFields: 0, dragging: 1, selected: 2, cardShape: 3, cardTemplate: 4, meta: 5 });
        }
    }
    function Vr(e, t, n) {
        const l = e.slice();
        return (l[24] = t[n]), l;
    }
    function Xr(e) {
        let t,
            l,
            o,
            c = e[0].label + "";
        return {
            c() {
                (t = z("div")), (l = z("div")), (o = O(c)), U(l, "class", "label-text svelte-173kp57"), U(t, "class", "collapsed-label svelte-173kp57");
            },
            m(e, n) {
                T(e, t, n), I(t, l), I(l, o);
            },
            p(e, t) {
                1 & t && c !== (c = e[0].label + "") && Y(o, c);
            },
            i: n,
            o: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function Qr(e) {
        let t,
            n,
            l,
            o,
            c,
            s = !e[2] && Jc(e[3], e[10]),
            r = e[1] && Wr(e),
            i = s && ti(e),
            a = e[8] && !e[6][e[10]].noFreeSpace && ni(e),
            u = e[6][e[10]].rowId && e[6][e[10]].limit && li(e);
        return {
            c() {
                r && r.c(), (t = F()), i && i.c(), (n = F()), (l = z("div")), a && a.c(), (o = F()), u && u.c(), U(l, "class", "controls-wrapper svelte-173kp57");
            },
            m(e, s) {
                r && r.m(e, s), T(e, t, s), i && i.m(e, s), T(e, n, s), T(e, l, s), a && a.m(l, null), I(l, o), u && u.m(l, null), (c = !0);
            },
            p(e, c) {
                e[1]
                    ? r
                        ? (r.p(e, c), 2 & c && Ae(r, 1))
                        : ((r = Wr(e)), r.c(), Ae(r, 1), r.m(t.parentNode, t))
                    : r &&
                    (De(),
                        Ee(r, 1, 1, () => {
                            r = null;
                        }),
                        Ie()),
                    1036 & c && (s = !e[2] && Jc(e[3], e[10])),
                    s ? (i ? i.p(e, c) : ((i = ti(e)), i.c(), i.m(n.parentNode, n))) : i && (i.d(1), (i = null)),
                    e[8] && !e[6][e[10]].noFreeSpace
                        ? a
                            ? (a.p(e, c), 1344 & c && Ae(a, 1))
                            : ((a = ni(e)), a.c(), Ae(a, 1), a.m(l, o))
                        : a &&
                        (De(),
                            Ee(a, 1, 1, () => {
                                a = null;
                            }),
                            Ie()),
                    e[6][e[10]].rowId && e[6][e[10]].limit ? (u ? u.p(e, c) : ((u = li(e)), u.c(), u.m(l, null))) : u && (u.d(1), (u = null));
            },
            i(e) {
                c || (Ae(r), Ae(a), (c = !0));
            },
            o(e) {
                Ee(r), Ee(a), (c = !1);
            },
            d(e) {
                r && r.d(e), e && L(t), i && i.d(e), e && L(n), e && L(l), a && a.d(), u && u.d();
            },
        };
    }
    function Wr(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[1];
        const s = (e) => e[24].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = Vr(e, c, t),
                r = s(n);
            o.set(r, (l[t] = ei(r, n)));
        }
        return {
            c() {
                for (let e = 0; e < l.length; e += 1) l[e].c();
                t = q();
            },
            m(e, o) {
                for (let t = 0; t < l.length; t += 1) l[t].m(e, o);
                T(e, t, o), (n = !0);
            },
            p(e, n) {
                10934 & n && ((c = e[1]), De(), (l = Oe(l, n, s, 1, e, c, o, t.parentNode, Ne, ei, t, Vr)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                for (let t = 0; t < l.length; t += 1) l[t].d(e);
                e && L(t);
            },
        };
    }
    function Zr(e) {
        let t;
        return {
            c() {
                (t = z("div")), U(t, "class", "drop-area svelte-173kp57"), G(t, "min-height", e[11] + "px");
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, n) {
                2048 & n && G(t, "min-height", e[11] + "px");
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function ei(e, t) {
        let n,
            l,
            o,
            c,
            s = Jc(t[24].id, t[2]),
            r = s && Zr(t);
        return (
            (o = new Jr({ props: { cardTemplate: t[7] || Yr, cardFields: t[24], dragging: t[13](t[24].id), selected: Vc(t[4], t[24].id), meta: t[9] && t[9][t[24].id], cardShape: t[5] } })),
            o.$on("action", t[22]),
            {
                key: e,
                first: null,
                c() {
                    (n = q()), r && r.c(), (l = F()), Pe(o.$$.fragment), (this.first = n);
                },
                m(e, t) {
                    T(e, n, t), r && r.m(e, t), T(e, l, t), Ke(o, e, t), (c = !0);
                },
                p(e, n) {
                    (t = e), 6 & n && (s = Jc(t[24].id, t[2])), s ? (r ? r.p(t, n) : ((r = Zr(t)), r.c(), r.m(l.parentNode, l))) : r && (r.d(1), (r = null));
                    const c = {};
                    128 & n && (c.cardTemplate = t[7] || Yr),
                        2 & n && (c.cardFields = t[24]),
                        8194 & n && (c.dragging = t[13](t[24].id)),
                        18 & n && (c.selected = Vc(t[4], t[24].id)),
                        514 & n && (c.meta = t[9] && t[9][t[24].id]),
                        32 & n && (c.cardShape = t[5]),
                        o.$set(c);
                },
                i(e) {
                    c || (Ae(o.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(o.$$.fragment, e), (c = !1);
                },
                d(e) {
                    e && L(n), r && r.d(e), e && L(l), Ue(o, e);
                },
            }
        );
    }
    function ti(e) {
        let t;
        return {
            c() {
                (t = z("div")), U(t, "class", "drop-area svelte-173kp57"), G(t, "min-height", e[11] + "px");
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, n) {
                2048 & n && G(t, "min-height", e[11] + "px");
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function ni(e) {
        let t, l, o, c, s, r, i;
        return (
            (l = new Ys({ props: { name: "plus" } })),
            {
                c() {
                    (t = z("div")), Pe(l.$$.fragment), (o = F()), (c = z("span")), (c.textContent = `${e[14]("Add new card")}...`), U(c, "class", "add-card-tip svelte-173kp57"), U(t, "class", "add-card-btn svelte-173kp57");
                },
                m(n, a) {
                    T(n, t, a), Ke(l, t, null), I(t, o), I(t, c), (s = !0), r || ((i = R(t, "click", e[15])), (r = !0));
                },
                p: n,
                i(e) {
                    s || (Ae(l.$$.fragment, e), (s = !0));
                },
                o(e) {
                    Ee(l.$$.fragment, e), (s = !1);
                },
                d(e) {
                    e && L(t), Ue(l), (r = !1), i();
                },
            }
        );
    }
    function li(e) {
        let t,
            n,
            l,
            o,
            c = e[6][e[10]].cardsCount + "",
            s = e[6][e[10]].limit + "";
        return {
            c() {
                (t = z("div")), (n = O(c)), (l = O("/")), (o = O(s)), U(t, "class", "swimlane-limit svelte-173kp57");
            },
            m(e, c) {
                T(e, t, c), I(t, n), I(t, l), I(t, o);
            },
            p(e, t) {
                1088 & t && c !== (c = e[6][e[10]].cardsCount + "") && Y(n, c), 1088 & t && s !== (s = e[6][e[10]].limit + "") && Y(o, s);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function oi(e) {
        let t, n, l, o;
        const c = [Qr, Xr],
            s = [];
        function r(e, t) {
            return e[0].collapsed ? 1 : 0;
        }
        return (
            (n = r(e)),
            (l = s[n] = c[n](e)),
            {
                c() {
                    (t = z("div")), l.c(), U(t, "class", "column svelte-173kp57"), U(t, "data-drop-area", e[10]), G(t, "min-height", e[12]), V(t, "collapsed", e[0].collapsed), V(t, "over-limit", e[6][e[10]].isOverLimit);
                },
                m(e, l) {
                    T(e, t, l), s[n].m(t, null), (o = !0);
                },
                p(e, [i]) {
                    let a = n;
                    (n = r(e)),
                        n === a
                            ? s[n].p(e, i)
                            : (De(),
                                Ee(s[a], 1, 1, () => {
                                    s[a] = null;
                                }),
                                Ie(),
                                (l = s[n]),
                                l ? l.p(e, i) : ((l = s[n] = c[n](e)), l.c()),
                                Ae(l, 1),
                                l.m(t, null)),
                        (!o || 1024 & i) && U(t, "data-drop-area", e[10]),
                        (!o || 4096 & i) && G(t, "min-height", e[12]),
                        1 & i && V(t, "collapsed", e[0].collapsed),
                        1088 & i && V(t, "over-limit", e[6][e[10]].isOverLimit);
                },
                i(e) {
                    o || (Ae(l), (o = !0));
                },
                o(e) {
                    Ee(l), (o = !1);
                },
                d(e) {
                    e && L(t), s[n].d();
                },
            }
        );
    }
    function ci(e, t, n) {
        let l, o, c, s, r;
        var i;
        let { column: a } = t,
            { row: u } = t,
            { cards: d } = t,
            { overCardId: p } = t,
            { movedCardId: f } = t,
            { movedCardCoords: $ } = t,
            { overColId: m } = t,
            { selected: h } = t,
            { dropAreasCoords: g } = t,
            { cardShape: v } = t,
            { areasMeta: y } = t,
            { cardTemplate: w = null } = t,
            { add: b = !0 } = t,
            { cardsMeta: x = null } = t;
        const k = ae("wx-i18n").getGroup("kanban"),
            S = re();
        return (
            (e.$$set = (e) => {
                "column" in e && n(0, (a = e.column)),
                    "row" in e && n(16, (u = e.row)),
                    "cards" in e && n(1, (d = e.cards)),
                    "overCardId" in e && n(2, (p = e.overCardId)),
                    "movedCardId" in e && n(17, (f = e.movedCardId)),
                    "movedCardCoords" in e && n(18, ($ = e.movedCardCoords)),
                    "overColId" in e && n(3, (m = e.overColId)),
                    "selected" in e && n(4, (h = e.selected)),
                    "dropAreasCoords" in e && n(19, (g = e.dropAreasCoords)),
                    "cardShape" in e && n(5, (v = e.cardShape)),
                    "areasMeta" in e && n(6, (y = e.areasMeta)),
                    "cardTemplate" in e && n(7, (w = e.cardTemplate)),
                    "add" in e && n(8, (b = e.add)),
                    "cardsMeta" in e && n(9, (x = e.cardsMeta));
            }),
            (e.$$.update = () => {
                131088 & e.$$.dirty && n(13, (l = (e) => !!f && (Jc(f, e) || ((null == h ? void 0 : h.length) > 1 && Vc(h, e))))),
                    65537 & e.$$.dirty && n(10, (o = Xc(a.id, u.id))),
                    1573888 & e.$$.dirty && n(21, (c = g && (null === n(20, (i = null == g ? void 0 : g.find((e) => e.id === o))) || void 0 === i ? void 0 : i.height))),
                    2097152 & e.$$.dirty && n(12, (s = c ? `${c}px` : "auto")),
                    262144 & e.$$.dirty && n(11, (r = (null == $ ? void 0 : $.height) - 2));
            }),
            [
                a,
                d,
                p,
                m,
                h,
                v,
                y,
                w,
                b,
                x,
                o,
                r,
                s,
                l,
                k,
                function (e) {
                    e.stopPropagation(), S("action", { action: "add-card", data: { columnId: a.id, rowId: u.id } });
                },
                u,
                f,
                $,
                g,
                i,
                c,
                function (t) {
                    ue.call(this, e, t);
                },
            ]
        );
    }
    class si extends Ye {
        constructor(e) {
            super(),
                He(this, e, ci, oi, a, {
                    column: 0,
                    row: 16,
                    cards: 1,
                    overCardId: 2,
                    movedCardId: 17,
                    movedCardCoords: 18,
                    overColId: 3,
                    selected: 4,
                    dropAreasCoords: 19,
                    cardShape: 5,
                    areasMeta: 6,
                    cardTemplate: 7,
                    add: 8,
                    cardsMeta: 9,
                });
        }
    }
    function ri(e, t) {
        for (const n in t) {
            const l = e[n],
                o = t[n];
            if (l !== o) {
                if (!Array.isArray(l) || !Array.isArray(o) || l.length !== o.length) return !1;
                for (let e = l.length - 1; e >= 0; e--) if (l[e] !== o[e]) return !1;
            }
        }
        return !0;
    }
    function ii(e, t, n) {
        let l = !1,
            o = null;
        const c = zn(e),
            { set: s } = c;
        let r = { ...e };
        return (
            (c.set = function (e) {
                ri(r, e) || ((r = { ...e }), s(e));
            }),
            (c.update = function (t) {
                const n = t(e);
                ri(r, n) || ((r = { ...n }), s(n));
            }),
            (c.reset = function (e) {
                (l = !1), (r = {}), c.set(e);
            }),
            c.subscribe((e) => {
                l ? e && (n && n.debounce ? (clearTimeout(o), (o = setTimeout(() => t(e), n.debounce))) : t(e)) : (l = !0);
            }),
            c
        );
    }
    function ai(e, t, n) {
        const l = e.slice();
        return (l[13] = t[n]), l;
    }
    function ui(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i,
            a = [],
            u = new Map(),
            d = e[1];
        const p = (e) => e[13].id;
        for (let t = 0; t < d.length; t += 1) {
            let n = ai(e, d, t),
                l = p(n);
            u.set(l, (a[t] = xi(l, n)));
        }
        return {
            c() {
                (t = z("div")), (n = z("div")), (l = z("i")), (o = F()), (c = z("div"));
                for (let e = 0; e < a.length; e += 1) a[e].c();
                U(l, "class", "far fa-times svelte-hvce4g"), U(n, "class", "header svelte-hvce4g"), U(c, "class", "list svelte-hvce4g"), U(t, "class", "layout svelte-hvce4g");
            },
            m(u, d) {
                T(u, t, d), I(t, n), I(n, l), I(t, o), I(t, c);
                for (let e = 0; e < a.length; e += 1) a[e].m(c, null);
                (s = !0), r || ((i = R(l, "click", e[3])), (r = !0));
            },
            p(e, t) {
                498 & t && ((d = e[1]), De(), (a = Oe(a, t, p, 1, e, d, u, c, Ne, xi, null, ai)), Ie());
            },
            i(e) {
                if (!s) {
                    for (let e = 0; e < d.length; e += 1) Ae(a[e]);
                    s = !0;
                }
            },
            o(e) {
                for (let e = 0; e < a.length; e += 1) Ee(a[e]);
                s = !1;
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < a.length; e += 1) a[e].d();
                (r = !1), i();
            },
        };
    }
    function di(e) {
        let t, l;
        return (
            (t = new Ys({ props: { name: "paperclip", size: 20 } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, n) {
                    Ke(t, e, n), (l = !0);
                },
                p: n,
                i(e) {
                    l || (Ae(t.$$.fragment, e), (l = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (l = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function pi(e) {
        let t;
        return {
            c() {
                (t = z("div")), U(t, "class", "thumb svelte-hvce4g"), G(t, "background-image", "url('" + (e[13].previewURL || e[13].url) + "')");
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, n) {
                2 & n && G(t, "background-image", "url('" + (e[13].previewURL || e[13].url) + "')");
            },
            i: n,
            o: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function fi(e) {
        let t,
            n,
            l = e[5](e[13].file.size) + "";
        return {
            c() {
                (t = z("div")), (n = O(l)), U(t, "class", "size");
            },
            m(e, l) {
                T(e, t, l), I(t, n);
            },
            p(e, t) {
                2 & t && l !== (l = e[5](e[13].file.size) + "") && Y(n, l);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function $i(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i,
            a,
            u = e[6](e[13]);
        function d() {
            return e[10](e[13]);
        }
        (l = new Ys({ props: { name: "external", clickable: !0 } })), (r = new Ys({ props: { name: "delete", click: d } }));
        let p = u && gi(e);
        return {
            c() {
                (t = z("div")),
                    (n = z("a")),
                    Pe(l.$$.fragment),
                    (s = F()),
                    Pe(r.$$.fragment),
                    (i = F()),
                    p && p.c(),
                    U(n, "class", "upload-link svelte-hvce4g"),
                    U(n, "href", (o = e[13].url)),
                    U(n, "download", (c = e[13].name)),
                    U(n, "target", "_blank"),
                    U(n, "rel", "noreferrer nofollow noopener"),
                    U(t, "class", "hidden svelte-hvce4g");
            },
            m(e, o) {
                T(e, t, o), I(t, n), Ke(l, n, null), I(t, s), Ke(r, t, null), I(t, i), p && p.m(t, null), (a = !0);
            },
            p(l, s) {
                (e = l), (!a || (2 & s && o !== (o = e[13].url))) && U(n, "href", o), (!a || (2 & s && c !== (c = e[13].name))) && U(n, "download", c);
                const i = {};
                2 & s && (i.click = d),
                    r.$set(i),
                    2 & s && (u = e[6](e[13])),
                    u
                        ? p
                            ? (p.p(e, s), 2 & s && Ae(p, 1))
                            : ((p = gi(e)), p.c(), Ae(p, 1), p.m(t, null))
                        : p &&
                        (De(),
                            Ee(p, 1, 1, () => {
                                p = null;
                            }),
                            Ie());
            },
            i(e) {
                a || (Ae(l.$$.fragment, e), Ae(r.$$.fragment, e), Ae(p), (a = !0));
            },
            o(e) {
                Ee(l.$$.fragment, e), Ee(r.$$.fragment, e), Ee(p), (a = !1);
            },
            d(e) {
                e && L(t), Ue(l), Ue(r), p && p.d();
            },
        };
    }
    function mi(e) {
        let t, n, l, o;
        function c() {
            return e[9](e[13]);
        }
        return (
            (t = new Ys({ props: { name: "alert" } })),
            (l = new Ys({ props: { name: "delete", click: c } })),
            {
                c() {
                    Pe(t.$$.fragment), (n = F()), Pe(l.$$.fragment);
                },
                m(e, c) {
                    Ke(t, e, c), T(e, n, c), Ke(l, e, c), (o = !0);
                },
                p(t, n) {
                    e = t;
                    const o = {};
                    2 & n && (o.click = c), l.$set(o);
                },
                i(e) {
                    o || (Ae(t.$$.fragment, e), Ae(l.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), Ee(l.$$.fragment, e), (o = !1);
                },
                d(e) {
                    Ue(t, e), e && L(n), Ue(l, e);
                },
            }
        );
    }
    function hi(e) {
        let t, l;
        return (
            (t = new Ys({ props: { name: "loading", spin: !0 } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, n) {
                    Ke(t, e, n), (l = !0);
                },
                p: n,
                i(e) {
                    l || (Ae(t.$$.fragment, e), (l = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (l = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function gi(e) {
        let t, n, l, o;
        const c = [yi, vi],
            s = [];
        function r(e, t) {
            return e[13].isCover ? 1 : 0;
        }
        return (
            (t = r(e)),
            (n = s[t] = c[t](e)),
            {
                c() {
                    n.c(), (l = q());
                },
                m(e, n) {
                    s[t].m(e, n), T(e, l, n), (o = !0);
                },
                p(e, o) {
                    let i = t;
                    (t = r(e)),
                        t === i
                            ? s[t].p(e, o)
                            : (De(),
                                Ee(s[i], 1, 1, () => {
                                    s[i] = null;
                                }),
                                Ie(),
                                (n = s[t]),
                                n ? n.p(e, o) : ((n = s[t] = c[t](e)), n.c()),
                                Ae(n, 1),
                                n.m(l.parentNode, l));
                },
                i(e) {
                    o || (Ae(n), (o = !0));
                },
                o(e) {
                    Ee(n), (o = !1);
                },
                d(e) {
                    s[t].d(e), e && L(l);
                },
            }
        );
    }
    function vi(e) {
        let t, n;
        return (
            (t = new e[2]({ props: { click: e[8], $$slots: { default: [wi] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    65536 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function yi(e) {
        let t, n;
        function l() {
            return e[11](e[13]);
        }
        return (
            (t = new e[2]({ props: { click: l, $$slots: { default: [bi] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(n, o) {
                    e = n;
                    const c = {};
                    2 & o && (c.click = l), 65536 & o && (c.$$scope = { dirty: o, ctx: e }), t.$set(c);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function wi(e) {
        let t;
        return {
            c() {
                t = O("Remove cover");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function bi(e) {
        let t;
        return {
            c() {
                t = O("Make cover");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function xi(e, t) {
        let n,
            l,
            o,
            c,
            s,
            r,
            i,
            a,
            u,
            d,
            p,
            f,
            $,
            m,
            h,
            g = t[13].name + "";
        const v = [pi, di],
            y = [];
        function w(e, t) {
            return 2 & t && (o = !!e[6](e[13])), o ? 0 : 1;
        }
        (c = w(t, -1)), (s = y[c] = v[c](t));
        let b = t[13].file && fi(t);
        const x = [hi, mi, $i],
            k = [];
        function S(e, t) {
            return "client" === e[13].status ? 0 : "error" === e[13].status ? 1 : e[13].status && "server" !== e[13].status ? -1 : 2;
        }
        return (
            ~(f = S(t)) && ($ = k[f] = x[f](t)),
            {
                key: e,
                first: null,
                c() {
                    (n = z("div")),
                        (l = z("div")),
                        s.c(),
                        (r = F()),
                        (i = z("div")),
                        (a = O(g)),
                        (u = F()),
                        b && b.c(),
                        (d = F()),
                        (p = z("div")),
                        $ && $.c(),
                        (m = F()),
                        U(l, "class", "file-icon svelte-hvce4g"),
                        U(i, "class", "name svelte-hvce4g"),
                        U(p, "class", "controls svelte-hvce4g"),
                        U(n, "class", "row svelte-hvce4g"),
                        (this.first = n);
                },
                m(e, t) {
                    T(e, n, t), I(n, l), y[c].m(l, null), I(n, r), I(n, i), I(i, a), I(n, u), b && b.m(n, null), I(n, d), I(n, p), ~f && k[f].m(p, null), I(n, m), (h = !0);
                },
                p(e, o) {
                    let r = c;
                    (c = w((t = e), o)),
                        c === r
                            ? y[c].p(t, o)
                            : (De(),
                                Ee(y[r], 1, 1, () => {
                                    y[r] = null;
                                }),
                                Ie(),
                                (s = y[c]),
                                s ? s.p(t, o) : ((s = y[c] = v[c](t)), s.c()),
                                Ae(s, 1),
                                s.m(l, null)),
                        (!h || 2 & o) && g !== (g = t[13].name + "") && Y(a, g),
                        t[13].file ? (b ? b.p(t, o) : ((b = fi(t)), b.c(), b.m(n, d))) : b && (b.d(1), (b = null));
                    let i = f;
                    (f = S(t)),
                        f === i
                            ? ~f && k[f].p(t, o)
                            : ($ &&
                                (De(),
                                    Ee(k[i], 1, 1, () => {
                                        k[i] = null;
                                    }),
                                    Ie()),
                                ~f ? (($ = k[f]), $ ? $.p(t, o) : (($ = k[f] = x[f](t)), $.c()), Ae($, 1), $.m(p, null)) : ($ = null));
                },
                i(e) {
                    h || (Ae(s), Ae($), (h = !0));
                },
                o(e) {
                    Ee(s), Ee($), (h = !1);
                },
                d(e) {
                    e && L(n), y[c].d(), b && b.d(), ~f && k[f].d();
                },
            }
        );
    }
    function ki(e) {
        let t,
            n,
            l = e[1].length && ui(e);
        return {
            c() {
                l && l.c(), (t = q());
            },
            m(e, o) {
                l && l.m(e, o), T(e, t, o), (n = !0);
            },
            p(e, [n]) {
                e[1].length
                    ? l
                        ? (l.p(e, n), 2 & n && Ae(l, 1))
                        : ((l = ui(e)), l.c(), Ae(l, 1), l.m(t.parentNode, t))
                    : l &&
                    (De(),
                        Ee(l, 1, 1, () => {
                            l = null;
                        }),
                        Ie());
            },
            i(e) {
                n || (Ae(l), (n = !0));
            },
            o(e) {
                Ee(l), (n = !1);
            },
            d(e) {
                l && l.d(e), e && L(t);
            },
        };
    }
    function Si(e, t, l) {
        let o,
            c = n,
            s = () => (c(), (c = p(i, (e) => l(1, (o = e)))), i);
        e.$$.on_destroy.push(() => c());
        const { Button: r } = Oc;
        let { data: i } = t;
        s();
        const a = ["b", "Kb", "Mb", "Gb", "Tb", "Pb", "Eb"];
        function u(e) {
            i.update((t) => t.filter((t) => t.id !== e));
        }
        function d(e) {
            i.update((t) => t.map((t) => (t.id === e ? Object.assign(Object.assign({}, t), { isCover: !0 }) : (delete t.isCover, t))));
        }
        return (
            (e.$$set = (e) => {
                "data" in e && s(l(0, (i = e.data)));
            }),
            [
                i,
                o,
                r,
                function () {
                    i.set([]);
                },
                u,
                function (e) {
                    let t = 0;
                    for (; e > 1024;) t++ , (e /= 1024);
                    return Math.round(100 * e) / 100 + " " + a[t];
                },
                function (e) {
                    var t, n;
                    const l = null === (t = null == e ? void 0 : e.url) || void 0 === t ? void 0 : t.split(".").pop();
                    return js(null === (n = null == e ? void 0 : e.previewURL) || void 0 === n ? void 0 : n.split(".").pop()) || js(l);
                },
                d,
                function () {
                    i.update((e) =>
                        e.map((e) => {
                            const t = Object.assign({}, e);
                            return delete t.isCover, t;
                        })
                    );
                },
                (e) => u(e.id),
                (e) => u(e.id),
                (e) => d(e.id),
            ]
        );
    }
    class Mi extends Ye {
        constructor(e) {
            super(), He(this, e, Si, ki, a, { data: 0 });
        }
    }
    function _i(e) {
        let t, n, l, c, s;
        t = new Mi({ props: { data: e[3] } });
        const r = [{ data: e[3] }, { uploadURL: e[0].uploadURL }, e[0].config];
        function i(t) {
            e[7](t);
        }
        let a = {};
        for (let e = 0; e < r.length; e += 1) a = o(a, r[e]);
        return (
            void 0 !== e[2][e[0].key] && (a.value = e[2][e[0].key]),
            (l = new e[4]({ props: a })),
            pe.push(() => Re(l, "value", i)),
            {
                c() {
                    Pe(t.$$.fragment), (n = F()), Pe(l.$$.fragment);
                },
                m(e, o) {
                    Ke(t, e, o), T(e, n, o), Ke(l, e, o), (s = !0);
                },
                p(e, [n]) {
                    const o = {};
                    8 & n && (o.data = e[3]), t.$set(o);
                    const s = 9 & n ? Fe(r, [8 & n && { data: e[3] }, 1 & n && { uploadURL: e[0].uploadURL }, 1 & n && qe(e[0].config)]) : {};
                    !c && 5 & n && ((c = !0), (s.value = e[2][e[0].key]), ye(() => (c = !1))), l.$set(s);
                },
                i(e) {
                    s || (Ae(t.$$.fragment, e), Ae(l.$$.fragment, e), (s = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), Ee(l.$$.fragment, e), (s = !1);
                },
                d(e) {
                    Ue(t, e), e && L(n), Ue(l, e);
                },
            }
        );
    }
    function Ci(e, t, l) {
        let o,
            c,
            s,
            r = n,
            i = n,
            a = () => (i(), (i = p(f, (e) => l(2, (s = e)))), f);
        e.$$.on_destroy.push(() => r()), e.$$.on_destroy.push(() => i());
        const { Uploader: u } = Oc;
        let { field: d } = t,
            { values: f } = t;
        a();
        let $ = !1;
        return (
            (e.$$set = (e) => {
                "field" in e && l(0, (d = e.field)), "values" in e && a(l(1, (f = e.values)));
            }),
            (e.$$.update = () => {
                97 & e.$$.dirty && ($ && x(f, (s[d.key] = c), s), l(5, ($ = !0))), 5 & e.$$.dirty && (l(3, (o = zn(s[d.key] || []))), r(), (r = p(o, (e) => l(6, (c = e)))));
            }),
            [
                d,
                f,
                s,
                o,
                u,
                $,
                c,
                function (t) {
                    e.$$.not_equal(s[d.key], t) && ((s[d.key] = t), f.set(s));
                },
            ]
        );
    }
    class Di extends Ye {
        constructor(e) {
            super(), He(this, e, Ci, _i, a, { field: 0, values: 1 });
        }
    }
    function Ii(e) {
        let t, n, l, o, c, s;
        return (
            (n = new e[8]({ props: { cancel: e[13], width: "unset", $$slots: { default: [Li] }, $$scope: { ctx: e } } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment);
                },
                m(r, i) {
                    T(r, t, i), Ke(n, t, null), (o = !0), c || ((s = k((l = zs.call(null, t, { container: ".wx-portal", at: e[4], position: "bottom", align: "end" })))), (c = !0));
                },
                p(e, t) {
                    const o = {};
                    8388705 & t && (o.$$scope = { dirty: t, ctx: e }), n.$set(o), l && i(l.update) && 16 & t && l.update.call(null, { container: ".wx-portal", at: e[4], position: "bottom", align: "end" });
                },
                i(e) {
                    o || (Ae(n.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(t), Ue(n), (c = !1), s();
                },
            }
        );
    }
    function Ai(e) {
        let t;
        return {
            c() {
                t = O("Done");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Ei(e) {
        let t;
        return {
            c() {
                t = O("Clear");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Ti(e) {
        let t;
        return {
            c() {
                t = O("Today");
            },
            m(e, n) {
                T(e, t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function Li(e) {
        let t, n, l, o, c, s, r, i, a, u, d, p, f, $, m, h, g, v, y, w;
        function b(t) {
            e[17](t);
        }
        function x(t) {
            e[18](t);
        }
        let k = { part: "left" };
        function S(t) {
            e[19](t);
        }
        function M(t) {
            e[20](t);
        }
        void 0 !== e[0] && (k.value = e[0]), void 0 !== e[5] && (k.current = e[5]), (o = new e[10]({ props: k })), pe.push(() => Re(o, "value", b)), pe.push(() => Re(o, "current", x));
        let _ = { part: "right" };
        return (
            void 0 !== e[0] && (_.value = e[0]),
            void 0 !== e[6] && (_.current = e[6]),
            (a = new e[10]({ props: _ })),
            pe.push(() => Re(a, "value", S)),
            pe.push(() => Re(a, "current", M)),
            (m = new e[9]({ props: { type: "primary", click: e[13], $$slots: { default: [Ai] }, $$scope: { ctx: e } } })),
            (g = new e[9]({ props: { type: "link", click: e[15], $$slots: { default: [Ei] }, $$scope: { ctx: e } } })),
            (y = new e[9]({ props: { type: "link", click: e[21], $$slots: { default: [Ti] }, $$scope: { ctx: e } } })),
            {
                c() {
                    (t = z("div")),
                        (n = z("div")),
                        (l = z("div")),
                        Pe(o.$$.fragment),
                        (r = F()),
                        (i = z("div")),
                        Pe(a.$$.fragment),
                        (p = F()),
                        (f = z("div")),
                        ($ = z("div")),
                        Pe(m.$$.fragment),
                        (h = F()),
                        Pe(g.$$.fragment),
                        (v = F()),
                        Pe(y.$$.fragment),
                        U(l, "class", "half svelte-cik0z8"),
                        U(i, "class", "half svelte-cik0z8"),
                        U(n, "class", "calendars svelte-cik0z8"),
                        U($, "class", "done svelte-cik0z8"),
                        U(f, "class", "buttons svelte-cik0z8"),
                        U(t, "class", "calendar svelte-cik0z8");
                },
                m(e, c) {
                    T(e, t, c), I(t, n), I(n, l), Ke(o, l, null), I(n, r), I(n, i), Ke(a, i, null), I(t, p), I(t, f), I(f, $), Ke(m, $, null), I(f, h), Ke(g, f, null), I(f, v), Ke(y, f, null), (w = !0);
                },
                p(e, t) {
                    const n = {};
                    !c && 1 & t && ((c = !0), (n.value = e[0]), ye(() => (c = !1))), !s && 32 & t && ((s = !0), (n.current = e[5]), ye(() => (s = !1))), o.$set(n);
                    const l = {};
                    !u && 1 & t && ((u = !0), (l.value = e[0]), ye(() => (u = !1))), !d && 64 & t && ((d = !0), (l.current = e[6]), ye(() => (d = !1))), a.$set(l);
                    const r = {};
                    8388608 & t && (r.$$scope = { dirty: t, ctx: e }), m.$set(r);
                    const i = {};
                    8388608 & t && (i.$$scope = { dirty: t, ctx: e }), g.$set(i);
                    const p = {};
                    8388608 & t && (p.$$scope = { dirty: t, ctx: e }), y.$set(p);
                },
                i(e) {
                    w || (Ae(o.$$.fragment, e), Ae(a.$$.fragment, e), Ae(m.$$.fragment, e), Ae(g.$$.fragment, e), Ae(y.$$.fragment, e), (w = !0));
                },
                o(e) {
                    Ee(o.$$.fragment, e), Ee(a.$$.fragment, e), Ee(m.$$.fragment, e), Ee(g.$$.fragment, e), Ee(y.$$.fragment, e), (w = !1);
                },
                d(e) {
                    e && L(t), Ue(o), Ue(a), Ue(m), Ue(g), Ue(y);
                },
            }
        );
    }
    function ji(e) {
        let t, n, l, o, c, s, i, a;
        n = new e[7]({ props: { value: e[3], id: e[1], readonly: !0, inputStyle: "cursor: pointer; text-overflow: ellipsis; padding-right: 18px;" } });
        let u = e[2] && Ii(e);
        return {
            c() {
                (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = z("i")), (c = F()), u && u.c(), U(o, "class", "icon wxi-calendar svelte-cik0z8"), U(t, "class", "layout svelte-cik0z8");
            },
            m(r, d) {
                T(r, t, d), Ke(n, t, null), I(t, l), I(t, o), I(t, c), u && u.m(t, null), e[22](t), (s = !0), i || ((a = [R(window, "scroll", e[13]), R(t, "click", e[16])]), (i = !0));
            },
            p(e, [l]) {
                const o = {};
                8 & l && (o.value = e[3]),
                    2 & l && (o.id = e[1]),
                    n.$set(o),
                    e[2]
                        ? u
                            ? (u.p(e, l), 4 & l && Ae(u, 1))
                            : ((u = Ii(e)), u.c(), Ae(u, 1), u.m(t, null))
                        : u &&
                        (De(),
                            Ee(u, 1, 1, () => {
                                u = null;
                            }),
                            Ie());
            },
            i(e) {
                s || (Ae(n.$$.fragment, e), Ae(u), (s = !0));
            },
            o(e) {
                Ee(n.$$.fragment, e), Ee(u), (s = !1);
            },
            d(l) {
                l && L(t), Ue(n), u && u.d(), e[22](null), (i = !1), r(a);
            },
        };
    }
    function zi(e, t, n) {
        let l, o;
        const { Text: c, Dropdown: s, Button: r, Calendar: i } = Oc;
        let a,
            { value: u = { start: null, end: null } } = t,
            { id: d = et() } = t;
        const p = zn(u && u.start ? new Date(u.start) : new Date());
        f(e, p, (e) => n(5, (l = e)));
        const $ = zn(l);
        function m(e, t) {
            t && (p.set(new Date(t)), n(0, (u = { start: new Date(t), end: null })));
        }
        f(e, $, (e) => n(6, (o = e))),
            p.subscribe((e) => {
                const t = new Date(e);
                t.setDate(1), t.setMonth(t.getMonth() + 1), t.valueOf() != o.valueOf() && $.set(t);
            }),
            $.subscribe((e) => {
                const t = new Date(e);
                t.setDate(1), t.setMonth(t.getMonth() - 1), t.valueOf() != l.valueOf() && p.set(t);
            });
        let h,
            g = "";
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (u = e.value)), "id" in e && n(1, (d = e.id));
            }),
            (e.$$.update = () => {
                1 & e.$$.dirty && (u.start ? n(3, (g = Ve(u.start) + (u.end ? ` - ${Ve(u.end)}` : ""))) : n(3, (g = "")));
            }),
            [
                u,
                d,
                a,
                g,
                h,
                l,
                o,
                c,
                s,
                r,
                i,
                p,
                $,
                function (e) {
                    e.stopPropagation(), u && u.start && p.set(new Date(u.start)), n(2, (a = null));
                },
                m,
                function () {
                    n(0, (u = { start: null, end: null }));
                },
                function () {
                    p.set(u.start ? new Date(u.start) : new Date()), n(2, (a = !0));
                },
                function (e) {
                    (u = e), n(0, u);
                },
                function (e) {
                    (l = e), p.set(l);
                },
                function (e) {
                    (u = e), n(0, u);
                },
                function (e) {
                    (o = e), $.set(o);
                },
                (e) => m(0, new Date()),
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (h = e), n(4, h);
                    });
                },
            ]
        );
    }
    class Ni extends Ye {
        constructor(e) {
            super(), He(this, e, zi, ji, a, { value: 0, id: 1 });
        }
    }
    function Oi(e) {
        let t, n, l;
        function o(t) {
            e[7](t);
        }
        let c = { id: e[2], label: e[0], format: e[1] };
        return (
            void 0 !== e[3] && (c.value = e[3]),
            (t = new Ni({ props: c })),
            pe.push(() => Re(t, "value", o)),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, n) {
                    Ke(t, e, n), (l = !0);
                },
                p(e, [l]) {
                    const o = {};
                    4 & l && (o.id = e[2]), 1 & l && (o.label = e[0]), 2 & l && (o.format = e[1]), !n && 8 & l && ((n = !0), (o.value = e[3]), ye(() => (n = !1))), t.$set(o);
                },
                i(e) {
                    l || (Ae(t.$$.fragment, e), (l = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (l = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Fi(e, t, n) {
        let l,
            { start: o } = t,
            { end: c } = t,
            { label: s } = t,
            { format: r } = t,
            { id: i = Gc() } = t;
        const a = ii({ start: o, end: c }, (e) => {
            n(5, (o = e.start)), n(6, (c = e.end));
        });
        return (
            f(e, a, (e) => n(3, (l = e))),
            (e.$$set = (e) => {
                "start" in e && n(5, (o = e.start)), "end" in e && n(6, (c = e.end)), "label" in e && n(0, (s = e.label)), "format" in e && n(1, (r = e.format)), "id" in e && n(2, (i = e.id));
            }),
            (e.$$.update = () => {
                96 & e.$$.dirty && a.reset({ start: o, end: c });
            }),
            [
                s,
                r,
                i,
                l,
                a,
                o,
                c,
                function (e) {
                    (l = e), a.set(l);
                },
            ]
        );
    }
    class qi extends Ye {
        constructor(e) {
            super(), He(this, e, Fi, Oi, a, { start: 5, end: 6, label: 0, format: 1, id: 2 });
        }
    }
    function Ri(e, t, n) {
        const l = e.slice();
        return (l[35] = t[n]), (l[36] = t), (l[37] = n), l;
    }
    function Pi(e) {
        let t,
            l = e[16]("Close") + "";
        return {
            c() {
                t = O(l);
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function Ki(e) {
        let t, n;
        return (
            (t = new e[4]({ props: { block: !0, type: "primary", click: e[17], $$slots: { default: [Ui] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    512 & n[1] && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Ui(e) {
        let t,
            l = e[16]("Save") + "";
        return {
            c() {
                t = O(l);
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function Hi(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[2];
        const s = (e) => e[35].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = Ri(e, c, t),
                r = s(n);
            o.set(r, (l[t] = $a(r, n)));
        }
        return {
            c() {
                for (let e = 0; e < l.length; e += 1) l[e].c();
                t = q();
            },
            m(e, o) {
                for (let t = 0; t < l.length; t += 1) l[t].m(e, o);
                T(e, t, o), (n = !0);
            },
            p(e, n) {
                (327692 & n[0]) | (384 & n[1]) && ((c = e[2]), De(), (l = Oe(l, n, s, 1, e, c, o, t.parentNode, Ne, $a, t, Ri)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                for (let t = 0; t < l.length; t += 1) l[t].d(e);
                e && L(t);
            },
        };
    }
    function Yi(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[16](e[35].label), position: "top", $$slots: { default: [ta] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[16](e[35].label)), (4 & n[0]) | (512 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Bi(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[16](e[35].label), position: "top", $$slots: { default: [na, ({ id: e }) => ({ 38: e }), ({ id: e }) => [0, e ? 128 : 0]] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[16](e[35].label)), (12 & n[0]) | (640 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Gi(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[16](e[35].label), position: "top", $$slots: { default: [la, ({ id: e }) => ({ 38: e }), ({ id: e }) => [0, e ? 128 : 0]] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[16](e[35].label)), (12 & n[0]) | (640 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Ji(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[35].label, position: "top", $$slots: { default: [ca, ({ id: e }) => ({ 38: e }), ({ id: e }) => [0, e ? 128 : 0]] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[35].label), (12 & n[0]) | (640 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Vi(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[16](e[35].label), position: "top", $$slots: { default: [sa, ({ id: e }) => ({ 38: e }), ({ id: e }) => [0, e ? 128 : 0]] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[16](e[35].label)), (12 & n[0]) | (640 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Xi(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[16](e[35].label), position: "top", $$slots: { default: [ra, ({ id: e }) => ({ 38: e }), ({ id: e }) => [0, e ? 128 : 0]] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[16](e[35].label)), (12 & n[0]) | (640 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Qi(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[35].label, position: "top", $$slots: { default: [ua, ({ id: e }) => ({ 38: e }), ({ id: e }) => [0, e ? 128 : 0]] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[35].label), (12 & n[0]) | (640 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Wi(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[35].label, position: "top", $$slots: { default: [da, ({ id: e }) => ({ 38: e }), ({ id: e }) => [0, e ? 128 : 0]] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[35].label), (12 & n[0]) | (640 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Zi(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[16](e[35].label), position: "top", $$slots: { default: [pa, ({ id: e }) => ({ 38: e }), ({ id: e }) => [0, e ? 128 : 0]] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[16](e[35].label)), (12 & n[0]) | (640 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function ea(e) {
        let t, n;
        return (
            (t = new e[7]({ props: { label: e[16](e[35].label), position: "top", $$slots: { default: [fa, ({ id: e }) => ({ 38: e }), ({ id: e }) => [0, e ? 128 : 0]] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.label = e[16](e[35].label)), (12 & n[0]) | (640 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function ta(e) {
        let t, n, l;
        return (
            (t = new Di({ props: { field: e[35], values: e[18] } })),
            {
                c() {
                    Pe(t.$$.fragment), (n = F());
                },
                m(e, o) {
                    Ke(t, e, o), T(e, n, o), (l = !0);
                },
                p(e, n) {
                    const l = {};
                    4 & n[0] && (l.field = e[35]), t.$set(l);
                },
                i(e) {
                    l || (Ae(t.$$.fragment, e), (l = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (l = !1);
                },
                d(e) {
                    Ue(t, e), e && L(n);
                },
            }
        );
    }
    function na(e) {
        let t, n, l, o, c;
        function s(t) {
            e[31](t, e[35]);
        }
        function r(t) {
            e[32](t, e[35]);
        }
        let i = { id: e[38], label: e[16](e[35].label), format: Es };
        return (
            void 0 !== e[3][e[35].key.start] && (i.start = e[3][e[35].key.start]),
            void 0 !== e[3][e[35].key.end] && (i.end = e[3][e[35].key.end]),
            (t = new qi({ props: i })),
            pe.push(() => Re(t, "start", s)),
            pe.push(() => Re(t, "end", r)),
            {
                c() {
                    Pe(t.$$.fragment), (o = F());
                },
                m(e, n) {
                    Ke(t, e, n), T(e, o, n), (c = !0);
                },
                p(o, c) {
                    e = o;
                    const s = {};
                    128 & c[1] && (s.id = e[38]),
                        4 & c[0] && (s.label = e[16](e[35].label)),
                        !n && 12 & c[0] && ((n = !0), (s.start = e[3][e[35].key.start]), ye(() => (n = !1))),
                        !l && 12 & c[0] && ((l = !0), (s.end = e[3][e[35].key.end]), ye(() => (l = !1))),
                        t.$set(s);
                },
                i(e) {
                    c || (Ae(t.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (c = !1);
                },
                d(e) {
                    Ue(t, e), e && L(o);
                },
            }
        );
    }
    function la(e) {
        let t, n, l, c;
        const s = [{ id: e[38] }, { label: e[16](e[35].label) }, { format: Es }, e[35].config];
        function r(t) {
            e[30](t, e[35]);
        }
        let i = {};
        for (let e = 0; e < s.length; e += 1) i = o(i, s[e]);
        return (
            void 0 !== e[3][e[35].key] && (i.value = e[3][e[35].key]),
            (t = new e[11]({ props: i })),
            pe.push(() => Re(t, "value", r)),
            {
                c() {
                    Pe(t.$$.fragment), (l = F());
                },
                m(e, n) {
                    Ke(t, e, n), T(e, l, n), (c = !0);
                },
                p(l, o) {
                    e = l;
                    const c = (65540 & o[0]) | (128 & o[1]) ? Fe(s, [128 & o[1] && { id: e[38] }, 65540 & o[0] && { label: e[16](e[35].label) }, 0 & o && { format: Es }, 4 & o[0] && qe(e[35].config)]) : {};
                    !n && 12 & o[0] && ((n = !0), (c.value = e[3][e[35].key]), ye(() => (n = !1))), t.$set(c);
                },
                i(e) {
                    c || (Ae(t.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (c = !1);
                },
                d(e) {
                    Ue(t, e), e && L(l);
                },
            }
        );
    }
    function oa(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r = e[39].label + "";
        return (
            (n = new ar({ props: { data: e[39] } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = z("span")), (c = O(r)), U(o, "class", "multiselect-label svelte-1s2o9kq"), U(t, "class", "multiselect-option svelte-1s2o9kq");
                },
                m(e, r) {
                    T(e, t, r), Ke(n, t, null), I(t, l), I(t, o), I(o, c), (s = !0);
                },
                p(e, t) {
                    const l = {};
                    256 & t[1] && (l.data = e[39]), n.$set(l), (!s || 256 & t[1]) && r !== (r = e[39].label + "") && Y(c, r);
                },
                i(e) {
                    s || (Ae(n.$$.fragment, e), (s = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (s = !1);
                },
                d(e) {
                    e && L(t), Ue(n);
                },
            }
        );
    }
    function ca(e) {
        let t, n, l, c;
        const s = [{ id: e[38] }, { options: e[35].values }, { canDelete: !0 }, e[35].config];
        function r(t) {
            e[29](t, e[35]);
        }
        let i = { $$slots: { default: [oa, ({ option: e }) => ({ 39: e }), ({ option: e }) => [0, e ? 256 : 0]] }, $$scope: { ctx: e } };
        for (let e = 0; e < s.length; e += 1) i = o(i, s[e]);
        return (
            void 0 !== e[3][e[35].key] && (i.selected = e[3][e[35].key]),
            (t = new e[9]({ props: i })),
            pe.push(() => Re(t, "selected", r)),
            {
                c() {
                    Pe(t.$$.fragment), (l = F());
                },
                m(e, n) {
                    Ke(t, e, n), T(e, l, n), (c = !0);
                },
                p(l, o) {
                    e = l;
                    const c = (4 & o[0]) | (128 & o[1]) ? Fe(s, [128 & o[1] && { id: e[38] }, 4 & o[0] && { options: e[35].values }, s[2], 4 & o[0] && qe(e[35].config)]) : {};
                    768 & o[1] && (c.$$scope = { dirty: o, ctx: e }), !n && 12 & o[0] && ((n = !0), (c.selected = e[3][e[35].key]), ye(() => (n = !1))), t.$set(c);
                },
                i(e) {
                    c || (Ae(t.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (c = !1);
                },
                d(e) {
                    Ue(t, e), e && L(l);
                },
            }
        );
    }
    function sa(e) {
        let t, n, l, c;
        const s = [{ id: e[38] }, { colors: e[35].values }, e[35].config];
        function r(t) {
            e[28](t, e[35]);
        }
        let i = {};
        for (let e = 0; e < s.length; e += 1) i = o(i, s[e]);
        return (
            void 0 !== e[3][e[35].key] && (i.value = e[3][e[35].key]),
            (t = new e[12]({ props: i })),
            pe.push(() => Re(t, "value", r)),
            {
                c() {
                    Pe(t.$$.fragment), (l = F());
                },
                m(e, n) {
                    Ke(t, e, n), T(e, l, n), (c = !0);
                },
                p(l, o) {
                    e = l;
                    const c = (4 & o[0]) | (128 & o[1]) ? Fe(s, [128 & o[1] && { id: e[38] }, 4 & o[0] && { colors: e[35].values }, 4 & o[0] && qe(e[35].config)]) : {};
                    !n && 12 & o[0] && ((n = !0), (c.value = e[3][e[35].key]), ye(() => (n = !1))), t.$set(c);
                },
                i(e) {
                    c || (Ae(t.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (c = !1);
                },
                d(e) {
                    Ue(t, e), e && L(l);
                },
            }
        );
    }
    function ra(e) {
        let t, n, l, c;
        const s = [{ id: e[38] }, { options: e[35].values }, e[35].config];
        function r(t) {
            e[27](t, e[35]);
        }
        let i = {};
        for (let e = 0; e < s.length; e += 1) i = o(i, s[e]);
        return (
            void 0 !== e[3][e[35].key] && (i.value = e[3][e[35].key]),
            (t = new e[8]({ props: i })),
            pe.push(() => Re(t, "value", r)),
            {
                c() {
                    Pe(t.$$.fragment), (l = F());
                },
                m(e, n) {
                    Ke(t, e, n), T(e, l, n), (c = !0);
                },
                p(l, o) {
                    e = l;
                    const c = (4 & o[0]) | (128 & o[1]) ? Fe(s, [128 & o[1] && { id: e[38] }, 4 & o[0] && { options: e[35].values }, 4 & o[0] && qe(e[35].config)]) : {};
                    !n && 12 & o[0] && ((n = !0), (c.value = e[3][e[35].key]), ye(() => (n = !1))), t.$set(c);
                },
                i(e) {
                    c || (Ae(t.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (c = !1);
                },
                d(e) {
                    Ue(t, e), e && L(l);
                },
            }
        );
    }
    function ia(e) {
        let t;
        return {
            c() {
                (t = z("div")), U(t, "class", "color svelte-1s2o9kq"), G(t, "background", e[39].color);
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, n) {
                256 & n[1] && G(t, "background", e[39].color);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function aa(e) {
        let t,
            n,
            l,
            o = e[39].label + "",
            c = e[39].color && ia(e);
        return {
            c() {
                (t = z("div")), c && c.c(), (n = F()), (l = O(o)), U(t, "class", "combo-option svelte-1s2o9kq");
            },
            m(e, o) {
                T(e, t, o), c && c.m(t, null), I(t, n), I(t, l);
            },
            p(e, s) {
                e[39].color ? (c ? c.p(e, s) : ((c = ia(e)), c.c(), c.m(t, n))) : c && (c.d(1), (c = null)), 256 & s[1] && o !== (o = e[39].label + "") && Y(l, o);
            },
            d(e) {
                e && L(t), c && c.d();
            },
        };
    }
    function ua(e) {
        let t, n, l, c;
        const s = [{ id: e[38] }, { options: e[35].values }, e[35].config];
        function r(t) {
            e[26](t, e[35]);
        }
        let i = { $$slots: { default: [aa, ({ option: e }) => ({ 39: e }), ({ option: e }) => [0, e ? 256 : 0]] }, $$scope: { ctx: e } };
        for (let e = 0; e < s.length; e += 1) i = o(i, s[e]);
        return (
            void 0 !== e[3][e[35].key] && (i.value = e[3][e[35].key]),
            (t = new e[10]({ props: i })),
            pe.push(() => Re(t, "value", r)),
            {
                c() {
                    Pe(t.$$.fragment), (l = F());
                },
                m(e, n) {
                    Ke(t, e, n), T(e, l, n), (c = !0);
                },
                p(l, o) {
                    e = l;
                    const c = (4 & o[0]) | (128 & o[1]) ? Fe(s, [128 & o[1] && { id: e[38] }, 4 & o[0] && { options: e[35].values }, 4 & o[0] && qe(e[35].config)]) : {};
                    768 & o[1] && (c.$$scope = { dirty: o, ctx: e }), !n && 12 & o[0] && ((n = !0), (c.value = e[3][e[35].key]), ye(() => (n = !1))), t.$set(c);
                },
                i(e) {
                    c || (Ae(t.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (c = !1);
                },
                d(e) {
                    Ue(t, e), e && L(l);
                },
            }
        );
    }
    function da(e) {
        let t, n, l, c;
        const s = [{ id: e[38] }, { min: 0 }, e[35].config];
        function r(t) {
            e[25](t, e[35]);
        }
        let i = {};
        for (let e = 0; e < s.length; e += 1) i = o(i, s[e]);
        return (
            void 0 !== e[3][e[35].key] && (i.value = e[3][e[35].key]),
            (t = new e[13]({ props: i })),
            pe.push(() => Re(t, "value", r)),
            {
                c() {
                    Pe(t.$$.fragment), (l = F());
                },
                m(e, n) {
                    Ke(t, e, n), T(e, l, n), (c = !0);
                },
                p(l, o) {
                    e = l;
                    const c = (4 & o[0]) | (128 & o[1]) ? Fe(s, [128 & o[1] && { id: e[38] }, s[1], 4 & o[0] && qe(e[35].config)]) : {};
                    !n && 12 & o[0] && ((n = !0), (c.value = e[3][e[35].key]), ye(() => (n = !1))), t.$set(c);
                },
                i(e) {
                    c || (Ae(t.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (c = !1);
                },
                d(e) {
                    Ue(t, e), e && L(l);
                },
            }
        );
    }
    function pa(e) {
        let t, n, l, c;
        const s = [{ id: e[38] }, e[35].config];
        function r(t) {
            e[24](t, e[35]);
        }
        let i = {};
        for (let e = 0; e < s.length; e += 1) i = o(i, s[e]);
        return (
            void 0 !== e[3][e[35].key] && (i.value = e[3][e[35].key]),
            (t = new e[6]({ props: i })),
            pe.push(() => Re(t, "value", r)),
            {
                c() {
                    Pe(t.$$.fragment), (l = F());
                },
                m(e, n) {
                    Ke(t, e, n), T(e, l, n), (c = !0);
                },
                p(l, o) {
                    e = l;
                    const c = (4 & o[0]) | (128 & o[1]) ? Fe(s, [128 & o[1] && { id: e[38] }, 4 & o[0] && qe(e[35].config)]) : {};
                    !n && 12 & o[0] && ((n = !0), (c.value = e[3][e[35].key]), ye(() => (n = !1))), t.$set(c);
                },
                i(e) {
                    c || (Ae(t.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (c = !1);
                },
                d(e) {
                    Ue(t, e), e && L(l);
                },
            }
        );
    }
    function fa(e) {
        let t, n, l, c;
        const s = [{ id: e[38] }, { focus: !0 }, e[35].config];
        function r(t) {
            e[23](t, e[35]);
        }
        let i = {};
        for (let e = 0; e < s.length; e += 1) i = o(i, s[e]);
        return (
            void 0 !== e[3][e[35].key] && (i.value = e[3][e[35].key]),
            (t = new e[5]({ props: i })),
            pe.push(() => Re(t, "value", r)),
            {
                c() {
                    Pe(t.$$.fragment), (l = F());
                },
                m(e, n) {
                    Ke(t, e, n), T(e, l, n), (c = !0);
                },
                p(l, o) {
                    e = l;
                    const c = (4 & o[0]) | (128 & o[1]) ? Fe(s, [128 & o[1] && { id: e[38] }, s[1], 4 & o[0] && qe(e[35].config)]) : {};
                    !n && 12 & o[0] && ((n = !0), (c.value = e[3][e[35].key]), ye(() => (n = !1))), t.$set(c);
                },
                i(e) {
                    c || (Ae(t.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (c = !1);
                },
                d(e) {
                    Ue(t, e), e && L(l);
                },
            }
        );
    }
    function $a(e, t) {
        let n, l, o, c, s;
        const r = [ea, Zi, Wi, Qi, Xi, Vi, Ji, Gi, Bi, Yi],
            i = [];
        function a(e, t) {
            return "text" === e[35].type
                ? 0
                : "textarea" === e[35].type
                    ? 1
                    : "progress" === e[35].type
                        ? 2
                        : "combo" === e[35].type
                            ? 3
                            : "select" === e[35].type
                                ? 4
                                : "color" === e[35].type
                                    ? 5
                                    : "multiselect" === e[35].type
                                        ? 6
                                        : "date" === e[35].type
                                            ? 7
                                            : "dateRange" === e[35].type
                                                ? 8
                                                : "files" === e[35].type
                                                    ? 9
                                                    : -1;
        }
        return (
            ~(l = a(t)) && (o = i[l] = r[l](t)),
            {
                key: e,
                first: null,
                c() {
                    (n = q()), o && o.c(), (c = q()), (this.first = n);
                },
                m(e, t) {
                    T(e, n, t), ~l && i[l].m(e, t), T(e, c, t), (s = !0);
                },
                p(e, n) {
                    let s = l;
                    (l = a((t = e))),
                        l === s
                            ? ~l && i[l].p(t, n)
                            : (o &&
                                (De(),
                                    Ee(i[s], 1, 1, () => {
                                        i[s] = null;
                                    }),
                                    Ie()),
                                ~l ? ((o = i[l]), o ? o.p(t, n) : ((o = i[l] = r[l](t)), o.c()), Ae(o, 1), o.m(c.parentNode, c)) : (o = null));
                },
                i(e) {
                    s || (Ae(o), (s = !0));
                },
                o(e) {
                    Ee(o), (s = !1);
                },
                d(e) {
                    e && L(n), ~l && i[l].d(e), e && L(c);
                },
            }
        );
    }
    function ma(e) {
        let t, n, l, o, c, s, r, i;
        l = new e[4]({ props: { block: !0, click: e[19], $$slots: { default: [Pi] }, $$scope: { ctx: e } } });
        let a = !e[0] && Ki(e),
            u = e[1] && Hi(e);
        return {
            c() {
                (t = z("div")),
                    (n = z("div")),
                    Pe(l.$$.fragment),
                    (o = F()),
                    a && a.c(),
                    (c = F()),
                    u && u.c(),
                    U(n, "class", "editor-controls svelte-1s2o9kq"),
                    U(t, "class", "editor svelte-1s2o9kq"),
                    U(t, "data-kanban-id", es),
                    V(t, "editor-open", e[1]);
            },
            m(d, p) {
                T(d, t, p), I(t, n), Ke(l, n, null), I(n, o), a && a.m(n, null), I(t, c), u && u.m(t, null), (s = !0), r || ((i = R(t, "click", K(e[22]))), (r = !0));
            },
            p(e, o) {
                const c = {};
                512 & o[1] && (c.$$scope = { dirty: o, ctx: e }),
                    l.$set(c),
                    e[0]
                        ? a &&
                        (De(),
                            Ee(a, 1, 1, () => {
                                a = null;
                            }),
                            Ie())
                        : a
                            ? (a.p(e, o), 1 & o[0] && Ae(a, 1))
                            : ((a = Ki(e)), a.c(), Ae(a, 1), a.m(n, null)),
                    e[1]
                        ? u
                            ? (u.p(e, o), 2 & o[0] && Ae(u, 1))
                            : ((u = Hi(e)), u.c(), Ae(u, 1), u.m(t, null))
                        : u &&
                        (De(),
                            Ee(u, 1, 1, () => {
                                u = null;
                            }),
                            Ie()),
                    2 & o[0] && V(t, "editor-open", e[1]);
            },
            i(e) {
                s || (Ae(l.$$.fragment, e), Ae(a), Ae(u), (s = !0));
            },
            o(e) {
                Ee(l.$$.fragment, e), Ee(a), Ee(u), (s = !1);
            },
            d(e) {
                e && L(t), Ue(l), a && a.d(), u && u.d(), (r = !1), i();
            },
        };
    }
    function ha(e, t, n) {
        let l, o, c, s;
        const { Button: r, Text: i, Area: a, Field: u, Select: d, MultiSelect: p, Combo: $, DatePicker: m, ColorPicker: h, Slider: g } = Oc;
        let { autoSave: v = !0 } = t,
            { api: y } = t;
        const { selected: w, editorShape: b } = y.getReactiveState();
        f(e, w, (e) => n(21, (c = e))), f(e, b, (e) => n(2, (o = e)));
        const x = ae("wx-i18n").getGroup("kanban");
        let k;
        function S() {
            y.exec("update-card", { card: Object.assign({}, k), id: k.id });
        }
        const M = ii({}, (e) => {
            (k = e), v && S();
        });
        return (
            f(e, M, (e) => n(3, (s = e))),
            (e.$$set = (e) => {
                "autoSave" in e && n(0, (v = e.autoSave)), "api" in e && n(20, (y = e.api));
            }),
            (e.$$.update = () => {
                3145728 & e.$$.dirty[0] && n(1, (l = y.getCard(1 === (null == c ? void 0 : c.length) && c[0]))),
                    2 & e.$$.dirty[0] &&
                    M.reset(
                        (function (e) {
                            const t = Object.assign({}, e);
                            return (
                                o.forEach((e) => {
                                    void 0 === t[e.key] && (t[e.key] = "files" === e.type ? [] : "date" === e.type ? null : "");
                                }),
                                t
                            );
                        })(l)
                    );
            }),
            [
                v,
                l,
                o,
                s,
                r,
                i,
                a,
                u,
                d,
                p,
                $,
                m,
                h,
                g,
                w,
                b,
                x,
                S,
                M,
                function () {
                    y.exec("unselect-card", { id: l.id });
                },
                y,
                c,
                function (t) {
                    ue.call(this, e, t);
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key], t) && ((s[n.key] = t), M.set(s));
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key], t) && ((s[n.key] = t), M.set(s));
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key], t) && ((s[n.key] = t), M.set(s));
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key], t) && ((s[n.key] = t), M.set(s));
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key], t) && ((s[n.key] = t), M.set(s));
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key], t) && ((s[n.key] = t), M.set(s));
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key], t) && ((s[n.key] = t), M.set(s));
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key], t) && ((s[n.key] = t), M.set(s));
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key.start], t) && ((s[n.key.start] = t), M.set(s));
                },
                function (t, n) {
                    e.$$.not_equal(s[n.key.end], t) && ((s[n.key.end] = t), M.set(s));
                },
            ]
        );
    }
    class ga extends Ye {
        constructor(e) {
            super(), He(this, e, ha, ma, a, { autoSave: 0, api: 20 }, null, [-1, -1]);
        }
    }
    function va(e, t, n) {
        const l = e.slice();
        return (l[26] = t[n]), (l[27] = t), (l[28] = n), l;
    }
    function ya(e) {
        let t, n, l, o;
        function c() {
            return e[17](e[26]);
        }
        return {
            c() {
                (t = z("div")), U(t, "class", "collapsed-column svelte-vj2omq"), G(t, "left", e[6][e[26].id].offsetLeft + "px");
            },
            m(s, r) {
                T(s, t, r), l || ((o = [R(t, "click", c), k((n = zs.call(null, t, { container: e[3] })))]), (l = !0));
            },
            p(l, o) {
                (e = l), 65 & o && G(t, "left", e[6][e[26].id].offsetLeft + "px"), n && i(n.update) && 8 & o && n.update.call(null, { container: e[3] });
            },
            d(e) {
                e && L(t), (l = !1), r(o);
            },
        };
    }
    function wa(e) {
        let t,
            n = e[26].label + "";
        return {
            c() {
                t = O(n);
            },
            m(e, n) {
                T(e, t, n);
            },
            p(e, l) {
                1 & l && n !== (n = e[26].label + "") && Y(t, n);
            },
            d(e) {
                e && L(t);
            },
        };
    }
    function ba(e) {
        let t, n, l, o;
        return {
            c() {
                (t = z("input")), U(t, "type", "text"), U(t, "class", "input svelte-vj2omq"), (t.value = n = e[26].label);
            },
            m(n, c) {
                T(n, t, c), l || ((o = [R(t, "input", e[12]), R(t, "keypress", e[13]), R(t, "blur", e[10]), k(Ia.call(null, t))]), (l = !0));
            },
            p(e, l) {
                1 & l && n !== (n = e[26].label) && t.value !== n && (t.value = n);
            },
            d(e) {
                e && L(t), (l = !1), r(o);
            },
        };
    }
    function xa(e) {
        let t,
            n,
            l,
            o,
            c,
            s = Jc(e[5], e[26].id),
            r = e[26].limit && ka(e);
        function i() {
            return e[19](e[26]);
        }
        l = new Ys({ props: { name: "dots-h", click: i } });
        let a = s && Sa(e);
        return {
            c() {
                r && r.c(), (t = F()), (n = z("div")), Pe(l.$$.fragment), (o = F()), a && a.c(), U(n, "class", "menu svelte-vj2omq");
            },
            m(e, s) {
                r && r.m(e, s), T(e, t, s), T(e, n, s), Ke(l, n, null), I(n, o), a && a.m(n, null), (c = !0);
            },
            p(o, c) {
                (e = o)[26].limit ? (r ? r.p(e, c) : ((r = ka(e)), r.c(), r.m(t.parentNode, t))) : r && (r.d(1), (r = null));
                const u = {};
                1 & c && (u.click = i),
                    l.$set(u),
                    33 & c && (s = Jc(e[5], e[26].id)),
                    s
                        ? a
                            ? (a.p(e, c), 33 & c && Ae(a, 1))
                            : ((a = Sa(e)), a.c(), Ae(a, 1), a.m(n, null))
                        : a &&
                        (De(),
                            Ee(a, 1, 1, () => {
                                a = null;
                            }),
                            Ie());
            },
            i(e) {
                c || (Ae(l.$$.fragment, e), Ae(a), (c = !0));
            },
            o(e) {
                Ee(l.$$.fragment, e), Ee(a), (c = !1);
            },
            d(e) {
                r && r.d(e), e && L(t), e && L(n), Ue(l), a && a.d();
            },
        };
    }
    function ka(e) {
        let t,
            n,
            l,
            o,
            c,
            s = e[1][e[26].id].cardsCount + "",
            r = e[1][e[26].id].limit + "";
        return {
            c() {
                (t = O("(")), (n = O(s)), (l = O("/")), (o = O(r)), (c = O(")"));
            },
            m(e, s) {
                T(e, t, s), T(e, n, s), T(e, l, s), T(e, o, s), T(e, c, s);
            },
            p(e, t) {
                3 & t && s !== (s = e[1][e[26].id].cardsCount + "") && Y(n, s), 3 & t && r !== (r = e[1][e[26].id].limit + "") && Y(o, r);
            },
            d(e) {
                e && L(t), e && L(n), e && L(l), e && L(o), e && L(c);
            },
        };
    }
    function Sa(e) {
        let t, n, l, o, c, s;
        return (
            (n = new e[7]({ props: { cancel: e[20], width: "auto", $$slots: { default: [_a] }, $$scope: { ctx: e } } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), U(t, "class", "menu-wrap svelte-vj2omq"), G(t, "left", e[6][e[26].id] ?.offsetLeft + e[6][e[26].id] ?.offsetWidth - 30 + "px");
                },
                m(r, i) {
                    T(r, t, i), Ke(n, t, null), (o = !0), c || ((s = k((l = zs.call(null, t, { container: e[3] })))), (c = !0));
                },
                p(e, c) {
                    const s = {};
                    32 & c && (s.cancel = e[20]),
                        1073741825 & c && (s.$$scope = { dirty: c, ctx: e }),
                        n.$set(s),
                        (!o || 65 & c) && G(t, "left", e[6][e[26].id] ?.offsetLeft + e[6][e[26].id] ?.offsetWidth - 30 + "px"),
                        l && i(l.update) && 8 & c && l.update.call(null, { container: e[3] });
                },
                i(e) {
                    o || (Ae(n.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(t), Ue(n), (c = !1), s();
                },
            }
        );
    }
    function Ma(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r = e[29].label + "";
        return (
            (n = new Ys({ props: { name: e[29].icon } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), (l = F()), (o = z("span")), (c = O(r)), U(o, "class", "svelte-vj2omq"), U(t, "class", "menu-item svelte-vj2omq");
                },
                m(e, r) {
                    T(e, t, r), Ke(n, t, null), I(t, l), I(t, o), I(o, c), (s = !0);
                },
                p(e, t) {
                    const l = {};
                    536870912 & t && (l.name = e[29].icon), n.$set(l), (!s || 536870912 & t) && r !== (r = e[29].label + "") && Y(c, r);
                },
                i(e) {
                    s || (Ae(n.$$.fragment, e), (s = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (s = !1);
                },
                d(e) {
                    e && L(t), Ue(n);
                },
            }
        );
    }
    function _a(e) {
        let t, n;
        return (
            (t = new e[8]({
                props: {
                    click: e[16],
                    data: [
                        { icon: "edit", label: e[9]("Rename"), id: 1 },
                        { icon: "arrow-left", label: e[9]("Move left"), id: e[28] > 0 ? 3 : "wx-list-disabled" },
                        { icon: "arrow-right", label: e[9]("Move right"), id: e[28] < e[0].length - 1 ? 4 : "wx-list-disabled" },
                        { icon: "delete", label: e[9]("Delete"), id: 2 },
                    ],
                    $$slots: { default: [Ma, ({ obj: e }) => ({ 29: e }), ({ obj: e }) => (e ? 536870912 : 0)] },
                    $$scope: { ctx: e },
                },
            })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    1 & n &&
                        (l.data = [
                            { icon: "edit", label: e[9]("Rename"), id: 1 },
                            { icon: "arrow-left", label: e[9]("Move left"), id: e[28] > 0 ? 3 : "wx-list-disabled" },
                            { icon: "arrow-right", label: e[9]("Move right"), id: e[28] < e[0].length - 1 ? 4 : "wx-list-disabled" },
                            { icon: "delete", label: e[9]("Delete"), id: 2 },
                        ]),
                        1610612736 & n && (l.$$scope = { dirty: n, ctx: e }),
                        t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Ca(e, t) {
        let n,
            l,
            o,
            c,
            s,
            r,
            i,
            a,
            u,
            d,
            p,
            f,
            $,
            m = !Jc(t[4], t[26].id) && t[2],
            h = t[26],
            g = t[26].collapsed && t[6][t[26].id] && t[3] && ya(t);
        function v() {
            return t[18](t[26]);
        }
        function y(e, t) {
            return (null == i || 17 & t) && (i = !!Jc(e[4], e[26].id)), i ? ba : wa;
        }
        c = new Ys({ props: { name: t[26].collapsed ? "angle-right" : "angle-left", clickable: !0, click: v } });
        let w = y(t, -1),
            b = w(t),
            x = m && xa(t);
        function k() {
            return t[21](t[26]);
        }
        const S = () => t[22](n, h),
            M = () => t[22](null, h);
        return {
            key: e,
            first: null,
            c() {
                (n = z("div")),
                    g && g.c(),
                    (l = F()),
                    (o = z("div")),
                    Pe(c.$$.fragment),
                    (s = F()),
                    (r = z("div")),
                    b.c(),
                    (a = F()),
                    x && x.c(),
                    (u = F()),
                    (d = F()),
                    U(o, "class", "collapse-icon svelte-vj2omq"),
                    U(r, "class", "label svelte-vj2omq"),
                    U(n, "class", "column svelte-vj2omq"),
                    V(n, "over-limit", t[1][t[26].id].isOverLimit),
                    V(n, "collapsed", t[26].collapsed),
                    (this.first = n);
            },
            m(e, t) {
                T(e, n, t), g && g.m(n, null), I(n, l), I(n, o), Ke(c, o, null), I(n, s), I(n, r), b.m(r, null), I(r, a), x && x.m(r, null), I(n, u), I(n, d), S(), (p = !0), f || (($ = R(r, "dblclick", k)), (f = !0));
            },
            p(e, o) {
                (t = e)[26].collapsed && t[6][t[26].id] && t[3] ? (g ? g.p(t, o) : ((g = ya(t)), g.c(), g.m(n, l))) : g && (g.d(1), (g = null));
                const s = {};
                1 & o && (s.name = t[26].collapsed ? "angle-right" : "angle-left"),
                    1 & o && (s.click = v),
                    c.$set(s),
                    w === (w = y(t, o)) && b ? b.p(t, o) : (b.d(1), (b = w(t)), b && (b.c(), b.m(r, a))),
                    21 & o && (m = !Jc(t[4], t[26].id) && t[2]),
                    m
                        ? x
                            ? (x.p(t, o), 21 & o && Ae(x, 1))
                            : ((x = xa(t)), x.c(), Ae(x, 1), x.m(r, null))
                        : x &&
                        (De(),
                            Ee(x, 1, 1, () => {
                                x = null;
                            }),
                            Ie()),
                    h !== t[26] && (M(), (h = t[26]), S()),
                    3 & o && V(n, "over-limit", t[1][t[26].id].isOverLimit),
                    1 & o && V(n, "collapsed", t[26].collapsed);
            },
            i(e) {
                p || (Ae(c.$$.fragment, e), Ae(x), (p = !0));
            },
            o(e) {
                Ee(c.$$.fragment, e), Ee(x), (p = !1);
            },
            d(e) {
                e && L(n), g && g.d(), Ue(c), b.d(), x && x.d(), M(), (f = !1), $();
            },
        };
    }
    function Da(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[0];
        const s = (e) => e[26].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = va(e, c, t),
                r = s(n);
            o.set(r, (l[t] = Ca(r, n)));
        }
        return {
            c() {
                t = z("div");
                for (let e = 0; e < l.length; e += 1) l[e].c();
                U(t, "class", "header svelte-vj2omq");
            },
            m(e, o) {
                T(e, t, o);
                for (let e = 0; e < l.length; e += 1) l[e].m(t, null);
                n = !0;
            },
            p(e, [n]) {
                537001599 & n && ((c = e[0]), De(), (l = Oe(l, n, s, 1, e, c, o, t, Ne, Ca, null, va)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                e && L(t);
                for (let e = 0; e < l.length; e += 1) l[e].d();
            },
        };
    }
    function Ia(e) {
        e.focus();
    }
    function Aa(e, t, n) {
        const { Dropdown: l, List: o } = Oc;
        let { columns: c } = t,
            { areasMeta: s } = t,
            { edit: r = !0 } = t,
            { contentEl: i } = t;
        const a = ae("wx-i18n").getGroup("kanban"),
            u = re();
        let d,
            p = null,
            f = null;
        function $() {
            p && (null == f ? void 0 : f.trim()) && u("action", { action: "update-column", data: { id: p, column: { label: f } } }), n(4, (p = null)), (f = null);
        }
        function m(e) {
            r && n(4, (p = e));
        }
        function h(e, t) {
            u("action", { action: "update-column", data: { id: e, column: { collapsed: !t } } });
        }
        function g(e) {
            n(5, (d = e));
        }
        function v(e) {
            var t;
            const n = c.findIndex((e) => e.id === d),
                l = null === (t = c["left" === e ? n - 1 : n + 2]) || void 0 === t ? void 0 : t.id;
            u("action", { action: "move-column", data: { id: d, before: l } });
        }
        let y = {};
        return (
            (e.$$set = (e) => {
                "columns" in e && n(0, (c = e.columns)), "areasMeta" in e && n(1, (s = e.areasMeta)), "edit" in e && n(2, (r = e.edit)), "contentEl" in e && n(3, (i = e.contentEl));
            }),
            [
                c,
                s,
                r,
                i,
                p,
                d,
                y,
                l,
                o,
                a,
                $,
                m,
                function (e) {
                    f = e.target.value;
                },
                function (e) {
                    13 === e.charCode && $();
                },
                h,
                g,
                function (e) {
                    1 === e && m(d), 2 === e && u("action", { action: "delete-column", data: { id: d } }), 3 === e && v("left"), 4 === e && v("right"), n(5, (d = null));
                },
                (e) => h(e.id, e.collapsed),
                (e) => h(e.id, e.collapsed),
                (e) => g(e.id),
                () => {
                    n(5, (d = null));
                },
                (e) => m(e.id),
                function (e, t) {
                    pe[e ? "unshift" : "push"](() => {
                        (y[t.id] = e), n(6, y), n(0, c);
                    });
                },
            ]
        );
    }
    class Ea extends Ye {
        constructor(e) {
            super(), He(this, e, Aa, Da, a, { columns: 0, areasMeta: 1, edit: 2, contentEl: 3 });
        }
    }
    function Ta(e, t, n) {
        const l = e.slice();
        return (l[54] = t[n]), l;
    }
    function La(e, t, n) {
        const l = e.slice();
        return (l[57] = t[n]), l;
    }
    function ja(e, t) {
        let n, l, o;
        return (
            (l = new si({
                props: {
                    column: t[57],
                    row: t[54],
                    overCardId: t[23],
                    overColId: t[24],
                    movedCardId: t[5],
                    movedCardCoords: t[17],
                    selected: t[18],
                    dropAreasCoords: t[25],
                    cards: t[26][Xc(t[57].id, t[54].id)],
                    cardShape: t[4],
                    cardTemplate: t[1],
                    cardsMeta: t[27],
                    areasMeta: t[20],
                    add: t[13],
                },
            })),
            l.$on("action", t[28]),
            {
                key: e,
                first: null,
                c() {
                    (n = q()), Pe(l.$$.fragment), (this.first = n);
                },
                m(e, t) {
                    T(e, n, t), Ke(l, e, t), (o = !0);
                },
                p(e, n) {
                    t = e;
                    const o = {};
                    524288 & n[0] && (o.column = t[57]),
                        2097152 & n[0] && (o.row = t[54]),
                        8388608 & n[0] && (o.overCardId = t[23]),
                        16777216 & n[0] && (o.overColId = t[24]),
                        32 & n[0] && (o.movedCardId = t[5]),
                        131072 & n[0] && (o.movedCardCoords = t[17]),
                        262144 & n[0] && (o.selected = t[18]),
                        33554432 & n[0] && (o.dropAreasCoords = t[25]),
                        69730304 & n[0] && (o.cards = t[26][Xc(t[57].id, t[54].id)]),
                        16 & n[0] && (o.cardShape = t[4]),
                        2 & n[0] && (o.cardTemplate = t[1]),
                        134217728 & n[0] && (o.cardsMeta = t[27]),
                        1048576 & n[0] && (o.areasMeta = t[20]),
                        8192 & n[0] && (o.add = t[13]),
                        l.$set(o);
                },
                i(e) {
                    o || (Ae(l.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(l.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(n), Ue(l, e);
                },
            }
        );
    }
    function za(e) {
        let t,
            n,
            l = [],
            o = new Map(),
            c = e[19];
        const s = (e) => e[57].id;
        for (let t = 0; t < c.length; t += 1) {
            let n = La(e, c, t),
                r = s(n);
            o.set(r, (l[t] = ja(r, n)));
        }
        return {
            c() {
                for (let e = 0; e < l.length; e += 1) l[e].c();
                t = F();
            },
            m(e, o) {
                for (let t = 0; t < l.length; t += 1) l[t].m(e, o);
                T(e, t, o), (n = !0);
            },
            p(e, n) {
                532553778 & n[0] && ((c = e[19]), De(), (l = Oe(l, n, s, 1, e, c, o, t.parentNode, Ne, ja, t, La)), Ie());
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < c.length; e += 1) Ae(l[e]);
                    n = !0;
                }
            },
            o(e) {
                for (let e = 0; e < l.length; e += 1) Ee(l[e]);
                n = !1;
            },
            d(e) {
                for (let t = 0; t < l.length; t += 1) l[t].d(e);
                e && L(t);
            },
        };
    }
    function Na(e, t) {
        let n, l, o;
        return (
            (l = new lr({ props: { row: t[54], rows: t[21], collapsable: !!t[22], contentEl: t[16], edit: t[3], $$slots: { default: [za] }, $$scope: { ctx: t } } })),
            l.$on("action", t[28]),
            {
                key: e,
                first: null,
                c() {
                    (n = q()), Pe(l.$$.fragment), (this.first = n);
                },
                m(e, t) {
                    T(e, n, t), Ke(l, e, t), (o = !0);
                },
                p(e, n) {
                    t = e;
                    const o = {};
                    2097152 & n[0] && (o.row = t[54]),
                        2097152 & n[0] && (o.rows = t[21]),
                        4194304 & n[0] && (o.collapsable = !!t[22]),
                        65536 & n[0] && (o.contentEl = t[16]),
                        8 & n[0] && (o.edit = t[3]),
                        (264118322 & n[0]) | (536870912 & n[1]) && (o.$$scope = { dirty: n, ctx: t }),
                        l.$set(o);
                },
                i(e) {
                    o || (Ae(l.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(l.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(n), Ue(l, e);
                },
            }
        );
    }
    function Oa(e) {
        let t, n;
        return (
            (t = new ga({ props: { api: e[2], autoSave: e[0] } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    1 & n[0] && (l.autoSave = e[0]), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Fa(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            a,
            u,
            d,
            p,
            f,
            $,
            m = [],
            h = new Map();
        (o = new Ea({ props: { columns: e[19], areasMeta: e[20], contentEl: e[16], edit: e[3] } })), o.$on("action", e[28]);
        let g = e[21];
        const v = (e) => e[54].id;
        for (let t = 0; t < g.length; t += 1) {
            let n = Ta(e, g, t),
                l = v(n);
            h.set(l, (m[t] = Na(l, n)));
        }
        let y = e[3] && !e[5] && Oa(e);
        return {
            c() {
                (t = z("div")), (n = z("div")), (l = z("div")), Pe(o.$$.fragment), (c = F());
                for (let e = 0; e < m.length; e += 1) m[e].c();
                (a = F()), y && y.c(), U(l, "class", "content svelte-1lpxmvc"), U(n, "class", "content-wrapper svelte-1lpxmvc"), U(n, "data-kanban-id", ts), U(t, "class", "kanban svelte-1lpxmvc"), V(t, "dragged", !!e[5]);
            },
            m(r, i) {
                T(r, t, i), I(t, n), I(n, l), Ke(o, l, null), I(l, c);
                for (let e = 0; e < m.length; e += 1) m[e].m(l, null);
                e[46](l),
                    I(t, a),
                    y && y.m(t, null),
                    (p = !0),
                    f ||
                    (($ = [
                        k((s = Is.call(null, l, { api: e[2], readonly: !1 === e[14] }))),
                        k((u = Ds.call(null, t, { onAction: e[35], api: e[2], readonly: !1 === e[15] }))),
                        k((d = As.call(null, t, { onAction: e[36], readonly: !1 === e[3] }))),
                    ]),
                        (f = !0));
            },
            p(e, n) {
                const c = {};
                524288 & n[0] && (c.columns = e[19]),
                    1048576 & n[0] && (c.areasMeta = e[20]),
                    65536 & n[0] && (c.contentEl = e[16]),
                    8 & n[0] && (c.edit = e[3]),
                    o.$set(c),
                    536813626 & n[0] && ((g = e[21]), De(), (m = Oe(m, n, v, 1, e, g, h, l, Ne, Na, null, Ta)), Ie()),
                    s && i(s.update) && 16384 & n[0] && s.update.call(null, { api: e[2], readonly: !1 === e[14] }),
                    e[3] && !e[5]
                        ? y
                            ? (y.p(e, n), 40 & n[0] && Ae(y, 1))
                            : ((y = Oa(e)), y.c(), Ae(y, 1), y.m(t, null))
                        : y &&
                        (De(),
                            Ee(y, 1, 1, () => {
                                y = null;
                            }),
                            Ie()),
                    u && i(u.update) && 32768 & n[0] && u.update.call(null, { onAction: e[35], api: e[2], readonly: !1 === e[15] }),
                    d && i(d.update) && 8 & n[0] && d.update.call(null, { onAction: e[36], readonly: !1 === e[3] }),
                    32 & n[0] && V(t, "dragged", !!e[5]);
            },
            i(e) {
                if (!p) {
                    Ae(o.$$.fragment, e);
                    for (let e = 0; e < g.length; e += 1) Ae(m[e]);
                    Ae(y), (p = !0);
                }
            },
            o(e) {
                Ee(o.$$.fragment, e);
                for (let e = 0; e < m.length; e += 1) Ee(m[e]);
                Ee(y), (p = !1);
            },
            d(n) {
                n && L(t), Ue(o);
                for (let e = 0; e < m.length; e += 1) m[e].d();
                e[46](null), y && y.d(), (f = !1), r($);
            },
        };
    }
    function qa(e) {
        let t, n;
        return (
            (t = new Kc({ props: { $$slots: { default: [Fa] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    (268427323 & n[0]) | (536870912 & n[1]) && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Ra(e, t, l) {
        let o,
            c,
            s,
            r,
            i,
            a,
            u,
            d,
            $,
            m,
            h,
            g,
            v,
            y,
            w = n,
            b = n,
            k = n,
            S = n,
            M = n,
            _ = n,
            C = n;
        var D;
        e.$$.on_destroy.push(() => w()), e.$$.on_destroy.push(() => b()), e.$$.on_destroy.push(() => k()), e.$$.on_destroy.push(() => S()), e.$$.on_destroy.push(() => M()), e.$$.on_destroy.push(() => _()), e.$$.on_destroy.push(() => C());
        let { columns: I } = t,
            { rows: A = null } = t,
            { cards: E } = t,
            { cardShape: T = Wc } = t,
            { editorShape: L = null } = t,
            { editorAutoSave: j = !0 } = t,
            { cardTemplate: z = null } = t,
            { readonly: N = !1 } = t,
            { columnKey: O = "column" } = t,
            { rowKey: F = "" } = t;
        const q = ae("wx-i18n");
        q ? (null === (D = null == q ? void 0 : q.data) || void 0 === D ? void 0 : D.kanban) || q.extend(Os) : ie("wx-i18n", Ns(Os));
        const R = re();
        var P;
        ce(() => {
            if (!document.querySelector(".wx-portal")) {
                const e = document.createElement("div");
                e.classList.add("wx-portal"), e.classList.add("wx-meta-theme"), document.body.appendChild(e);
            }
        }),
            (P = () => {
                var e;
                null === (e = document.querySelector(".wx-portal")) || void 0 === e || e.remove();
            }),
            oe().$$.on_destroy.push(P);
        const K = new xs((e) => zn(e));
        let U = new Yc(R);
        K.out.setNext(U.exec);
        const H = (function (e, t) {
            let n = t;
            return {
                exec: e.in.exec.bind(e.in),
                on: e.out.on.bind(e.out),
                intercept: e.in.on.bind(e.in),
                getState: e.getState.bind(e),
                getReactiveState: e.getReactive.bind(e),
                setNext: (e) => {
                    n.setNext(e.exec), (n = e);
                },
                getStores: () => ({ data: e }),
                getCard: (t) => {
                    const { cards: n } = e.getState();
                    return n.find((e) => e.id == t);
                },
                serialize: () => {
                    const { cards: t, columns: n, rows: l } = e.getState();
                    return { cards: t, columns: n, rows: l };
                },
                getAreaCards: (t, n) => {
                    const { cardsMap: l } = e.getState();
                    return l[Xc(t, n)];
                },
            };
        })(K, U);
        let Y, B, G, J, V, X, Q;
        const W = K.getReactive(),
            Z = W.dragItemId;
        f(e, Z, (e) => l(5, (r = e)));
        const ee = W.before;
        f(e, ee, (e) => l(23, (m = e)));
        const te = W.overAreaId;
        f(e, te, (e) => l(24, (h = e)));
        const ne = W.dropAreasCoords;
        f(e, ne, (e) => l(25, (g = e)));
        const le = W.dragItemsCoords;
        f(e, le, (e) => l(45, (i = e)));
        const se = W.selected;
        let ue, de, fe, $e, me;
        return (
            f(e, se, (e) => l(18, (c = e))),
            (e.$$set = (e) => {
                "columns" in e && l(37, (I = e.columns)),
                    "rows" in e && l(38, (A = e.rows)),
                    "cards" in e && l(39, (E = e.cards)),
                    "cardShape" in e && l(40, (T = e.cardShape)),
                    "editorShape" in e && l(41, (L = e.editorShape)),
                    "editorAutoSave" in e && l(0, (j = e.editorAutoSave)),
                    "cardTemplate" in e && l(1, (z = e.cardTemplate)),
                    "readonly" in e && l(42, (N = e.readonly)),
                    "columnKey" in e && l(43, (O = e.columnKey)),
                    "rowKey" in e && l(44, (F = e.rowKey));
            }),
            (e.$$.update = () => {
                if (14272 & e.$$.dirty[1]) {
                    K.init({ columnKey: O, rowKey: F, columns: I, rows: A, cards: E, cardsMap: {}, cardsMeta: null, cardShape: T, editorShape: L });
                    const e = K.getReactive();
                    l(6, (Y = e.columns)),
                        b(),
                        (b = p(Y, (e) => l(19, (a = e)))),
                        l(8, (G = e.rows)),
                        S(),
                        (S = p(G, (e) => l(21, (d = e)))),
                        l(7, (B = e.rowKey)),
                        M(),
                        (M = p(B, (e) => l(22, ($ = e)))),
                        l(11, (X = e.cardShape)),
                        w(),
                        (w = p(X, (e) => l(4, (s = e)))),
                        l(10, (V = e.cardsMap)),
                        _(),
                        (_ = p(V, (e) => l(26, (v = e)))),
                        l(12, (Q = e.cardsMeta)),
                        C(),
                        (C = p(Q, (e) => l(27, (y = e)))),
                        l(9, (J = e.areasMeta)),
                        k(),
                        (k = p(J, (e) => l(20, (u = e))));
                }
                (32 & e.$$.dirty[0]) | (16384 & e.$$.dirty[1]) && l(17, (o = i && i[r])),
                    (24 & e.$$.dirty[0]) | (2048 & e.$$.dirty[1]) &&
                    ("object" == typeof N
                        ? (l(3, (ue = null == N ? void 0 : N.edit)), l(13, (de = null == N ? void 0 : N.add)), l(14, (fe = null == N ? void 0 : N.select)), l(15, ($e = null == N ? void 0 : N.dnd)))
                        : l(3, (ue = l(13, (de = l(14, (fe = l(15, ($e = !0 !== N)))))))),
                        ue || (x(X, (s.menu = s.menu || {}), s), x(X, (s.menu.show = !1), s)));
            }),
            [
                j,
                z,
                H,
                ue,
                s,
                r,
                Y,
                B,
                G,
                J,
                V,
                X,
                Q,
                de,
                fe,
                $e,
                me,
                o,
                c,
                a,
                u,
                d,
                $,
                m,
                h,
                g,
                v,
                y,
                ({ detail: { action: e, data: t } }) => H.exec(e, t),
                Z,
                ee,
                te,
                ne,
                le,
                se,
                function (e, t) {
                    H.exec(e, t);
                },
                function (e, t) {
                    const { hotkey: n } = t;
                    switch (n) {
                        case "delete":
                            (null == c ? void 0 : c.length) &&
                                c.map((e) => {
                                    H.exec("delete-card", { id: e });
                                });
                    }
                },
                I,
                A,
                E,
                T,
                L,
                N,
                O,
                F,
                i,
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (me = e), l(16, me);
                    });
                },
            ]
        );
    }
    class Pa extends Ye {
        constructor(e) {
            super(), He(this, e, Ra, qa, a, { columns: 37, rows: 38, cards: 39, cardShape: 40, editorShape: 41, editorAutoSave: 0, cardTemplate: 1, readonly: 42, columnKey: 43, rowKey: 44, api: 2 }, null, [-1, -1]);
        }
        get api() {
            return this.$$.ctx[2];
        }
    }
    function Ka(e) {
        let t, n;
        const l = e[1].default,
            o = $(l, e, e[0], null);
        return {
            c() {
                (t = z("div")), o && o.c(), U(t, "class", "toolbar svelte-1hfhlkm");
            },
            m(e, l) {
                T(e, t, l), o && o.m(t, null), (n = !0);
            },
            p(e, [t]) {
                o && o.p && (!n || 1 & t) && g(o, l, e, e[0], n ? h(l, e[0], t, null) : v(e[0]), null);
            },
            i(e) {
                n || (Ae(o, e), (n = !0));
            },
            o(e) {
                Ee(o, e), (n = !1);
            },
            d(e) {
                e && L(t), o && o.d(e);
            },
        };
    }
    function Ua(e, t, n) {
        let { $$slots: l = {}, $$scope: o } = t;
        var c;
        const s = ae("wx-i18n");
        return (
            s ? (null === (c = null == s ? void 0 : s.data) || void 0 === c ? void 0 : c.kanban) || s.extend(Os) : ie("wx-i18n", Ns(Os)),
            (e.$$set = (e) => {
                "$$scope" in e && n(0, (o = e.$$scope));
            }),
            [o, l]
        );
    }
    class Ha extends Ye {
        constructor(e) {
            super(), He(this, e, Ua, Ka, a, {});
        }
    }
    function Ya(e) {
        let t, l, o;
        return (
            (l = new Ys({ props: { name: "close", clickable: !0, click: e[15] } })),
            {
                c() {
                    (t = z("div")), Pe(l.$$.fragment), U(t, "class", "close-icon svelte-1o6cy5r");
                },
                m(e, n) {
                    T(e, t, n), Ke(l, t, null), (o = !0);
                },
                p: n,
                i(e) {
                    o || (Ae(l.$$.fragment, e), (o = !0));
                },
                o(e) {
                    Ee(l.$$.fragment, e), (o = !1);
                },
                d(e) {
                    e && L(t), Ue(l);
                },
            }
        );
    }
    function Ba(e) {
        let t, n;
        return (
            (t = new e[8]({ props: { cancel: e[11], width: "auto", $$slots: { default: [Qa] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    16842768 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Ga(e) {
        let t, n;
        const l = e[18].default,
            o = $(l, e, e[24], null);
        return {
            c() {
                (t = z("div")), o && o.c(), U(t, "class", "settings svelte-1o6cy5r");
            },
            m(e, l) {
                T(e, t, l), o && o.m(t, null), (n = !0);
            },
            p(e, t) {
                o && o.p && (!n || 16777216 & t) && g(o, l, e, e[24], n ? h(l, e[24], t, null) : v(e[24]), null);
            },
            i(e) {
                n || (Ae(o, e), (n = !0));
            },
            o(e) {
                Ee(o, e), (n = !1);
            },
            d(e) {
                e && L(t), o && o.d(e);
            },
        };
    }
    function Ja(e) {
        let t;
        return {
            c() {
                (t = z("div")), (t.textContent = `${e[10]("No results")}`), U(t, "class", "list-item no-results svelte-1o6cy5r");
            },
            m(e, n) {
                T(e, t, n);
            },
            p: n,
            i: n,
            o: n,
            d(e) {
                e && L(t);
            },
        };
    }
    function Va(e) {
        let t, n, l;
        return (
            (n = new e[9]({ props: { data: e[4], $$slots: { default: [Xa, ({ obj: e }) => ({ 26: e }), ({ obj: e }) => (e ? 67108864 : 0)] }, $$scope: { ctx: e } } })),
            {
                c() {
                    (t = z("div")), Pe(n.$$.fragment), U(t, "class", "results svelte-1o6cy5r");
                },
                m(e, o) {
                    T(e, t, o), Ke(n, t, null), (l = !0);
                },
                p(e, t) {
                    const l = {};
                    16 & t && (l.data = e[4]), 83886080 & t && (l.$$scope = { dirty: t, ctx: e }), n.$set(l);
                },
                i(e) {
                    l || (Ae(n.$$.fragment, e), (l = !0));
                },
                o(e) {
                    Ee(n.$$.fragment, e), (l = !1);
                },
                d(e) {
                    e && L(t), Ue(n);
                },
            }
        );
    }
    function Xa(e) {
        let t,
            n,
            l,
            o,
            c,
            s = e[26].label + "";
        function r() {
            return e[22](e[26]);
        }
        return {
            c() {
                (t = z("div")), (n = z("span")), (l = O(s)), U(n, "class", "list-item-text svelte-1o6cy5r"), U(t, "class", "list-item svelte-1o6cy5r");
            },
            m(e, s) {
                T(e, t, s), I(t, n), I(n, l), o || ((c = R(t, "click", r)), (o = !0));
            },
            p(t, n) {
                (e = t), 67108864 & n && s !== (s = e[26].label + "") && Y(l, s);
            },
            d(e) {
                e && L(t), (o = !1), c();
            },
        };
    }
    function Qa(e) {
        let t,
            n,
            l,
            o,
            c,
            s = e[16] ?.default && Ga(e);
        const r = [Va, Ja],
            i = [];
        function a(e, t) {
            return e[4] ? 0 : 1;
        }
        return (
            (l = a(e)),
            (o = i[l] = r[l](e)),
            {
                c() {
                    (t = z("div")), s && s.c(), (n = F()), o.c(), U(t, "class", "search-popup svelte-1o6cy5r");
                },
                m(e, o) {
                    T(e, t, o), s && s.m(t, null), I(t, n), i[l].m(t, null), (c = !0);
                },
                p(e, c) {
                    e[16] ?.default
                        ? s
                            ? (s.p(e, c), 65536 & c && Ae(s, 1))
                            : ((s = Ga(e)), s.c(), Ae(s, 1), s.m(t, n))
                        : s &&
                        (De(),
                            Ee(s, 1, 1, () => {
                                s = null;
                            }),
                            Ie());
                    let u = l;
                    (l = a(e)),
                        l === u
                            ? i[l].p(e, c)
                            : (De(),
                                Ee(i[u], 1, 1, () => {
                                    i[u] = null;
                                }),
                                Ie(),
                                (o = i[l]),
                                o ? o.p(e, c) : ((o = i[l] = r[l](e)), o.c()),
                                Ae(o, 1),
                                o.m(t, null));
                },
                i(e) {
                    c || (Ae(s), Ae(o), (c = !0));
                },
                o(e) {
                    Ee(s), Ee(o), (c = !1);
                },
                d(e) {
                    e && L(t), s && s.d(), i[l].d();
                },
            }
        );
    }
    function Wa(e) {
        let t, n, l, o, c, s, i, a, u, d, p;
        l = new Ys({ props: { name: "search" } });
        let f = !!e[0] && Ya(e),
            $ = e[5] && Ba(e);
        return {
            c() {
                (t = z("div")),
                    (n = z("div")),
                    Pe(l.$$.fragment),
                    (o = F()),
                    (c = z("input")),
                    (i = F()),
                    f && f.c(),
                    (a = F()),
                    $ && $.c(),
                    U(n, "class", "search-icon svelte-1o6cy5r"),
                    U(c, "id", (s = `${e[1]}`)),
                    (c.readOnly = e[2]),
                    U(c, "placeholder", e[3]),
                    U(c, "class", "svelte-1o6cy5r"),
                    U(t, "class", "search svelte-1o6cy5r"),
                    U(t, "tabindex", 1);
            },
            m(s, r) {
                T(s, t, r),
                    I(t, n),
                    Ke(l, n, null),
                    I(t, o),
                    I(t, c),
                    B(c, e[0]),
                    e[21](c),
                    I(t, i),
                    f && f.m(t, null),
                    I(t, a),
                    $ && $.m(t, null),
                    e[23](t),
                    (u = !0),
                    d || ((p = [R(c, "input", e[20]), R(c, "focus", e[12]), R(c, "blur", e[13]), R(t, "click", K(e[19]))]), (d = !0));
            },
            p(e, [n]) {
                (!u || (2 & n && s !== (s = `${e[1]}`))) && U(c, "id", s),
                    (!u || 4 & n) && (c.readOnly = e[2]),
                    (!u || 8 & n) && U(c, "placeholder", e[3]),
                    1 & n && c.value !== e[0] && B(c, e[0]),
                    e[0]
                        ? f
                            ? (f.p(e, n), 1 & n && Ae(f, 1))
                            : ((f = Ya(e)), f.c(), Ae(f, 1), f.m(t, a))
                        : f &&
                        (De(),
                            Ee(f, 1, 1, () => {
                                f = null;
                            }),
                            Ie()),
                    e[5]
                        ? $
                            ? ($.p(e, n), 32 & n && Ae($, 1))
                            : (($ = Ba(e)), $.c(), Ae($, 1), $.m(t, null))
                        : $ &&
                        (De(),
                            Ee($, 1, 1, () => {
                                $ = null;
                            }),
                            Ie());
            },
            i(e) {
                u || (Ae(l.$$.fragment, e), Ae(f), Ae($), (u = !0));
            },
            o(e) {
                Ee(l.$$.fragment, e), Ee(f), Ee($), (u = !1);
            },
            d(n) {
                n && L(t), Ue(l), e[21](null), f && f.d(), $ && $.d(), e[23](null), (d = !1), r(p);
            },
        };
    }
    function Za(e, t, n) {
        let { $$slots: l = {}, $$scope: o } = t;
        const c = (function (e) {
            const t = {};
            for (const n in e) t[n] = !0;
            return t;
        })(l),
            { Dropdown: s, List: r } = Oc,
            i = ae("wx-i18n").getGroup("kanban");
        let { value: a = "" } = t,
            { id: u = Gc() } = t,
            { readonly: d = !1 } = t,
            { focus: p = !1 } = t,
            { placeholder: f = i("Search") } = t,
            { searchResults: $ = null } = t;
        const m = re();
        let h,
            g,
            v = !1;
        function y(e) {
            m("action", { action: "result-click", id: e }), n(5, (v = !1));
        }
        p && ce(() => h.focus());
        return (
            (e.$$set = (e) => {
                "value" in e && n(0, (a = e.value)),
                    "id" in e && n(1, (u = e.id)),
                    "readonly" in e && n(2, (d = e.readonly)),
                    "focus" in e && n(17, (p = e.focus)),
                    "placeholder" in e && n(3, (f = e.placeholder)),
                    "searchResults" in e && n(4, ($ = e.searchResults)),
                    "$$scope" in e && n(24, (o = e.$$scope));
            }),
            [
                a,
                u,
                d,
                f,
                $,
                v,
                h,
                g,
                s,
                r,
                i,
                function (e) {
                    g.contains(e.target) || (n(5, (v = !1)), n(0, (a = "")));
                },
                function () {
                    n(5, (v = !0)), m("action", { action: "search-focus" });
                },
                function () {
                    m("action", { action: "search-blur" });
                },
                y,
                function () {
                    n(0, (a = ""));
                },
                c,
                p,
                l,
                function (t) {
                    ue.call(this, e, t);
                },
                function () {
                    (a = this.value), n(0, a);
                },
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (h = e), n(6, h);
                    });
                },
                (e) => y(e.id),
                function (e) {
                    pe[e ? "unshift" : "push"](() => {
                        (g = e), n(7, g);
                    });
                },
                o,
            ]
        );
    }
    class eu extends Ye {
        constructor(e) {
            super(), He(this, e, Za, Wa, a, { value: 0, id: 1, readonly: 2, focus: 17, placeholder: 3, searchResults: 4 });
        }
    }
    function tu(e) {
        let t, n, l, o, c, s;
        function r(t) {
            e[12](t);
        }
        let i = { options: e[7] };
        return (
            void 0 !== e[2].by && (i.value = e[2].by),
            (o = new e[4]({ props: i })),
            pe.push(() => Re(o, "value", r)),
            {
                c() {
                    (t = z("div")), (n = z("div")), (n.textContent = `${e[5]("Search in")}:`), (l = F()), Pe(o.$$.fragment), U(n, "class", "title svelte-85g0vm"), U(t, "class", "select svelte-85g0vm");
                },
                m(e, c) {
                    T(e, t, c), I(t, n), I(t, l), Ke(o, t, null), (s = !0);
                },
                p(e, t) {
                    const n = {};
                    !c && 4 & t && ((c = !0), (n.value = e[2].by), ye(() => (c = !1))), o.$set(n);
                },
                i(e) {
                    s || (Ae(o.$$.fragment, e), (s = !0));
                },
                o(e) {
                    Ee(o.$$.fragment, e), (s = !1);
                },
                d(e) {
                    e && L(t), Ue(o);
                },
            }
        );
    }
    function nu(e) {
        let t,
            n,
            l = e[0] && tu(e);
        return {
            c() {
                l && l.c(), (t = q());
            },
            m(e, o) {
                l && l.m(e, o), T(e, t, o), (n = !0);
            },
            p(e, n) {
                e[0]
                    ? l
                        ? (l.p(e, n), 1 & n && Ae(l, 1))
                        : ((l = tu(e)), l.c(), Ae(l, 1), l.m(t.parentNode, t))
                    : l &&
                    (De(),
                        Ee(l, 1, 1, () => {
                            l = null;
                        }),
                        Ie());
            },
            i(e) {
                n || (Ae(l), (n = !0));
            },
            o(e) {
                Ee(l), (n = !1);
            },
            d(e) {
                l && l.d(e), e && L(t);
            },
        };
    }
    function lu(e) {
        let t, n, l;
        function o(t) {
            e[13](t);
        }
        let c = { searchResults: e[1], $$slots: { default: [nu] }, $$scope: { ctx: e } };
        return (
            void 0 !== e[2].value && (c.value = e[2].value),
            (t = new eu({ props: c })),
            pe.push(() => Re(t, "value", o)),
            t.$on("action", e[8]),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, n) {
                    Ke(t, e, n), (l = !0);
                },
                p(e, [l]) {
                    const o = {};
                    2 & l && (o.searchResults = e[1]), 16389 & l && (o.$$scope = { dirty: l, ctx: e }), !n && 4 & l && ((n = !0), (o.value = e[2].value), ye(() => (n = !1))), t.$set(o);
                },
                i(e) {
                    l || (Ae(t.$$.fragment, e), (l = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (l = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function ou(e, t, l) {
        let o,
            c,
            s,
            r = n;
        e.$$.on_destroy.push(() => r());
        const { Select: i } = Oc,
            a = ae("wx-i18n").getGroup("kanban");
        let { api: u } = t,
            { showOptions: d = !0 } = t,
            $ = null;
        const m = ii({ value: "", by: null }, ({ value: e, by: t }) => {
            u.exec("set-search", { value: e, by: t });
        });
        let h;
        f(e, m, (e) => l(2, (c = e)));
        const g = [
            { id: null, label: a("Everywhere") },
            { id: "label", label: a("Label") },
            { id: "description", label: a("Description") },
        ];
        return (
            (e.$$set = (e) => {
                "api" in e && l(9, (u = e.api)), "showOptions" in e && l(0, (d = e.showOptions));
            }),
            (e.$$.update = () => {
                512 & e.$$.dirty && (l(3, (o = u && u.getReactiveState().cardsMeta)), r(), (r = p(o, (e) => l(11, (s = e))))),
                    2562 & e.$$.dirty && s && (l(1, ($ = Object.keys(s).reduce((e, t) => (s[t].found && e.push(null == u ? void 0 : u.getCard(t)), e), []))), $.length || l(1, ($ = null))),
                    1540 & e.$$.dirty &&
                    u &&
                    !h &&
                    (l(
                        10,
                        (h = (e) => {
                            ((null == e ? void 0 : e.value) === c.value && (null == e ? void 0 : e.by) === (null == c ? void 0 : c.by)) || m.reset(e);
                        })
                    ),
                        u.on("set-search", h));
            }),
            [
                d,
                $,
                c,
                o,
                i,
                a,
                m,
                g,
                function ({ detail: e }) {
                    const { id: t, action: n } = e;
                    switch (n) {
                        case "result-click":
                            u.exec("select-card", { id: t });
                            break;
                        case "search-focus":
                            c.value && u.exec("set-search", { value: c.value, by: c.by });
                    }
                },
                u,
                h,
                s,
                function (t) {
                    e.$$.not_equal(c.by, t) && ((c.by = t), m.set(c));
                },
                function (t) {
                    e.$$.not_equal(c.value, t) && ((c.value = t), m.set(c));
                },
            ]
        );
    }
    class cu extends Ye {
        constructor(e) {
            super(), He(this, e, ou, lu, a, { api: 9, showOptions: 0 });
        }
    }
    function su(e) {
        let t, l, o, c, s, r;
        return (
            (l = new Ys({ props: { name: "table-row-plus-after" } })),
            {
                c() {
                    (t = z("div")), Pe(l.$$.fragment), U(t, "class", "control svelte-14z0x6o"), U(t, "title", (o = e[2]("Add new row")));
                },
                m(n, o) {
                    T(n, t, o), Ke(l, t, null), (c = !0), s || ((r = R(t, "click", e[4])), (s = !0));
                },
                p: n,
                i(e) {
                    c || (Ae(l.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(l.$$.fragment, e), (c = !1);
                },
                d(e) {
                    e && L(t), Ue(l), (s = !1), r();
                },
            }
        );
    }
    function ru(e) {
        let t,
            n,
            l,
            o,
            c,
            s,
            r,
            i = e[1] && su(e);
        return (
            (o = new Ys({ props: { name: "table-column-plus-after" } })),
            {
                c() {
                    (t = z("div")), i && i.c(), (n = F()), (l = z("div")), Pe(o.$$.fragment), U(l, "class", "control svelte-14z0x6o"), U(l, "title", e[2]("Add new column")), U(t, "class", "controls svelte-14z0x6o");
                },
                m(a, u) {
                    T(a, t, u), i && i.m(t, null), I(t, n), I(t, l), Ke(o, l, null), (c = !0), s || ((r = R(l, "click", e[3])), (s = !0));
                },
                p(e, [l]) {
                    e[1]
                        ? i
                            ? (i.p(e, l), 2 & l && Ae(i, 1))
                            : ((i = su(e)), i.c(), Ae(i, 1), i.m(t, n))
                        : i &&
                        (De(),
                            Ee(i, 1, 1, () => {
                                i = null;
                            }),
                            Ie());
                },
                i(e) {
                    c || (Ae(i), Ae(o.$$.fragment, e), (c = !0));
                },
                o(e) {
                    Ee(i), Ee(o.$$.fragment, e), (c = !1);
                },
                d(e) {
                    e && L(t), i && i.d(), Ue(o), (s = !1), r();
                },
            }
        );
    }
    function iu(e, t, l) {
        let o,
            c = n;
        e.$$.on_destroy.push(() => c());
        let { api: s } = t;
        const r = ae("wx-i18n").getGroup("kanban");
        let i;
        return (
            (e.$$set = (e) => {
                "api" in e && l(5, (s = e.api));
            }),
            (e.$$.update = () => {
                32 & e.$$.dirty && s && (l(0, (i = s.getReactiveState().rowKey)), c(), (c = p(i, (e) => l(1, (o = e)))));
            }),
            [
                i,
                o,
                r,
                function () {
                    s.exec("add-column", { id: Gc(), column: { label: r("Untitled") } });
                },
                function () {
                    s.exec("add-row", { id: Gc(), row: { label: r("Untitled") } });
                },
                s,
            ]
        );
    }
    class au extends Ye {
        constructor(e) {
            super(), He(this, e, iu, ru, a, { api: 5 });
        }
    }
    function uu(e) {
        let t, n;
        return {
            c() {
                (t = new Q()), (n = q()), (t.a = n);
            },
            m(l, o) {
                t.m(e[0], l, o), T(l, n, o);
            },
            p(e, n) {
                1 & n && t.p(e[0]);
            },
            d(e) {
                e && L(n), e && t.d();
            },
        };
    }
    function du(e) {
        let t,
            l = e[0] && uu(e);
        return {
            c() {
                l && l.c(), (t = q());
            },
            m(e, n) {
                l && l.m(e, n), T(e, t, n);
            },
            p(e, [n]) {
                e[0] ? (l ? l.p(e, n) : ((l = uu(e)), l.c(), l.m(t.parentNode, t))) : l && (l.d(1), (l = null));
            },
            i: n,
            o: n,
            d(e) {
                l && l.d(e), e && L(t);
            },
        };
    }
    function pu(e, t, n) {
        let l;
        const c = ["template"];
        let s = w(t, c),
            { template: r } = t;
        return (
            (e.$$set = (e) => {
                (t = o(o({}, t), y(e))), n(2, (s = w(t, c))), "template" in e && n(1, (r = e.template));
            }),
            (e.$$.update = () => {
                n(0, (l = "function" == typeof r ? r(Object.assign({}, s)) : r));
            }),
            [l, r]
        );
    }
    class fu extends Ye {
        constructor(e) {
            super(), He(this, e, pu, du, a, { template: 1 });
        }
    }
    var $u = (function () {
        function e(e) {
            this._api = e;
        }
        return (
            (e.prototype.on = function (e, t) {
                this._api.on(e, t);
            }),
            (e.prototype.exec = function (e, t) {
                this._api.exec(e, t);
            }),
            e
        );
    })(),
        mu = (function () {
            function e(e, t) {
                (this.container = "string" == typeof e ? document.querySelector(e) : e), (this.config = t), this._init();
            }
            return (
                (e.prototype.destructor = function () {
                    this._kanban.$destroy(), (this._kanban = this.api = this.events = null);
                }),
                (e.prototype.setConfig = function (e) {
                    var n = this.serialize();
                    (this.config = t(t(t({}, this.config), n), e)), this._init();
                }),
                (e.prototype.parse = function (e) {
                    var t = e.cards,
                        n = e.columns,
                        l = e.rows;
                    (t || n || l) && (t && (this.config.cards = t), n && (this.config.columns = n), l && (this.config.rows = l), this._kanban.$set(this._configToProps(this.config)));
                }),
                (e.prototype.serialize = function () {
                    var e = this.api.getState();
                    return { cards: e.cards, columns: e.columns, rows: e.rows };
                }),
                (e.prototype.getCard = function (e) {
                    return this.api.getCard(e);
                }),
                (e.prototype.getAreaCards = function (e, t) {
                    return this.api.getAreaCards(e, t);
                }),
                (e.prototype.getSelection = function () {
                    return this.api.getState().selected;
                }),
                (e.prototype.addCard = function (e) {
                    this.api.exec("add-card", e);
                }),
                (e.prototype.updateCard = function (e) {
                    this.api.exec("update-card", e);
                }),
                (e.prototype.deleteCard = function (e) {
                    this.api.exec("delete-card", e);
                }),
                (e.prototype.moveCard = function (e) {
                    this.api.exec("move-card", e);
                }),
                (e.prototype.addColumn = function (e) {
                    this.api.exec("add-column", e);
                }),
                (e.prototype.updateColumn = function (e) {
                    this.api.exec("update-column", e);
                }),
                (e.prototype.addRow = function (e) {
                    this.api.exec("add-row", e);
                }),
                (e.prototype.updateRow = function (e) {
                    this.api.exec("update-row", e);
                }),
                (e.prototype.moveColumn = function (e) {
                    this.api.exec("move-column", e);
                }),
                (e.prototype.moveRow = function (e) {
                    this.api.exec("move-row", e);
                }),
                (e.prototype.deleteColumn = function (e) {
                    this.api.exec("delete-column", e);
                }),
                (e.prototype.deleteRow = function (e) {
                    this.api.exec("delete-row", e);
                }),
                (e.prototype.selectCard = function (e) {
                    this.api.exec("select-card", e);
                }),
                (e.prototype.unselectCard = function (e) {
                    this.api.exec("unselect-card", e);
                }),
                (e.prototype.setSearch = function (e) {
                    this.api.exec("set-search", e);
                }),
                (e.prototype.setLocale = function (e) {
                    e && this.setConfig({ locale: e });
                }),
                (e.prototype._init = function () {
                    this._kanban && this.destructor();
                    var e = new Map([
                        ["templates", { card: this.config.cardTemplate }],
                        ["wx-i18n", Ns(this.config.locale || Os)],
                    ]);
                    (this._kanban = new Pa({ target: this.container, props: this._configToProps(this.config), context: e })), (this.api = this._kanban.api), (this.events = new $u(this.api));
                }),
                (e.prototype._configToProps = function (e) {
                    return (null == e ? void 0 : e.cardTemplate)
                        ? t(t({}, e), {
                            cardTemplate:
                                ((n = fu),
                                    (l = null == e ? void 0 : e.cardTemplate),
                                    new Proxy(n, {
                                        construct: function (e, t) {
                                            var n = t[0].props || {};
                                            return (
                                                (n.template = l),
                                                (t[0].props = n),
                                                new (e.bind.apply(
                                                    e,
                                                    (function (e, t, n) {
                                                        if (n || 2 === arguments.length) for (var l, o = 0, c = t.length; o < c; o++) (!l && o in t) || (l || (l = Array.prototype.slice.call(t, 0, o)), (l[o] = t[o]));
                                                        return e.concat(l || Array.prototype.slice.call(t));
                                                    })([void 0], t)
                                                ))()
                                            );
                                        },
                                    })),
                        })
                        : e;
                    var n, l;
                }),
                e
            );
        })();
    function hu(e, t, n) {
        const l = e.slice();
        return (l[2] = t[n]), l;
    }
    function gu(e) {
        let t, n;
        return (
            (t = new fu({ props: { template: e[2] } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    2 & n && (l.template = e[2]), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function vu(e) {
        let t, n;
        return (
            (t = new au({ props: { api: e[0] } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    1 & n && (l.api = e[0]), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function yu(e) {
        let t, n;
        return (
            (t = new cu({ props: { api: e[0] } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    1 & n && (l.api = e[0]), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function wu(e) {
        let t, n, l, o;
        const c = [yu, vu, gu],
            s = [];
        function r(e, t) {
            return "search" === e[2] ? 0 : "controls" === e[2] ? 1 : e[2] ? 2 : -1;
        }
        return (
            ~(t = r(e)) && (n = s[t] = c[t](e)),
            {
                c() {
                    n && n.c(), (l = q());
                },
                m(e, n) {
                    ~t && s[t].m(e, n), T(e, l, n), (o = !0);
                },
                p(e, o) {
                    let i = t;
                    (t = r(e)),
                        t === i
                            ? ~t && s[t].p(e, o)
                            : (n &&
                                (De(),
                                    Ee(s[i], 1, 1, () => {
                                        s[i] = null;
                                    }),
                                    Ie()),
                                ~t ? ((n = s[t]), n ? n.p(e, o) : ((n = s[t] = c[t](e)), n.c()), Ae(n, 1), n.m(l.parentNode, l)) : (n = null));
                },
                i(e) {
                    o || (Ae(n), (o = !0));
                },
                o(e) {
                    Ee(n), (o = !1);
                },
                d(e) {
                    ~t && s[t].d(e), e && L(l);
                },
            }
        );
    }
    function bu(e) {
        let t,
            n,
            l = e[1],
            o = [];
        for (let t = 0; t < l.length; t += 1) o[t] = wu(hu(e, l, t));
        const c = (e) =>
            Ee(o[e], 1, 1, () => {
                o[e] = null;
            });
        return {
            c() {
                for (let e = 0; e < o.length; e += 1) o[e].c();
                t = q();
            },
            m(e, l) {
                for (let t = 0; t < o.length; t += 1) o[t].m(e, l);
                T(e, t, l), (n = !0);
            },
            p(e, n) {
                if (3 & n) {
                    let s;
                    for (l = e[1], s = 0; s < l.length; s += 1) {
                        const c = hu(e, l, s);
                        o[s] ? (o[s].p(c, n), Ae(o[s], 1)) : ((o[s] = wu(c)), o[s].c(), Ae(o[s], 1), o[s].m(t.parentNode, t));
                    }
                    for (De(), s = l.length; s < o.length; s += 1) c(s);
                    Ie();
                }
            },
            i(e) {
                if (!n) {
                    for (let e = 0; e < l.length; e += 1) Ae(o[e]);
                    n = !0;
                }
            },
            o(e) {
                o = o.filter(Boolean);
                for (let e = 0; e < o.length; e += 1) Ee(o[e]);
                n = !1;
            },
            d(e) {
                j(o, e), e && L(t);
            },
        };
    }
    function xu(e) {
        let t, n;
        return (
            (t = new Ha({ props: { $$slots: { default: [bu] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, n) {
                    const l = {};
                    35 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function ku(e) {
        let t, n;
        return (
            (t = new Kc({ props: { $$slots: { default: [xu] }, $$scope: { ctx: e } } })),
            {
                c() {
                    Pe(t.$$.fragment);
                },
                m(e, l) {
                    Ke(t, e, l), (n = !0);
                },
                p(e, [n]) {
                    const l = {};
                    35 & n && (l.$$scope = { dirty: n, ctx: e }), t.$set(l);
                },
                i(e) {
                    n || (Ae(t.$$.fragment, e), (n = !0));
                },
                o(e) {
                    Ee(t.$$.fragment, e), (n = !1);
                },
                d(e) {
                    Ue(t, e);
                },
            }
        );
    }
    function Su(e, t, n) {
        let { api: l } = t,
            { items: o = ["search", "controls"] } = t;
        return (
            (e.$$set = (e) => {
                "api" in e && n(0, (l = e.api)), "items" in e && n(1, (o = e.items));
            }),
            [l, o]
        );
    }
    class Mu extends Ye {
        constructor(e) {
            super(), He(this, e, Su, ku, a, { api: 0, items: 1 });
        }
    }
    var _u = (function () {
        function e(e, t) {
            (this.container = "string" == typeof e ? document.querySelector(e) : e), (this.config = t), this._init();
        }
        return (
            (e.prototype.destructor = function () {
                this._toolbar.$destroy(), (this._toolbar = this.api = this.events = null);
            }),
            (e.prototype.setConfig = function (e) {
                e && ((this.config = t(t({}, this.config), e)), this._init());
            }),
            (e.prototype.setLocale = function (e) {
                e && this.setConfig({ locale: e });
            }),
            (e.prototype._init = function () {
                var e;
                this._toolbar && this.destructor();
                var t = new Map([["wx-i18n", Ns((null === (e = this.config) || void 0 === e ? void 0 : e.locale) || Os)]]);
                (this._toolbar = new Mu({ target: this.container, props: this._configToProps(this.config), context: t })), (this.events = new $u(this.api));
            }),
            (e.prototype._configToProps = function (e) {
                return t({}, e);
            }),
            e
        );
    })();
    const Cu = Symbol();
    class Du {
        constructor() {
            (this._awaitAddingQueue = []), (this._queue = {}), (this._idPool = {}), (this.add = this.add.bind(this));
        }
        add(e, t, n) {
            if (n.debounce) {
                const l = `${e}"/"${t.id}`,
                    o = this._queue[l];
                return (
                    o && clearTimeout(o),
                    void (this._queue[l] = setTimeout(() => {
                        this.add(e, t, { ...n, debounce: !1 });
                    }, n.debounce))
                );
            }
            const l = this.tryExec(e, t, n);
            null === l
                ? this._awaitAddingQueue.push({ action: e, data: t, proc: n })
                : l.then((e) => {
                    e && e.id && e.id != t.id && this.isTempID(t.id) && ((this._idPool[t.id] = e.id), this.execQueue());
                });
        }
        tryExec(e, t, n) {
            const l = this.exec(e, t, n);
            return null === l && this._awaitAddingQueue.push({ action: e, data: t, proc: n }), l;
        }
        exec(e, t, n) {
            const l = this.correctID(t, n.ignoreID ? t.id : null);
            return l === Cu ? null : n.handler(l, e, t);
        }
        isTempID(e) {
            return "string" == typeof e && 20 === e.length && parseInt(e.substr(7)) > 1e12;
        }
        correctID(e, t) {
            let n = null;
            for (const l in e) {
                const o = e[l];
                if ("object" == typeof o) {
                    const c = this.correctID(o, t);
                    if (c !== o) {
                        if (c === Cu) return Cu;
                        null === n && (n = { ...e }), (n[l] = c);
                    }
                } else if (o !== t && this.isTempID(o)) {
                    const t = this._idPool[o];
                    if (!t) return Cu;
                    null === n && (n = { ...e }), (n[l] = t);
                }
            }
            return n || e;
        }
        execQueue() {
            this._awaitAddingQueue = this._awaitAddingQueue.map((e) => (this.tryExec(e.action, e.data, e.proc) ? null : e)).filter((e) => null !== e);
        }
    }
    return (
        (e.Kanban = mu),
        (e.RestDataProvider = class extends Uc {
            constructor(e) {
                super(), (this._url = e), (this._queue = new Du());
                const t = {
                    "add-card": { ignoreID: !0, handler: (e) => this.send("cards", "POST", e) },
                    "update-card": { debounce: 500, handler: (e) => this.send(`cards/${e.id}`, "PUT", e) },
                    "move-card": { handler: (e) => this.send(`cards/${e.id}/move`, "PUT", e) },
                    "delete-card": { handler: (e) => this.send(`cards/${e.id}`, "DELETE") },
                    "add-column": { ignoreID: !0, handler: (e) => this.send("columns", "POST", e) },
                    "update-column": { debounce: 500, handler: (e) => this.send(`columns/${e.id}`, "PUT", e) },
                    "delete-column": { handler: (e) => this.send(`columns/${e.id}`, "DELETE", e) },
                    "add-row": { ignoreID: !0, handler: (e) => this.send("rows", "POST", e) },
                    "update-row": { debounce: 500, handler: (e) => this.send(`rows/${e.id}`, "PUT", e) },
                    "delete-row": { handler: (e) => this.send(`rows/${e.id}`, "DELETE", e) },
                },
                    n = this.getHandlers(t);
                for (const e in n) this.on(e, (t) => this._queue.add(e, t, n[e]));
            }
            getCards() {
                return this.send("cards", "GET").then(this.parseCards);
            }
            getColumns() {
                return this.send("columns", "GET");
            }
            getRows() {
                return this.send("rows", "GET");
            }
            getHandlers(e) {
                return e;
            }
            send(e, t, n, l = {}) {
                const o = { method: t, headers: { "Content-Type": "application/json", ...l } };
                return n && (o.body = "object" == typeof n ? JSON.stringify(n) : n), fetch(`${this._url}/${e}`, o).then((e) => e.json());
            }
            parseCards(e) {
                return e.forEach((e) => (e.end_date && (e.end_date = new Date(e.end_date)), e.start_date && (e.start_date = new Date(e.start_date)), e)), e;
            }
        }),
        (e.Toolbar = _u),
        (e.cn = qs),
        (e.defaultCardShape = Wc),
        (e.defaultEditorShape = Zc),
        (e.en = Os),
        (e.ru = Rs),
        (e.uid = Gc),
        Object.defineProperty(e, "__esModule", { value: !0 }),
        e
    );
})({});
//# sourceMappingURL=kanban.js.map
