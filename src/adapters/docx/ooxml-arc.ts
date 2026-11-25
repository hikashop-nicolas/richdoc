// DrawingML's arcTo, shared by the docx outlines and odt's OOXML arcs (LibreOffice's G command).

type Pt = [number, number];
const DEG = 60000; // angles are in 1/60 000 degree
const rad = (a: number) => (a / DEG) * (Math.PI / 180);
const r = (n: number) => String(Math.round(n * 100) / 100);

/** DrawingML's arcTo: from the current point along the wR x hR ellipse, from angle stAng through
 *  swAng. Angles are visual (the direction from the centre), y down. */
export function arcTo(cur: Pt, wR: number, hR: number, st: number, sw: number): { d: string; to: Pt; firstCtl: Pt; lastCtl: Pt } {
  // The parametric angle of a point at visual angle `a` on the ellipse.
  const param = (a: number) => Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
  const a0 = rad(st), a1 = rad(st + sw);
  const t0 = param(a0);
  let t1 = param(a1);
  let sweep = t1 - t0;
  if (Math.abs(sw) >= 21600000) sweep = Math.sign(sw) * 2 * Math.PI;
  else if (sw > 0 && sweep < 0) sweep += 2 * Math.PI;
  else if (sw < 0 && sweep > 0) sweep -= 2 * Math.PI;
  t1 = t0 + sweep;
  const cx = cur[0] - wR * Math.cos(t0), cy = cur[1] - hR * Math.sin(t0);
  const at = (t: number): Pt => [cx + wR * Math.cos(t), cy + hR * Math.sin(t)];
  const n = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
  const step = sweep / n;
  const k = (4 / 3) * Math.tan(step / 4);
  let d = "";
  let firstCtl: Pt = cur, lastCtl: Pt = cur;
  for (let s = 0; s < n; s++) {
    const ta = t0 + s * step, tb = ta + step;
    const p0 = at(ta), p3 = at(tb);
    const c1: Pt = [p0[0] - k * wR * Math.sin(ta), p0[1] + k * hR * Math.cos(ta)];
    const c2: Pt = [p3[0] + k * wR * Math.sin(tb), p3[1] - k * hR * Math.cos(tb)];
    if (s === 0) firstCtl = c1;
    lastCtl = c2;
    d += `C${r(c1[0])},${r(c1[1])} ${r(c2[0])},${r(c2[1])} ${r(p3[0])},${r(p3[1])} `;
  }
  if (!wR || !hR || !sweep) return { d: "", to: cur, firstCtl: cur, lastCtl: cur };
  return { d, to: at(t1), firstCtl, lastCtl };
}
