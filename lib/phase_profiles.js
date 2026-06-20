// ================================================================
// Bellweather — Phase Load Profiles  (v1, derived June 2026)
// Derived from 3-year design+dev timecard aggregate via image analysis
// ================================================================
//
// Each function returns relative intensity f(t) for t ∈ [0,1]
// where t=0 = phase start week, t=1 = phase end week.
// Use distributeHours() to turn a total-hours estimate into
// a week-by-week array.

// Phase I: Gaussian bell on floor
// Peak at t≈0.65 (week ~65% through, e.g. week 4 of 6)
// Floor of 0.68 — Phase I starts immediately at moderate intensity
function phaseIProfile(t) {
  const floor = 0.678, mu = 0.647, sigma = 0.311;
  return floor + (1 - floor) * Math.exp(-Math.pow((t - mu) / sigma, 2));
}

// Phase II: Bimodal — early spike then sustained plateau
// Spike at t≈0.07 (first 1–2 wks), main effort peaks at t≈0.59
// Floor of 0.39; dips between the two peaks (~t=0.13–0.20)
function phaseIIProfile(t) {
  const floor = 0.394;
  const A1 = 0.411, mu1 = 0.067, s1 = 0.101;  // early spike
  const A2 = 0.700, mu2 = 0.593, s2 = 0.434;  // main effort
  const spike = A1 * Math.exp(-Math.pow((t - mu1) / s1, 2));
  const main  = A2 * Math.exp(-Math.pow((t - mu2) / s2, 2));
  return Math.min(1, floor + spike + main);
}

// Phase III: Gamma-shaped decay + floor
// Peak at t≈0.02 (early burst), decays to sustained floor ~0.14
// Slight uptick at end not modeled (punch-list / closeout)
function phaseIIIProfile(t) {
  const A = 0.734, tau = 0.082, alpha = 0.300, floor = 0.140;
  const ts = Math.max(t, 1e-5);
  const shape = Math.pow(ts / tau, alpha) * Math.exp(-ts / tau);
  const norm  = Math.pow(alpha, alpha) * Math.exp(-alpha); // shape at t=alpha*tau
  return A * (shape / norm) + floor;
}

// Apply a profile to a total-hours estimate
// Returns an integer array of length numWeeks summing to totalHours
function distributeHours(totalHours, numWeeks, profileFn) {
  const weights = Array.from({length: numWeeks}, (_, i) =>
    Math.max(0, profileFn((i + 0.5) / numWeeks))
  );
  const sum = weights.reduce((a, b) => a + b, 0);
  const dist = weights.map(w => Math.round(w / sum * totalHours));
  dist[dist.length - 1] += totalHours - dist.reduce((a, b) => a + b, 0);
  return dist;
}
