// ================================================================
// Bellweather — Phase Load Profiles  (v2, calibrated June 2026)
// Derived from 3-year designer+developer timecard aggregate
// Y-axis calibrated against source chart gridlines (0–1400 hr scale)
// Fit quality: Phase I R²=0.937 · Phase II R²=0.918 · Phase III R²=0.970
// ================================================================
//
// Each function returns aggregate hours intensity at normalized phase
// position t ∈ [0,1] (t=0 = phase start, t=1 = phase end).
//
// Values are on the 0–1400 hr absolute scale of the source data.
// To distribute a project's total phase hours week-by-week, use
// distributeHours() — it normalizes the shape automatically.

const BELLWEATHER_PHASE_PROFILES = (() => {

  // Phase I — Initial Design (~6 weeks)
  // Gaussian bell on a floor.
  // Peaks at t≈0.64 (e.g. week 4 of 6). Floor=625 hrs — starts immediately.
  function phaseI(t) {
    const floor = 624.509, mu = 0.6436, sigma = 0.3461, amp = 300.681;
    return floor + amp * Math.exp(-Math.pow((t - mu) / sigma, 2));
  }

  // Phase II — Design Development (16–20 weeks)
  // Bimodal: early kickoff spike at t≈0.17, main sustained effort peaking at t≈0.62.
  // The dip between them (~t=0.13–0.20) is the client-review pause.
  function phaseII(t) {
    const floor = 461.162;
    const A1 = 152.986, mu1 = 0.1747, s1 = 0.2000;
    const A2 = 900.000, mu2 = 0.6157, s2 = 0.4963;
    return Math.min(1400,
      floor
      + A1 * Math.exp(-Math.pow((t - mu1) / s1, 2))
      + A2 * Math.exp(-Math.pow((t - mu2) / s2, 2))
    );
  }

  // Phase III — Construction Design (24–30 weeks)
  // Gamma-shaped early burst decaying to a sustained floor.
  // Peak at t≈0.048. Floor=142 hrs — persistent RFI/detail activity.
  function phaseIII(t) {
    const A = 543.288, tau = 0.1604, alpha = 0.3000, floor = 142.126;
    const ts  = Math.max(t, 1e-5);
    const shape = Math.pow(ts / tau, alpha) * Math.exp(-ts / tau);
    const norm  = Math.pow(alpha, alpha) * Math.exp(-alpha);
    return Math.max(0, A * (shape / norm) + floor);
  }

  // Distribute total phase hours across numWeeks using a profile shape.
  // Returns integer array of length numWeeks summing to totalHours.
  function distributeHours(totalHours, numWeeks, profileFn) {
    const weights = Array.from({ length: numWeeks }, (_, i) =>
      Math.max(0, profileFn((i + 0.5) / numWeeks))
    );
    const sum = weights.reduce((a, b) => a + b, 0);
    const dist = weights.map(w => Math.round(w / sum * totalHours));
    dist[dist.length - 1] += totalHours - dist.reduce((a, b) => a + b, 0);
    return dist;
  }

  // Cumulative burn — useful for "% expected complete at week N" comparisons
  function cumulativeHours(totalHours, numWeeks, profileFn) {
    const weekly = distributeHours(totalHours, numWeeks, profileFn);
    let cum = 0;
    return weekly.map(w => (cum += w));
  }

  // Expected % complete at a given phase position t (0–1)
  // Used in Hours Analyzer for actual-vs-predicted comparison
  function expectedPctComplete(t, profileFn, steps = 200) {
    const dt = 1 / steps;
    let total = 0, elapsed = 0;
    for (let i = 0; i < steps; i++) {
      const mid = (i + 0.5) * dt;
      const w = Math.max(0, profileFn(mid));
      total += w;
      if (mid <= t) elapsed += w;
    }
    return total > 0 ? elapsed / total : 0;
  }

  return { phaseI, phaseII, phaseIII, distributeHours, cumulativeHours, expectedPctComplete };
})();
