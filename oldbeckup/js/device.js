// js/device.js
// Runtime environment detection for adaptive UX/UI.
// - Estimates screen PPI (approximate) and diagonal size.
// - Detects device type (phone/tablet/desktop) heuristically.
// - Publishes results to window.__deviceEnv and CSS variables on :root.
//
// Notes:
// 1) True physical PPI cannot be reliably known in browsers on all platforms.
//    We use the standard CSS inch (96 CSS px) + devicePixelRatio as an estimate.
// 2) Optional user calibration can override estimate (stored in localStorage).

const LS_KEY_PPI = "dilovodstvo_ppi_calibrated_v1";

function clamp(n, a, b){ return Math.max(a, Math.min(b, n)); }

export function isTouchDevice(){
  return ("ontouchstart" in window) || (navigator.maxTouchPoints||0) > 0;
}

export function getCalibratedPPI(){
  const v = localStorage.getItem(LS_KEY_PPI);
  if(!v) return null;
  const n = Number(v);
  if(!Number.isFinite(n) || n <= 0) return null;
  return n;
}

export function setCalibratedPPI(ppi){
  const n = Number(ppi);
  if(!Number.isFinite(n) || n <= 0) throw new Error("Invalid PPI");
  localStorage.setItem(LS_KEY_PPI, String(Math.round(n)));
}

export function clearCalibratedPPI(){
  localStorage.removeItem(LS_KEY_PPI);
}

export function estimatePPI(){
  const calibrated = getCalibratedPPI();
  if(calibrated) return { ppi: calibrated, method: "calibrated" };

  // Standard approximation: 96 CSS DPI * devicePixelRatio.
  const dpr = Number(window.devicePixelRatio || 1);
  const ppi = 96 * (Number.isFinite(dpr) ? dpr : 1);
  return { ppi, method: "css-dpi" };
}

function estimateDiagonalInches(ppi){
  try{
    const dpr = Number(window.devicePixelRatio || 1);
    const wpx = (screen?.width || window.innerWidth || 0) * dpr;
    const hpx = (screen?.height || window.innerHeight || 0) * dpr;
    if(!wpx || !hpx || !ppi) return null;
    const diagPx = Math.sqrt(wpx*wpx + hpx*hpx);
    return diagPx / ppi;
  }catch(_){
    return null;
  }
}

function detectDeviceType({touch, diagIn, vw}){
  // Primary: diagonal size if available.
  if(diagIn && Number.isFinite(diagIn)){
    if(touch && diagIn < 7.0) return "phone";
    if(touch && diagIn < 13.0) return "tablet";
    // Large touch devices are treated as desktop/tablet-like.
    return "desktop";
  }

  // Fallback: viewport width heuristics.
  if(touch){
    if(vw < 600) return "phone";
    if(vw < 1024) return "tablet";
    return "desktop";
  }
  return "desktop";
}

export function computeDeviceEnv(){
  const touch = isTouchDevice();
  const dpr = Number(window.devicePixelRatio || 1);
  const { ppi, method } = estimatePPI();
  const diagIn = estimateDiagonalInches(ppi);
  const vw = window.innerWidth || 0;
  const vh = window.innerHeight || 0;

  const deviceType = detectDeviceType({ touch, diagIn, vw });

  // A conservative UI scale hint for future tuning.
  // (Not applied automatically to layout unless you use the CSS var.)
  let uiScale = 1;
  if(deviceType === "phone") uiScale = 1.15;
  else if(deviceType === "tablet") uiScale = 1.08;
  uiScale = clamp(uiScale, 0.9, 1.3);

  return {
    ppi,
    ppiMethod: method,
    dpr,
    touch,
    diagIn: diagIn ? Math.round(diagIn * 10) / 10 : null,
    viewport: { w: vw, h: vh },
    deviceType,
    uiScale,
    userAgent: navigator.userAgent,
  };
}

export function applyDeviceEnvToDOM(env){
  const root = document.documentElement;

  // Data attributes for CSS hooks.
  root.dataset.deviceType = env.deviceType;
  root.dataset.touch = env.touch ? "1" : "0";

  // CSS variables for future sizing logic.
  root.style.setProperty("--device-ppi", String(Math.round(env.ppi)));
  root.style.setProperty("--device-dpr", String(env.dpr));
  root.style.setProperty("--device-diag-in", env.diagIn ? String(env.diagIn) : "");
  root.style.setProperty("--device-ui-scale", String(env.uiScale));
}

let _env = null;
let _installed = false;

export function getDeviceEnv(){
  return _env || computeDeviceEnv();
}

export function initDeviceEnv(){
  _env = computeDeviceEnv();
  applyDeviceEnvToDOM(_env);
  // publish for non-module usage/debug
  window.__deviceEnv = _env;

  if(_installed) return _env;
  _installed = true;

  // Keep env fresh on resize/orientation change.
  let t = null;
  const onChange = ()=>{
    clearTimeout(t);
    t = setTimeout(()=>{
      _env = computeDeviceEnv();
      applyDeviceEnvToDOM(_env);
      window.__deviceEnv = _env;
      window.dispatchEvent(new CustomEvent("deviceenv:change", { detail: _env }));
    }, 120);
  };
  window.addEventListener("resize", onChange);
  window.addEventListener("orientationchange", onChange);

  return _env;
}
