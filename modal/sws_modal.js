(function(){
  const $overlay = () => document.getElementById("swsOverlay");
  const $window = () => document.getElementById("swsWindow");
  const $stack  = () => document.getElementById("swsStack");

  const isTextArea = (el)=> el && el.tagName === "TEXTAREA";

  const state = {
    inited: false,
    mode: "auto",
    theme: "light",
    tokens: null,
    stack: [],
    closeOnRootBack: true,
  };

  function applyTheme(theme){
    state.theme = theme === "dark" ? "dark" : "light";
    const ov = $overlay();
    if(!ov) return;
    ov.setAttribute("data-sws-theme", state.theme);
  }

  function applyTokens(tokens){
    state.tokens = tokens || null;
    const w = $window();
    if(!w) return;
    const style = w.style;

    const c = tokens?.colors || {};
    if(c.overlay)  style.setProperty("--sws-overlay", c.overlay);
    if(c.windowBg) style.setProperty("--sws-window-bg", c.windowBg);
    if(c.surface)  style.setProperty("--sws-surface", c.surface);
    if(c.text)     style.setProperty("--sws-text", c.text);
    if(c.muted)    style.setProperty("--sws-muted", c.muted);
    if(c.border)   style.setProperty("--sws-border", c.border);
    if(c.primary)  style.setProperty("--sws-primary", c.primary);
    if(c.primaryText) style.setProperty("--sws-primary-text", c.primaryText);

    const r = tokens?.radius || {};
    if(r.window!=null) style.setProperty("--sws-radius-window", r.window + "px");
    if(r.card!=null)   style.setProperty("--sws-radius-card", r.card + "px");
    if(r.button!=null) style.setProperty("--sws-radius-btn", r.button + "px");

    const f = tokens?.font || {};
    if(f.basePx!=null) style.setProperty("--sws-font-base", f.basePx + "px");
    if(f.titlePx!=null) style.setProperty("--sws-font-title", f.titlePx + "px");

    const s = tokens?.spacing || {};
    if(s.pad!=null) style.setProperty("--sws-pad", s.pad + "px");
  }

  function resolveMode(mode){
    if(mode && mode !== "auto") return mode;
    const ua = (navigator.userAgent || "").toLowerCase();
    const isWV = ua.includes(" wv") || ua.includes("; wv") || (ua.includes("android") && ua.includes("version/"));
    const isMobile = /android|iphone|ipad|ipod|mobile/.test(ua);
    if(isWV || isMobile) return "webview";
    return "desktop";
  }

  function openOverlay(){
    const ov = $overlay();
    if(!ov) return;
    ov.classList.add("sws-open");
    ov.setAttribute("aria-hidden","false");
    document.body.style.overflow = "hidden";
  }
  function closeOverlay(){
    const ov = $overlay();
    if(!ov) return;
    ov.classList.remove("sws-open");
    ov.setAttribute("aria-hidden","true");
    document.body.style.overflow = "";
    state.stack = [];
    const st = $stack();
    if(st) st.innerHTML = "";
  }

  function top(){ return state.stack[state.stack.length - 1] || null; }

  function setWindowClass(){
    const w = $window();
    if(!w) return;
    const m = resolveMode(state.mode);
    w.classList.toggle("sws-desktop", m === "desktop");
    w.classList.toggle("sws-webview", m !== "desktop");
  }

  function updateScreenClasses(){
    state.stack.forEach((s, i) => {
      const el = s.screenEl;
      if(!el) return;
      const isTop = i === state.stack.length - 1;
      el.classList.toggle("sws-active", isTop);
      el.classList.toggle("sws-inactive", !isTop);
      el.classList.toggle("sws-left", !isTop);
    });
  }

  function buildUIHelpers(){
    const ui = {};

    ui.el = (tag, cls, text) => {
      const e = document.createElement(tag);
      if(cls) e.className = cls;
      if(text!=null) e.textContent = text;
      return e;
    };

    ui.card = ({title, description, children}) => {
      const c = ui.el("div","sws-card");
      if(title){
        const h = ui.el("div","sws-card-title", title);
        c.appendChild(h);
      }
      if(description){
        const d = ui.el("div","sws-muted", description);
        d.style.marginBottom = "8px";
        c.appendChild(d);
      }
      if(children){
        if(Array.isArray(children)) children.forEach(ch => ch && c.appendChild(ch));
        else if(children instanceof HTMLElement) c.appendChild(children);
      }
      return c;
    };

    ui.list = (items) => {
      const wrap = ui.el("div","sws-list");
      (items||[]).forEach(it => {
        const row = ui.el("div","sws-item");
        const left = ui.el("div","sws-item-left");
        left.appendChild(ui.el("div","sws-item-label", it.label || ""));
        if(it.description) left.appendChild(ui.el("div","sws-item-desc", it.description));
        const che = ui.el("div","sws-chevron","›");
        row.appendChild(left);
        row.appendChild(che);
        row.onclick = () => { if(typeof it.onOpen === "function") it.onOpen(); };
        wrap.appendChild(row);
      });
      return wrap;
    };

    ui.controlRow = ({label, help, controlEl}) => {
      const row = ui.el("div","sws-control");
      const left = ui.el("div","");
      left.style.minWidth = "0";
      left.appendChild(ui.el("div","sws-control-label", label || ""));
      if(help) left.appendChild(ui.el("div","sws-control-help", help));
      row.appendChild(left);
      if(controlEl) row.appendChild(controlEl);
      return row;
    };

    ui.input = ({value="", placeholder="", type="text", onChange}) => {
      const i = document.createElement("input");
      i.className = "sws-input";
      i.type = type;
      i.value = value ?? "";
      i.placeholder = placeholder || "";
      i.addEventListener("input", ()=> { if(onChange) onChange(i.value); });
      return i;
    };

    ui.select = ({value="", options=[], onChange}) => {
      const s = document.createElement("select");
      s.className = "sws-select";
      (options||[]).forEach(opt => {
        const o = document.createElement("option");
        o.value = opt.value;
        o.textContent = opt.label;
        s.appendChild(o);
      });
      s.value = value ?? "";
      s.addEventListener("change", ()=> { if(onChange) onChange(s.value); });
      return s;
    };

    ui.toggle = ({value=false, onChange}) => {
      const t = ui.el("div","sws-toggle");
      t.setAttribute("role","switch");
      t.tabIndex = 0;
      t.dataset.on = value ? "1" : "0";
      const set = (v)=>{ t.dataset.on = v ? "1" : "0"; if(onChange) onChange(!!v); };
      t.onclick = ()=> set(t.dataset.on !== "1");
      t.addEventListener("keydown",(e)=>{
        if(e.key===" " || e.key==="Enter"){
          e.preventDefault();
          set(t.dataset.on !== "1");
        }
      });
      return t;
    };

    ui.push = (opts)=> SettingsWindow.push(opts);
    ui.pop = ()=> SettingsWindow.pop();
    ui.close = ()=> SettingsWindow.close();

    return ui;
  }

  function renderScreen({title, subtitle, content, onSave, saveLabel, canSave, ctx}){
    const screen = document.createElement("div");
    screen.className = "sws-screen";

    const header = document.createElement("div");
    header.className = "sws-header";

    const back = document.createElement("button");
    back.className = "sws-back";
    back.textContent = "←";
    back.title = "Назад (Esc)";
    back.onclick = () => SettingsWindow.pop();

    const headtxt = document.createElement("div");
    headtxt.className = "sws-headtxt";

    const t = document.createElement("div");
    t.className = "sws-title";
    t.textContent = title || "Налаштування";

    const sub = document.createElement("div");
    sub.className = "sws-subtitle";
    sub.textContent = subtitle || "";

    headtxt.appendChild(t);
    headtxt.appendChild(sub);

    header.appendChild(back);
    header.appendChild(headtxt);

    const body = document.createElement("div");
    body.className = "sws-body";

    const footer = document.createElement("div");
    footer.className = "sws-footer";

    const save = document.createElement("button");
    save.className = "sws-save";
    save.textContent = saveLabel || "Зберегти";
    save.title = "Зберегти (Enter)";
    save.onclick = () => {
      const s = top();
      if(!s) return;
      if(s.canSave && !s.canSave(s.ctx)) return;
      if(typeof s.onSave === "function") s.onSave(s.ctx);
    };

    footer.appendChild(save);
    screen.appendChild(header);
    screen.appendChild(body);
    screen.appendChild(footer);

    const ui = buildUIHelpers();
    const localCtx = Object.assign({}, ctx || {}, {
      ui,
      setSaveEnabled: (enabled)=> { save.disabled = !enabled; },
      setSubtitle: (txt)=> { sub.textContent = txt || ""; },
      setTitle: (txt)=> { t.textContent = txt || ""; },
      setDirty: (dirty)=> { sub.textContent = (subtitle || "") + (dirty ? " • змінено" : ""); },
    });

    if(typeof canSave === "function"){
      try{ save.disabled = !canSave(localCtx); }catch(_){}
    }

    if(typeof content === "function"){
      const out = content(localCtx);
      if(out instanceof HTMLElement) body.appendChild(out);
    } else if(content instanceof HTMLElement){
      body.appendChild(content);
    } else {
      body.innerHTML = "<div class='sws-card'><div class='sws-muted'>Порожній екран</div></div>";
    }

    return { screen, localCtx };
  }

  function push(opts){
    const st = $stack();
    if(!st) return;

    const r = renderScreen({
      title: opts.title,
      subtitle: opts.subtitle,
      content: opts.content,
      onSave: opts.onSave,
      saveLabel: opts.saveLabel,
      canSave: opts.canSave,
      ctx: opts.ctx || {},
    });

    const screenObj = {
      screenEl: r.screen,
      onSave: opts.onSave,
      canSave: opts.canSave,
      ctx: r.localCtx,
    };

    st.appendChild(r.screen);
    r.screen.getBoundingClientRect(); // reflow
    state.stack.push(screenObj);
    updateScreenClasses();

    setTimeout(()=>{
      const el = r.screen.querySelector("button, [tabindex], input, select, textarea");
      if(el && el.focus) el.focus();
    }, 0);
  }

  function pop(){
    if(state.stack.length <= 1){
      if(state.closeOnRootBack) SettingsWindow.close();
      return;
    }
    const leaving = state.stack.pop();
    updateScreenClasses();
    if(leaving?.screenEl){
      const el = leaving.screenEl;
      el.classList.remove("sws-active");
      el.style.transform = "translateX(100%)";
      setTimeout(()=>{ el.remove(); }, 240);
    }
  }

  function wrap(children){
    const d = document.createElement("div");
    (children||[]).forEach(ch => ch && d.appendChild(ch));
    return d;
  }

  function openRoot(opts){
    openOverlay();
    setWindowClass();

    const st = $stack();
    if(st) st.innerHTML = "";
    state.stack = [];

    const items = opts.items || [];
    push({
      title: opts.title || "Налаштування",
      subtitle: opts.subtitle || "Оберіть розділ",
      onSave: opts.onSave || null,
      saveLabel: opts.saveLabel || "Зберегти",
      canSave: opts.canSave || null,
      ctx: opts.ctx || {},
      content: (ctx)=>{
        const ui = ctx.ui;
        const hint = ui.card({
          title: "Налаштування",
          description: "Навігація: Esc — назад, Enter — зберегти. Кожен пункт відкриває наступне вікно поверх.",
          children: []
        });
        const list = ui.list(items);
        return wrap([hint, list]);
      }
    });
  }

  function openEmptyRoot(opts){
    openOverlay();
    setWindowClass();

    const st = $stack();
    if(st) st.innerHTML = "";
    state.stack = [];

    push({
      title: opts?.title || "Налаштування",
      subtitle: opts?.subtitle || "",
      onSave: opts?.onSave || null,
      saveLabel: opts?.saveLabel || "Зберегти",
      canSave: opts?.canSave || null,
      ctx: opts?.ctx || {},
      content: ()=> wrap([]) // IMPORTANT: empty body
    });
  }

  function onKeyDown(e){
    const ov = $overlay();
    if(!ov || !ov.classList.contains("sws-open")) return;

    if(e.key === "Escape"){
      e.preventDefault();
      SettingsWindow.pop();
      return;
    }
    if(e.key === "Enter"){
      if(isTextArea(document.activeElement)) return;
      e.preventDefault();
      const s = top();
      if(!s) return;
      if(s.canSave && !s.canSave(s.ctx)) return;
      if(typeof s.onSave === "function") s.onSave(s.ctx);
    }
  }

  function init(opts){
    if(!document.getElementById("swsOverlay")){
      console.warn("SettingsWindow.init(): markup not found. Insert sws_modal.html into DOM.");
      return;
    }
    state.inited = true;
    state.mode = opts?.mode || "auto";
    state.closeOnRootBack = (opts?.closeOnRootBack !== false);

    applyTheme(opts?.theme || "light");
    setWindowClass();
    if(opts?.tokens) applyTokens(opts.tokens);

    $overlay().addEventListener("mousedown", (ev)=>{
      if(ev.target === $overlay()) SettingsWindow.close();
    });

    window.addEventListener("keydown", onKeyDown, true);
  }

  const SettingsWindow = {
    init,
    openRoot,
    openEmptyRoot,
    push,
    pushList: (opts)=>{
      const items = opts.items || [];
      SettingsWindow.push({
        title: opts.title,
        subtitle: opts.subtitle,
        onSave: opts.onSave,
        saveLabel: opts.saveLabel,
        canSave: opts.canSave,
        ctx: opts.ctx,
        content: (ctx)=>{
          const ui = ctx.ui;
          const nodes = [];
          if(opts.introTitle || opts.introText){
            nodes.push(ui.card({ title: opts.introTitle || null, description: opts.introText || null, children: [] }));
          }
          nodes.push(ui.list(items));
          return wrap(nodes);
        }
      });
    },
    pop,
    close: closeOverlay,
    setTheme: applyTheme,
    setThemeTokens: applyTokens,
    getStackDepth: ()=> state.stack.length,
  };

  window.SettingsWindow = SettingsWindow;
})();
