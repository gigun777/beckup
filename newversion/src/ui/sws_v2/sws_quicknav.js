(function(){
  // SWS Quick Navigation Panel (ported from project quicknav.js to SettingsWindowSystem v2)
  // - No external deps
  // - Designed to be embedded inside SettingsWindow.push({content})

  const norm = (v)=>String(v||'').toLowerCase().trim();

  function rand1to9(){ return Math.floor(Math.random()*9)+1; }
  async function confirmDeleteNumber(title){
    const n = rand1to9();
    const v = prompt(`${title}\n\nПідтвердження видалення.\nВведіть число: ${n}`);
    if(v===null) return false;
    return String(v).trim() === String(n);
  }

  function mkMiniBtn(ui, {text, title, variant}){
    const b = ui.el('button', `sws-qnav-mini ${variant||''}`.trim(), text);
    b.type = 'button';
    if(title) b.title = title;
    return b;
  }

  function mkCaret(ui, {hasKids, expanded, title}){
    const b = mkMiniBtn(ui, {text: hasKids ? (expanded ? '▾' : '▸') : ' ', title: title||''});
    b.classList.add('sws-qnav-caret');
    if(!hasKids){ b.style.opacity='0'; b.style.pointerEvents='none'; }
    return b;
  }

  /**
   * Create an embeddable quick nav panel.
   *
   * opts:
   *  - ui: ctx.ui from SettingsWindow
   *  - getData(): Promise<{spaces, activeSpaceId, jtree, activeJournalId}>
   *  - showSpaces, showJournals, allowAdd, allowDelete
   *  - onGoSpace(spaceId), onAddSpace(parentSpaceId), onDeleteSpace(spaceId)
   *  - onGoJournalPath(pathIds), onAddJournalChild(pathIds), onDeleteJournal(journalId)
   */
  async function createPanel(opts){
    const {
      ui,
      title = 'Швидка навігація',
      showSpaces = true,
      showJournals = true,
      allowAdd = true,
      allowDelete = true,
      defaultCollapsed = true,
      getData,
      onGoSpace,
      onAddSpace,
      onDeleteSpace,
      onGoJournalPath,
      onAddJournalChild,
      onDeleteJournal,
    } = opts || {};

    if(!ui) throw new Error('SWSQuickNav.createPanel(): opts.ui is required');

    // Live data snapshots
    let _spaces = [];
    let _activeSpaceId = null;
    let _jtree = null;
    let _activeJournalId = null;

    async function refreshData(){
      if(typeof getData !== 'function') return;
      const d = await getData();
      if(!d || typeof d !== 'object') return;
      if(Array.isArray(d.spaces)) _spaces = d.spaces;
      if(typeof d.activeSpaceId === 'string' || d.activeSpaceId===null) _activeSpaceId = d.activeSpaceId;
      if(d.jtree && typeof d.jtree === 'object') _jtree = d.jtree;
      if(typeof d.activeJournalId === 'string' || d.activeJournalId===null) _activeJournalId = d.activeJournalId;
    }

    await refreshData();

    // Indexes (rebuilt on demand)
    let onlySpaces = [];
    let bySpaceId = {};
    let spaceChildren = {};
    let nodes = {};
    let topIds = [];

    const nodeById = (id)=>nodes[id] || null;
    const nodeTitle = (n)=>{
      if(!n) return '';
      if(n.title) return String(n.title);
      if(n.key) return String(n.key);
      return String(n.id||'');
    };

    function rebuildIndexes(){
      onlySpaces = (_spaces||[]).filter(s=>s && (s.kind==='space' || !s.kind));
      bySpaceId = Object.fromEntries(onlySpaces.map(s=>[s.id, s]));
      spaceChildren = {};
      for(const s of onlySpaces){
        const pid = s.parentId || null;
        if(!spaceChildren[pid]) spaceChildren[pid] = [];
        spaceChildren[pid].push(s);
      }
      for(const k of Object.keys(spaceChildren)){
        spaceChildren[k].sort((a,b)=>String(a.name||'').localeCompare(String(b.name||'')));
      }
      nodes = (_jtree && _jtree.nodes) ? _jtree.nodes : {};
      topIds = (_jtree && Array.isArray(_jtree.topIds)) ? _jtree.topIds : [];
    }
    rebuildIndexes();

    // Root container
    const root = ui.el('div', 'sws-qnav');

    // Header line
    const head = ui.el('div', 'sws-qnav-head');
    head.appendChild(ui.el('div','sws-muted', title));
    const toggle = mkMiniBtn(ui, {text: defaultCollapsed ? '▸' : '▾', title:'Перемикає: гілки згорнуті/розгорнуті'});
    toggle.classList.add('sws-qnav-toggle');
    head.appendChild(toggle);
    root.appendChild(head);

    // Search line
    const srow = ui.el('div','sws-qnav-search');
    const search = document.createElement('input');
    search.type = 'search';
    search.className = 'sws-input';
    search.placeholder = 'Пошук (простір/журнал)…';
    search.style.flex = '1';
    search.style.minWidth = '0';
    const clear = mkMiniBtn(ui, {text:'✕', title:'Очистити пошук'});
    clear.onclick = (e)=>{ e.preventDefault(); e.stopPropagation(); search.value=''; renderTree(); search.focus(); };
    srow.appendChild(search);
    srow.appendChild(clear);
    root.appendChild(srow);

    // Tree container
    const tree = ui.el('div','sws-qnav-tree');
    root.appendChild(tree);

    // Expanded state
    const expandedSpaces = new Set();
    const expandedJournals = new Set();
    let collapsedByDefault = !!defaultCollapsed;

    const ensureDefaultExpanded = ()=>{
      if(collapsedByDefault){
        expandedSpaces.clear();
        expandedJournals.clear();
        return;
      }
      expandedSpaces.clear();
      expandedJournals.clear();
      for(const s of onlySpaces){ if((spaceChildren[s.id]||[]).length) expandedSpaces.add(s.id); }
      for(const id of Object.keys(nodes)){
        const n = nodes[id];
        const kids = (n?.children||[]).filter(cid=>!!nodes[cid]);
        if(kids.length) expandedJournals.add(id);
      }
    };

    // Search filters
    let visibleSpaceIds = null;
    let visibleJournalIds = null;

    const makeSpaceNode = (s, depth=0, num='')=>{
      if(visibleSpaceIds && !visibleSpaceIds.has(s.id)) return [];
      const kidsAll = spaceChildren[s.id] || [];
      const kids = visibleSpaceIds ? kidsAll.filter(k=>visibleSpaceIds.has(k.id)) : kidsAll;
      if(!collapsedByDefault && kids.length && !expandedSpaces.has(s.id)) expandedSpaces.add(s.id);

      const row = ui.el('div','sws-qnav-row');
      if(s.id===_activeSpaceId) row.classList.add('sws-qnav-active');

      const caret = mkCaret(ui, {hasKids: !!kids.length, expanded: expandedSpaces.has(s.id), title: kids.length ? 'Згорнути/розгорнути' : ''});
      const label = num ? `${num} ${s.name}` : `${s.name}`;
      const b = ui.el('button','sws-qnav-btn', `📁 ${label}`);
      b.style.marginLeft = (depth*14)+'px';
      b.type='button';
      b.onclick = async ()=>{
        if(onGoSpace){
          await onGoSpace(s.id);
          await refreshData();
          rebuildIndexes();
          renderTree();
        }
      };

      const addBtn = mkMiniBtn(ui, {text:'＋', title:'Додати підпростір'});
      if(!allowAdd || !onAddSpace) addBtn.style.display='none';
      addBtn.onclick = async (e)=>{
        e.preventDefault(); e.stopPropagation();
        await onAddSpace(s.id);
        await refreshData();
        rebuildIndexes();
        renderTree();
      };

      const delBtn = mkMiniBtn(ui, {text:'🗑', title:'Видалити простір', variant:'danger'});
      if(!allowDelete || !onDeleteSpace) delBtn.style.display='none';
      delBtn.onclick = async (e)=>{
        e.preventDefault(); e.stopPropagation();
        const ok = await confirmDeleteNumber(`Видалити простір "${(bySpaceId[s.id]?.name)||s.id}"?\n\nУвага: будуть видалені також усі підпростори.`);
        if(!ok) return;
        await onDeleteSpace(s.id);
        await refreshData();
        rebuildIndexes();
        renderTree();
      };

      const childWrap = ui.el('div','');
      childWrap.style.display = expandedSpaces.has(s.id) ? '' : 'none';
      for(let i=0;i<kids.length;i++){
        const c = kids[i];
        const cnum = num ? `${num}.${i+1}` : `${i+1}`;
        for(const n of makeSpaceNode(c, depth+1, cnum)) childWrap.appendChild(n);
      }

      caret.onclick = (e)=>{
        e.preventDefault(); e.stopPropagation();
        const next = !expandedSpaces.has(s.id);
        if(next) expandedSpaces.add(s.id); else expandedSpaces.delete(s.id);
        caret.textContent = next ? '▾' : '▸';
        childWrap.style.display = next ? '' : 'none';
      };

      row.appendChild(caret);
      row.appendChild(b);
      row.appendChild(addBtn);
      row.appendChild(delBtn);
      return [row, childWrap];
    };

    const makeJournalNode = (id, path, depth=0, num='')=>{
      if(visibleJournalIds && !visibleJournalIds.has(id)) return [];
      const n = nodeById(id);
      if(!n) return [];
      const kidsAll = (n.children||[]).filter(cid=>!!nodeById(cid));
      const kids = visibleJournalIds ? kidsAll.filter(cid=>visibleJournalIds.has(cid)) : kidsAll;
      if(!collapsedByDefault && kids.length && !expandedJournals.has(id)) expandedJournals.add(id);

      const row = ui.el('div','sws-qnav-row');
      if(_activeJournalId===id) row.classList.add('sws-qnav-active');

      const caret = mkCaret(ui, {hasKids: !!kids.length, expanded: expandedJournals.has(id), title: kids.length ? 'Згорнути/розгорнути' : ''});
      const label = num ? `${num} ${nodeTitle(n)}` : `${nodeTitle(n)}`;
      const b = ui.el('button','sws-qnav-btn', `📄 ${label}`);
      b.style.marginLeft = (depth*14)+'px';
      b.type='button';
      b.onclick = async ()=>{
        if(onGoJournalPath){
          await onGoJournalPath(path.slice());
          await refreshData();
          rebuildIndexes();
          renderTree();
        }
      };

      const addBtn = mkMiniBtn(ui, {text:'＋', title:'Додати піджурнал'});
      if(!allowAdd || !onAddJournalChild) addBtn.style.display='none';
      addBtn.onclick = async (e)=>{
        e.preventDefault(); e.stopPropagation();
        await onAddJournalChild(path.slice());
        await refreshData();
        rebuildIndexes();
        renderTree();
      };

      const canDelete = !String(id).startsWith('root:');
      const delBtn = mkMiniBtn(ui, {text:'🗑', title: canDelete ? 'Видалити журнал' : 'Кореневий журнал видаляти не можна', variant:'danger'});
      if(!allowDelete || !onDeleteJournal) delBtn.style.display='none';
      else if(!canDelete){ delBtn.style.opacity='0.35'; delBtn.style.pointerEvents='none'; }
      delBtn.onclick = async (e)=>{
        e.preventDefault(); e.stopPropagation();
        if(!canDelete) return;
        const ok = await confirmDeleteNumber(`Видалити журнал "${nodeTitle(n)}"?\n\nУвага: будуть видалені також усі його піджурнали.`);
        if(!ok) return;
        await onDeleteJournal(id);
        await refreshData();
        rebuildIndexes();
        renderTree();
      };

      const childWrap = ui.el('div','');
      childWrap.style.display = expandedJournals.has(id) ? '' : 'none';
      for(let i=0;i<kids.length;i++){
        const cid = kids[i];
        const cnum = num ? `${num}.${i+1}` : `${i+1}`;
        for(const line of makeJournalNode(cid, path.concat([cid]), depth+1, cnum)) childWrap.appendChild(line);
      }

      caret.onclick = (e)=>{
        e.preventDefault(); e.stopPropagation();
        const next = !expandedJournals.has(id);
        if(next) expandedJournals.add(id); else expandedJournals.delete(id);
        caret.textContent = next ? '▾' : '▸';
        childWrap.style.display = next ? '' : 'none';
      };

      row.appendChild(caret);
      row.appendChild(b);
      row.appendChild(addBtn);
      row.appendChild(delBtn);
      return [row, childWrap];
    };

    const renderTree = ()=>{
      tree.innerHTML='';
      const q = norm(search.value||'');
      const searching = !!q;

      // Visible sets for search
      if(searching){
        const vS = new Set();
        const vJ = new Set();
        const expS = new Set();
        const expJ = new Set();

        // Spaces: match + ancestors
        for(const s of onlySpaces){
          if(norm(s.name).includes(q)){
            let cur = s;
            while(cur){
              vS.add(cur.id);
              const pid = cur.parentId || null;
              if(pid) expS.add(pid);
              cur = pid ? bySpaceId[pid] : null;
            }
          }
        }

        // Journals: match + ancestors
        for(const id of Object.keys(nodes)){
          const n = nodes[id];
          if(norm(nodeTitle(n)).includes(q)){
            let cur = n;
            while(cur){
              vJ.add(cur.id);
              const pid = cur.parentId || null;
              if(pid) expJ.add(pid);
              cur = pid ? nodes[pid] : null;
            }
          }
        }

        visibleSpaceIds = vS;
        visibleJournalIds = vJ;
        expandedSpaces.clear();
        expandedJournals.clear();
        for(const id of expS) expandedSpaces.add(id);
        for(const id of expJ) expandedJournals.add(id);
      }else{
        visibleSpaceIds = null;
        visibleJournalIds = null;
        ensureDefaultExpanded();
      }

      if(showSpaces){
        tree.appendChild(ui.el('div','sws-qnav-section sws-muted','Простори'));
        const roots = spaceChildren[null] || spaceChildren[undefined] || [];
        for(let i=0;i<roots.length;i++){
          const r = roots[i];
          if(searching && visibleSpaceIds && visibleSpaceIds.size && !visibleSpaceIds.has(r.id)) continue;
          for(const n of makeSpaceNode(r, 0, String(i+1))) tree.appendChild(n);
        }
      }

      if(showJournals){
        tree.appendChild(ui.el('div','sws-qnav-section sws-muted', showSpaces ? 'Журнали (поточний простір)' : 'Журнали'));
        const top = (topIds||[]).filter(id=>!!nodeById(id));
        for(let i=0;i<top.length;i++){
          const id = top[i];
          if(searching && visibleJournalIds && visibleJournalIds.size && !visibleJournalIds.has(id)) continue;
          for(const n of makeJournalNode(id, [id], 0, String(i+1))) tree.appendChild(n);
        }
      }

      if(searching){
        const hasAny = (visibleSpaceIds?.size||0) > 0 || (visibleJournalIds?.size||0) > 0;
        if(!hasAny) tree.appendChild(ui.el('div','sws-muted','Нічого не знайдено.'));
      }
    };

    search.oninput = ()=>renderTree();
    toggle.onclick = (e)=>{
      e.preventDefault(); e.stopPropagation();
      collapsedByDefault = !collapsedByDefault;
      toggle.textContent = collapsedByDefault ? '▸' : '▾';
      renderTree();
    };

    // Initial render
    renderTree();

    return { root, refresh: async ()=>{ await refreshData(); rebuildIndexes(); renderTree(); } };
  }

  function openQuickNavScreen({
    title='Спрощена навігація',
    subtitle='Дерево просторів і журналів',
    getData,
    showSpaces=true,
    showJournals=true,
    allowAdd=true,
    allowDelete=true,
    defaultCollapsed=true,
    onGoSpace,
    onAddSpace,
    onDeleteSpace,
    onGoJournalPath,
    onAddJournalChild,
    onDeleteJournal,
  }={}){
    if(!window.SettingsWindow) throw new Error('SettingsWindow is not available');
    window.SettingsWindow.push({
      title,
      subtitle,
      ctx: { model: {} },
      content: (ctx)=>{
        const wrap = ctx.ui.el('div','');
        // NOTE: content() cannot be async in SWS, so we create a placeholder and hydrate.
        const placeholder = ctx.ui.el('div','sws-muted','Завантаження…');
        wrap.appendChild(placeholder);
        (async ()=>{
          try{
            const panel = await createPanel({
              ui: ctx.ui,
              title: 'Швидка навігація',
              showSpaces,
              showJournals,
              allowAdd,
              allowDelete,
              defaultCollapsed,
              getData,
              onGoSpace,
              onAddSpace,
              onDeleteSpace,
              onGoJournalPath,
              onAddJournalChild,
              onDeleteJournal,
            });
            wrap.replaceChild(panel.root, placeholder);
          }catch(e){
            placeholder.textContent = 'Помилка завантаження навігації: ' + (e?.message || e);
          }
        })();
        return wrap;
      },
      // no save for nav
      onSave: null,
      saveLabel: 'OK',
      canSave: ()=>false,
    });
  }

  window.SWSQuickNav = { createPanel, openQuickNavScreen };
})();
