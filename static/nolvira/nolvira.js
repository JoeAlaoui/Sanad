// ===================================================================================
// THÈME NOLVIRA — habillage chargé UNIQUEMENT sur /nolvira, APRÈS app.js.
// Ne touche à aucune logique métier : il réorganise quelques nœuds existants (en conservant
// leurs id, donc les gestionnaires d'app.js restent valides), remplace les pictos emoji par
// des icônes SVG au trait, ajoute bascule clair/sombre + repli de la sidebar, et aligne les
// couleurs Chart.js sur la palette Nolvira. Retirer ce fichier = revenir au thème actuel.
// ===================================================================================
(function(){
  const root = document.documentElement;
  const $ = id => document.getElementById(id);

  // ---------- Icônes (traits 24×24, style Lucide) ----------
  const P = {
    grid:'<rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/>',
    phone:'<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.8 19.8 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.12 4.18 2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.13.96.36 1.9.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.91.34 1.85.57 2.81.7A2 2 0 0 1 22 16.92z"/>',
    mail:'<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>',
    pie:'<path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/>',
    trend:'<path d="m22 7-8.5 8.5-5-5L2 17"/><path d="M16 7h6v6"/>',
    sparkles:'<path d="M12 3l1.9 5.1L19 10l-5.1 1.9L12 17l-1.9-5.1L5 10l5.1-1.9z"/><path d="M19 3v4M21 5h-4"/>',
    upload:'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5"/><path d="M12 3v12"/>',
    file:'<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8"/>',
    book:'<path d="M2 4h6a4 4 0 0 1 4 4v13a3 3 0 0 0-3-3H2z"/><path d="M22 4h-6a4 4 0 0 0-4 4v13a3 3 0 0 1 3-3h7z"/>',
    sliders:'<path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3M1 14h6M9 8h6M17 16h6"/>',
    users:'<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
    chat:'<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
    library:'<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20V2H6.5A2.5 2.5 0 0 0 4 4.5z"/><path d="M4 19.5A2.5 2.5 0 0 0 6.5 22H20v-5"/>',
    activity:'<path d="M22 12h-4l-3 9L9 3l-3 9H2"/>',
    shield:'<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>',
    db:'<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/><path d="M3 12c0 1.66 4 3 9 3s9-1.34 9-3"/>',
    star:'<path d="m12 2 3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01z"/>',
    search:'<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
    panel:'<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18"/><path d="m16 15-3-3 3-3"/>',
    sun:'<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/>',
    moon:'<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9z"/>',
    chevL:'<path d="m15 18-6-6 6-6"/>',
    chevR:'<path d="m9 18 6-6-6-6"/>',
  };
  const svg = (name, cls) => `<svg class="nv-i${cls ? " " + cls : ""}" viewBox="0 0 24 24" aria-hidden="true">${P[name] || ""}</svg>`;

  const TAB_ICONS = {
    "dashboard":"grid", "calls":"phone", "emails":"mail", "analysis":"pie", "annual":"trend", "prompts":"sparkles",
    "p-dashboard":"grid", "p-import":"upload", "p-livrables":"file", "p-analysis":"pie",
    "m-dashboard":"grid", "m-sources":"upload", "m-analysis":"pie", "m-annual":"trend", "m-guide":"book",
    "admin":"sliders", "users":"users", "notes":"chat", "knowledge":"library",
    "quality":"activity", "compliance":"shield", "backups":"db",
  };
  function iconizeNav(scope){
    (scope || document).querySelectorAll(".nav-item .nav-ico:not([data-nv])").forEach(el=>{
      const item = el.closest(".nav-item");
      const name = item.dataset.realModule ? "star" : TAB_ICONS[item.dataset.tab];
      if(!name) return;
      el.dataset.nv = "1";
      el.innerHTML = svg(name);           // l'emoji disparaît aussi du fil d'Ariane (textContent)
      item.title = item.textContent.trim().replace(/\s+/g, " ");
    });
  }
  iconizeNav();
  // Favoris : réinjectés par renderFavorites() → on ré-iconise à chaque mutation
  const favGroup = $("favoritesGroup");
  if(favGroup) new MutationObserver(()=>iconizeNav(favGroup)).observe(favGroup, {childList:true});
  document.querySelectorAll(".nav-subgroup-title").forEach(t=>{
    t.childNodes.forEach(n=>{ if(n.nodeType === 3) n.textContent = n.textContent.replace(/^\s+/, ""); });
  });
  // Le fil d'Ariane a pu être rendu avant l'iconisation (emoji dans le libellé) → rafraîchir
  const activeTab = document.querySelector(".nav-item.active")?.dataset.tab;
  if(activeTab && typeof updateBreadcrumb === "function") updateBreadcrumb(activeTab);

  // ---------- Sidebar : marque + pied utilisateur ----------
  const brand = $("sanadBrand");
  if(brand && !brand.querySelector(".nv-mark")){
    const mark = document.createElement("div");
    mark.className = "nv-mark";
    brand.prepend(mark);
    const syncMark = ()=>{ mark.textContent = ($("sanadLogoText")?.textContent.trim()[0] || "S").toUpperCase(); };
    syncMark();
    new MutationObserver(syncMark).observe($("sanadLogoText"), {childList:true, characterData:true, subtree:true});
  }

  const initials = name => (name || "").trim().split(/\s+/).filter(Boolean).slice(0, 2).map(w=>w[0]).join("").toUpperCase() || "?";
  const footer = document.querySelector(".sidebar-footer");
  const nameEl = $("userMenuName"), roleEl = $("userMenuRole");
  const sideAvatar = document.createElement("div");
  sideAvatar.className = "nv-avatar";
  if(footer && nameEl){
    const org = footer.textContent.trim();
    footer.textContent = "";
    const box = document.createElement("div");
    box.className = "nv-user";
    const orgEl = document.createElement("div");
    orgEl.className = "nv-org";
    box.append(nameEl, orgEl);
    footer.append(sideAvatar, box);
    const syncUser = ()=>{
      sideAvatar.textContent = initials(nameEl.textContent);
      topAvatar.textContent = sideAvatar.textContent;
      topAvatar.title = nameEl.textContent + (roleEl ? " — " + roleEl.textContent : "");
      orgEl.textContent = [roleEl?.textContent, org].filter(Boolean).join(" · ");
    };
    var topAvatar = document.createElement("div");
    topAvatar.className = "nv-avatar";
    syncUser();
    new MutationObserver(syncUser).observe(nameEl, {childList:true, characterData:true, subtree:true});
    if(roleEl) new MutationObserver(syncUser).observe(roleEl, {childList:true, characterData:true, subtree:true});
  }

  // ---------- Topbar : [repli] [recherche ⌃K] ……… [mois] [sync] [thème] [actions] | [avatar] ----------
  const topbar = document.querySelector(".topbar");
  const main = document.querySelector(".main");
  if(topbar && main){
    const collapseBtn = document.createElement("button");
    collapseBtn.type = "button";
    collapseBtn.className = "nv-icon-btn";
    collapseBtn.title = "Replier / déplier la barre latérale";
    collapseBtn.innerHTML = svg("panel");
    collapseBtn.addEventListener("click", ()=>{
      const c = root.classList.toggle("nv-collapsed");
      try{ localStorage.setItem("nolvira_sidebar", c ? "collapsed" : "expanded"); }catch(e){}
    });

    const searchWrap = topbar.querySelector(".global-search-wrap");
    if(searchWrap){
      searchWrap.insertAdjacentHTML("afterbegin", svg("search", "nv-search-ico"));
      searchWrap.insertAdjacentHTML("beforeend", "<kbd>Ctrl K</kbd>");
      const gs = $("globalSearch");
      if(gs) gs.placeholder = "Rechercher… (mois, catégories, contacts)";
    }

    // Fil d'Ariane + favori : sortis du topbar vers un bandeau de page (id conservés)
    const pagebar = document.createElement("div");
    pagebar.className = "nv-pagebar";
    const bc = $("breadcrumb"), fav = $("favoriteToggleBtn");
    if(bc) pagebar.append(bc);
    if(fav) pagebar.append(fav);
    topbar.after(pagebar);

    const spacer = document.createElement("div");
    spacer.className = "nv-spacer";

    const themeBtn = document.createElement("button");
    themeBtn.type = "button";
    themeBtn.className = "nv-icon-btn";
    const paintThemeBtn = ()=>{
      const dark = root.classList.contains("dark");
      themeBtn.innerHTML = svg(dark ? "sun" : "moon");
      themeBtn.title = dark ? "Passer en thème clair" : "Passer en thème sombre";
    };
    paintThemeBtn();
    themeBtn.addEventListener("click", ()=>{
      const dark = root.classList.toggle("dark");
      try{ localStorage.setItem("nolvira_theme", dark ? "dark" : "light"); }catch(e){}
      paintThemeBtn();
      applyChartTheme();
      rerenderActiveTab();
    });

    const prevM = $("monthPrevBtn"), nextM = $("monthNextBtn");
    if(prevM) prevM.innerHTML = svg("chevL");
    if(nextM) nextM.innerHTML = svg("chevR");

    const sep = document.createElement("div");
    sep.className = "nv-topbar-sep";

    topbar.prepend(collapseBtn);
    if(searchWrap) collapseBtn.after(searchWrap);
    (searchWrap || collapseBtn).after(spacer);
    const userMenu = topbar.querySelector(".user-menu");
    if(userMenu){
      userMenu.prepend(themeBtn);
      const logout = $("logoutBtn");
      if(logout) userMenu.insertBefore(sep, logout);
      if(topAvatar) userMenu.insertBefore(topAvatar, logout || null);
    }
  }

  // Ctrl/Cmd + K → recherche globale
  document.addEventListener("keydown", e=>{
    if((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k"){
      const gs = $("globalSearch");
      if(gs && gs.offsetParent !== null){ e.preventDefault(); gs.focus(); gs.select(); }
    }
  });

  // ---------- Chart.js : typographie, grilles et palette Nolvira ----------
  // Les couleurs des séries sont codées dans app.js : on les transpose à la volée (plugin),
  // sans modifier app.js. Les exports serveur (PDF/PPTX/Excel/email) ne sont pas concernés.
  const COLOR_MAP = {
    // teinte conservée pour les couleurs sémantiques (statuts PCHC vert/bleu/orange/rouge)
    "#0B4965":"#0D9488", "#1794CF":"#2563EB", "#25935F":"#16A34A", "#6FC49A":"#5EC79A",
    "#F59F0A":"#D97706", "#F2994A":"#EA580C", "#DC2828":"#DC2626", "#8E5AA6":"#7C3AED",
    "#A9C7D6":"#9FD8D1", "#B9D3E0":"#B3E0DA", "#55616B":"#706E66", "#94A3AD":"#9CA3AF", "#0D1926":"#1B1B19",
  };
  const remap = v=>{
    if(typeof v === "string" && v[0] === "#" && v.length >= 7){
      const m = COLOR_MAP[v.slice(0, 7).toUpperCase()];
      return m ? m + v.slice(7) : v;
    }
    if(Array.isArray(v)) return v.map(remap);
    return v;
  };
  const COLOR_KEYS = ["backgroundColor", "borderColor", "pointBackgroundColor", "pointBorderColor", "hoverBackgroundColor", "hoverBorderColor"];
  const css = n => getComputedStyle(root).getPropertyValue(n).trim();
  function applyChartTheme(){
    if(!window.Chart) return;
    Chart.defaults.font.family = "Inter, system-ui, sans-serif";
    Chart.defaults.font.size = 12;
    Chart.defaults.color = css("--nv-muted");
    Chart.defaults.borderColor = css("--nv-border");
    Chart.defaults.plugins.legend.labels.usePointStyle = true;
    Chart.defaults.plugins.legend.labels.pointStyle = "rectRounded";
    Chart.defaults.plugins.legend.labels.boxWidth = 8;
    Chart.defaults.plugins.legend.labels.boxHeight = 8;
    Chart.defaults.plugins.tooltip.backgroundColor = css("--nv-fg");
    Chart.defaults.plugins.tooltip.titleColor = css("--nv-bg");
    Chart.defaults.plugins.tooltip.bodyColor = css("--nv-bg");
    Chart.defaults.plugins.tooltip.cornerRadius = 8;
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.elements.bar.borderRadius = 4;
    Chart.defaults.elements.arc.borderColor = css("--nv-surface");
  }
  if(window.Chart){
    applyChartTheme();
    Chart.register({
      id:"nolviraPalette",
      beforeUpdate(chart){
        (chart.config.data.datasets || []).forEach(ds=>{
          if(ds.__nv) return;
          COLOR_KEYS.forEach(k=>{ if(ds[k] !== undefined) ds[k] = remap(ds[k]); });
          ds.__nv = true;
        });
      },
    });
  }

  function rerenderActiveTab(){
    const tab = document.querySelector(".nav-item.active:not([data-real-module])")?.dataset.tab
             || document.querySelector(".nav-item.active")?.dataset.tab;
    if(tab && typeof renderTab === "function"){ try{ renderTab(tab); }catch(e){ console.warn(e); } }
  }
})();
