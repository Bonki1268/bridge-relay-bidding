(function(){
  "use strict";

  // 手機版導覽選單開關
  var nav = document.querySelector(".site-nav");
  var toggle = document.querySelector(".nav-toggle");
  if (toggle && nav) {
    toggle.addEventListener("click", function(){
      nav.classList.toggle("menu-open");
    });
  }

  // 下拉選單（點擊切換，含桌機 hover 與行動點擊）
  var items = document.querySelectorAll(".nav-item.has-menu");
  items.forEach(function(item){
    var link = item.querySelector(".nav-link");
    link.addEventListener("click", function(e){
      e.preventDefault();
      var isOpen = item.classList.contains("open");
      items.forEach(function(i){ i.classList.remove("open"); });
      if (!isOpen) item.classList.add("open");
    });
  });
  document.addEventListener("click", function(e){
    if (!e.target.closest(".nav-item")) {
      items.forEach(function(i){ i.classList.remove("open"); });
    }
  });

  // 叫序資料頁：工作表分頁籤
  var tabGroups = document.querySelectorAll("[data-sheet-tabs]");
  tabGroups.forEach(function(group){
    var tabs = group.querySelectorAll(".sheet-tab");
    var panelWrap = document.querySelector(group.getAttribute("data-sheet-tabs"));
    tabs.forEach(function(tab){
      tab.addEventListener("click", function(){
        var target = tab.getAttribute("data-target");
        tabs.forEach(function(t){ t.classList.remove("is-active"); });
        tab.classList.add("is-active");
        panelWrap.querySelectorAll(".sheet-panel").forEach(function(p){
          p.classList.toggle("is-active", p.id === target);
        });
      });
    });
  });

  // 叫序資料頁：分支跳轉連結（切換到目標分頁籤，並捲動到對應分支）
  document.addEventListener("click", function(e){
    var link = e.target.closest(".jump-link");
    if (!link) return;
    e.preventDefault();
    var sheetId = link.getAttribute("data-sheet");
    var anchorId = link.getAttribute("data-anchor");
    if (!sheetId) return;
    var tabBtn = document.querySelector('.sheet-tab[data-target="' + sheetId + '"]');
    if (tabBtn) tabBtn.click();
    var targetEl = document.getElementById(anchorId || sheetId);
    if (targetEl) {
      window.requestAnimationFrame(function(){
        targetEl.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    }
  });
})();
