/* bh-shell.js — opens and closes the slide-out side menu below 960px.
   Needs: a .bh-menu-btn in the header, a .bh-side-close inside .bh-side, and a <button class="bh-scrim"> after .bh-side. */
(function(){
  function set(open){
    document.body.classList.toggle('bh-menu-open', open);
    var b=document.querySelector('.bh-menu-btn'); if(b) b.setAttribute('aria-expanded', open?'true':'false');
    if(open){var f=document.querySelector('.bh-side .bh-side-close'); if(f) f.focus();}
  }
  document.addEventListener('click', function(e){
    if(e.target.closest('.bh-menu-btn')) set(true);
    else if(e.target.closest('.bh-side-close, .bh-scrim')) set(false);
    else if(e.target.closest('.bh-side .bh-side-item') && matchMedia('(max-width:960px)').matches) set(false);
  });
  document.addEventListener('keydown', function(e){ if(e.key==='Escape' && document.body.classList.contains('bh-menu-open')) set(false); });
})();
