(function () {
  var toggle = document.getElementById("navToggle");
  var links = document.getElementById("navLinks");
  if (!toggle || !links) return;

  toggle.addEventListener("click", function () {
    var open = links.classList.toggle("is-open");
    toggle.setAttribute("aria-expanded", open ? "true" : "false");
  });

  Array.prototype.slice.call(links.querySelectorAll("a")).forEach(function (a) {
    a.addEventListener("click", function () {
      links.classList.remove("is-open");
      toggle.setAttribute("aria-expanded", "false");
    });
  });
})();
