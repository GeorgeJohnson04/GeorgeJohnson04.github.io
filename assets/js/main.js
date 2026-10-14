/* George Johnson — highlight the current page in the nav and set the footer year */
(function () {
  "use strict";

  var path = location.pathname.split("/").pop() || "index.html";
  document.querySelectorAll(".site-nav a").forEach(function (a) {
    if ((a.getAttribute("href") || "") === path) a.classList.add("active");
  });

  var y = document.getElementById("year");
  if (y) y.textContent = new Date().getFullYear();
})();
