(function () {
  var videos = Array.prototype.slice.call(document.querySelectorAll(".card-media video"));

  videos.forEach(function (v) {
    v.muted = true;
    v.playsInline = true;
  });

  if ("IntersectionObserver" in window) {
    var io = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          var v = entry.target;
          if (entry.isIntersecting) {
            v.play().catch(function () {});
          } else {
            v.pause();
          }
        });
      },
      { root: null, rootMargin: "0px 400px 0px 400px", threshold: 0.01 }
    );
    videos.forEach(function (v) { io.observe(v); });
  } else {
    videos.forEach(function (v) { v.play().catch(function () {}); });
  }

  var track = document.getElementById("stripTrack");
  if (track) {
    track.addEventListener("focusin", function () { track.style.animationPlayState = "paused"; });
    track.addEventListener("focusout", function () { track.style.animationPlayState = "running"; });
  }
})();
