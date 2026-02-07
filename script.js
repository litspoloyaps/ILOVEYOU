let isOpen = false;

function toggleEnvelope() {
  innerPolygon.classList.toggle("inner-open");
  outer.classList.toggle("outer-open");
  heartBtn.classList.toggle("hide");
  closeBtn.classList.toggle("show");
  message1.classList.toggle("hide");
  message2.classList.toggle("show");

  // Open link only the FIRST time
  if (!isOpen) {
    window.open("https://litspoloyaps.github.io/CybersecurityDataManagement/", "_blank");
    isOpen = true;
  }
}
