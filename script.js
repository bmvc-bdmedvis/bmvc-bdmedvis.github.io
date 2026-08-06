const header = document.querySelector("[data-header]");
const navToggle = document.querySelector("[data-nav-toggle]");
const navLinks = document.querySelector("[data-nav-links]");
const progressBar = document.querySelector("[data-progress]");
const navAnchors = [...document.querySelectorAll('.site-nav a[href^="#"]')];
const sections = navAnchors
  .map((anchor) => document.querySelector(anchor.getAttribute("href")))
  .filter(Boolean);

const setHeaderState = () => {
  header?.classList.toggle("is-scrolled", window.scrollY > 8);
};

const setProgress = () => {
  if (!progressBar) return;
  const scrollable = document.documentElement.scrollHeight - window.innerHeight;
  const progress = scrollable > 0 ? (window.scrollY / scrollable) * 100 : 0;
  progressBar.style.width = `${Math.min(progress, 100)}%`;
};

const setActiveNav = () => {
  const current = sections
    .filter((section) => section.getBoundingClientRect().top <= 140)
    .at(-1);

  navAnchors.forEach((anchor) => {
    const isActive = current && anchor.getAttribute("href") === `#${current.id}`;
    anchor.classList.toggle("active", Boolean(isActive));
  });
};

const closeNav = () => {
  document.body.classList.remove("nav-open");
  header?.classList.remove("is-open");
  navLinks?.classList.remove("is-open");
  navToggle?.setAttribute("aria-expanded", "false");
};

setHeaderState();
setProgress();
setActiveNav();

window.addEventListener(
  "scroll",
  () => {
    setHeaderState();
    setProgress();
    setActiveNav();
  },
  { passive: true }
);

navToggle?.addEventListener("click", () => {
  const isOpen = navLinks?.classList.toggle("is-open") ?? false;
  document.body.classList.toggle("nav-open", isOpen);
  header?.classList.toggle("is-open", isOpen);
  navToggle.setAttribute("aria-expanded", String(isOpen));
});

navLinks?.addEventListener("click", (event) => {
  if (event.target instanceof HTMLAnchorElement) {
    closeNav();
  }
});

window.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeNav();
  }
});
