/* ============================================================
   APEX WRESTLING ROBE — Main Script
   ============================================================ */

document.addEventListener('DOMContentLoaded', () => {

  /* ── Nav Scroll State ── */
  const nav = document.querySelector('.nav');
  if (nav) {
    window.addEventListener('scroll', () => {
      nav.classList.toggle('scrolled', window.scrollY > 40);
    }, { passive: true });
  }

  /* ── Hamburger Menu ── */
  const hamburger = document.querySelector('.nav-hamburger');
  const navLinks  = document.querySelector('.nav-links');
  if (hamburger && navLinks) {
    hamburger.addEventListener('click', () => {
      const open = navLinks.classList.toggle('open');
      hamburger.classList.toggle('open', open);
      hamburger.setAttribute('aria-expanded', open);
      document.body.style.overflow = open ? 'hidden' : '';
    });
    // Close on nav link click
    navLinks.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        navLinks.classList.remove('open');
        hamburger.classList.remove('open');
        hamburger.setAttribute('aria-expanded', 'false');
        document.body.style.overflow = '';
      });
    });
  }

  /* ── Active Nav Link ── */
  const currentPage = location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav-links a').forEach(link => {
    const href = link.getAttribute('href');
    if (href === currentPage) link.classList.add('active');
  });

  /* ── Scroll Reveal ── */
  const revealEls = document.querySelectorAll('.reveal, .reveal-left, .reveal-right');
  if (revealEls.length) {
    const observer = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add('visible');
          observer.unobserve(entry.target);
        }
      });
    }, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });
    revealEls.forEach(el => observer.observe(el));
  }

  /* ── Slideshow (product / gallery) ── */
  function initSlideshow(containerSelector, dotsSelector) {
    const container = document.querySelector(containerSelector);
    if (!container) return;
    const slides = container.querySelectorAll('.slide, .gallery-slide');
    const dots   = document.querySelectorAll(dotsSelector);
    if (!slides.length) return;

    let current = 0;
    let timer;

    const goTo = (i) => {
      slides[current].classList.remove('active');
      if (dots[current]) dots[current].classList.remove('active');
      current = (i + slides.length) % slides.length;
      slides[current].classList.add('active');
      if (dots[current]) dots[current].classList.add('active');
    };

    const start = () => { timer = setInterval(() => goTo(current + 1), 4500); };
    const stop  = () => clearInterval(timer);

    // Init
    slides[0].classList.add('active');
    if (dots[0]) dots[0].classList.add('active');
    start();

    // Dots
    dots.forEach((dot, i) => {
      dot.addEventListener('click', () => { stop(); goTo(i); start(); });
    });

    // Arrows
    const prevBtn = container.querySelector('.slide-arrow.prev, .gallery-prev');
    const nextBtn = container.querySelector('.slide-arrow.next, .gallery-next');
    if (prevBtn) prevBtn.addEventListener('click', () => { stop(); goTo(current - 1); start(); });
    if (nextBtn) nextBtn.addEventListener('click', () => { stop(); goTo(current + 1); start(); });

    // Pause on hover
    container.addEventListener('mouseenter', stop);
    container.addEventListener('mouseleave', start);
  }

  initSlideshow('.slideshow',        '.slide-dots .dot');
  initSlideshow('.gallery-slideshow', '.gallery-dots .dot');

  /* ── Stat Counter Animation ── */
  const stats = document.querySelectorAll('.stat-number[data-count]');
  if (stats.length) {
    const countObserver = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (!entry.isIntersecting) return;
        const el = entry.target;
        const target = parseFloat(el.dataset.count);
        const suffix = el.dataset.suffix || '';
        const prefix = el.dataset.prefix || '';
        const decimals = el.dataset.decimals ? parseInt(el.dataset.decimals) : 0;
        const duration = 1800;
        const start = performance.now();

        const tick = (now) => {
          const progress = Math.min((now - start) / duration, 1);
          const eased = 1 - Math.pow(1 - progress, 3);
          const value = (eased * target).toFixed(decimals);
          el.textContent = prefix + value + suffix;
          if (progress < 1) requestAnimationFrame(tick);
        };
        requestAnimationFrame(tick);
        countObserver.unobserve(el);
      });
    }, { threshold: 0.5 });
    stats.forEach(el => countObserver.observe(el));
  }

  /* ── Contact Form ── */
  const contactForm = document.getElementById('contactForm');
  if (contactForm) {
    contactForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      const btn = contactForm.querySelector('.btn-submit');
      const originalText = btn.innerHTML;

      btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Sending...';
      btn.disabled = true;

      try {
        const res = await fetch(contactForm.action, {
          method: 'POST',
          body: new FormData(contactForm),
          headers: { 'Accept': 'application/json' }
        });

        if (res.ok) {
          contactForm.style.display = 'none';
          const success = document.querySelector('.form-success');
          if (success) success.style.display = 'block';
        } else {
          throw new Error('Form submission failed');
        }
      } catch {
        btn.innerHTML = '<i class="fas fa-exclamation-circle"></i> Error — try emailing directly';
        btn.disabled = false;
        setTimeout(() => {
          btn.innerHTML = originalText;
          btn.disabled = false;
        }, 3000);
      }
    });
  }

  /* ── Smooth hero parallax ── */
  const heroVisual = document.querySelector('.hero-visual');
  if (heroVisual) {
    window.addEventListener('scroll', () => {
      const scrolled = window.scrollY;
      heroVisual.style.transform = `translateY(${scrolled * 0.06}px)`;
    }, { passive: true });
  }

});
