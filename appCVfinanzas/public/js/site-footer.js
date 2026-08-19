(function () {
  class CvFooter extends HTMLElement {
    connectedCallback() {
      this.innerHTML = `
        <footer class="cv-site-footer">
          <div class="cv-site-footer__inner">
            <a class="cv-site-footer__brand" href="https://cvfinanzas.com" target="_blank" rel="noreferrer">
              <svg class="cv-site-footer__mark" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M7.2 7.1 10.9 17l6-14" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"></path>
                <path d="M5.2 4.8h.1" stroke="currentColor" stroke-width="3" stroke-linecap="round"></path>
              </svg>
              <span>cvfinanzas.com</span>
            </a>

            <nav class="cv-site-footer__social" aria-label="Redes sociales">
              <a href="https://instagram.com/carlosvalerincr" target="_blank" rel="noreferrer">
                <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                  <rect x="3" y="3" width="18" height="18" rx="5" stroke="currentColor" stroke-width="2"></rect>
                  <circle cx="12" cy="12" r="4" stroke="currentColor" stroke-width="2"></circle>
                  <path d="M17.5 6.5h.01" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"></path>
                </svg>
                <span>@carlosvalerincr</span>
              </a>

              <a href="https://tiktok.com/@carlosvalerincr" target="_blank" rel="noreferrer">
                <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                  <path d="M16.7 3c.3 2.2 1.6 3.9 4.1 4.2v3.1c-1.5 0-2.9-.4-4.1-1.2v6.1c0 3.5-2.6 5.8-5.9 5.8A5.6 5.6 0 0 1 5 15.4c0-3.3 2.7-5.7 5.9-5.7.4 0 .8 0 1.1.1v3.3c-.3-.1-.7-.2-1.1-.2-1.3 0-2.4 1-2.4 2.4 0 1.5 1.1 2.5 2.4 2.5s2.4-.9 2.4-2.5V3h3.4Z"></path>
                </svg>
                <span>@carlosvalerincr</span>
              </a>
            </nav>

            <p class="cv-site-footer__data">v.1.0 2026</p>
          </div>
        </footer>
      `;
    }
  }

  if (!customElements.get('cv-footer')) {
    customElements.define('cv-footer', CvFooter);
  }
}());
