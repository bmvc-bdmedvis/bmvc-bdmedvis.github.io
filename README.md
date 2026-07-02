# Biomarker-Driven Medical Vision Intelligence

Static website for the BMVC 2026 workshop **Biomarker-Driven Medical Vision Intelligence**.

The site is a no-build static GitHub Pages site with:

- `index.html` for content and structure
- `styles.css` for the flat responsive visual system
- `script.js` for scroll progress, active navigation, and mobile menu behavior
- `.github/workflows/pages.yml` for GitHub Pages deployment through Actions

## Local preview

Open `index.html` directly in a browser, or run a lightweight local server:

```bash
python3 -m http.server 8000
```

Then visit `http://localhost:8000`.

## Publish on GitHub Pages

This site has no build step. To publish it:

1. Create or open a GitHub repository for the workshop site.
2. Push this directory to the repository's `main` branch.
3. In GitHub, go to `Settings -> Pages`.
4. Set the source to either `Deploy from a branch` with `main` and `/root`, or use the included GitHub Actions workflow.

If using the GitHub CLI for a new public repository:

```bash
git init -b main
git add .
git commit -m "Create BMVC 2026 workshop website"
gh repo create bmvc-2026-bdmedicalvision --public --source=. --remote=origin --push
```

After GitHub Pages is enabled, the site will be available at the repository's Pages URL.

## QA performed

- Desktop browser check at `1440 x 1000`
- Mobile browser check at `390 x 844`
- Image load check for all page assets
- Internal anchor check for navigation links
- Mobile menu open/close and Program anchor jump check
- Console warning/error check
- Horizontal overflow check

## Asset notes

- The BMVC 2026 logo and Lancaster Town Hall venue image are adapted from the public BMVC 2026 website.
- The medical vision hero image was generated for this workshop website and compressed for web delivery.
