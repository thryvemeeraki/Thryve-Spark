# Thryve Spark Form

Static deployment package for the Thryve Spark assessment form.

## Files

- `index.html` - final form file, generated in a later step
- `Code.gs` - Google Apps Script backend, generated in a later step
- `assets/logo.png` - Spark logo
- `assets/palette.png` - supplied colour palette reference
- `.nojekyll` - GitHub Pages compatibility file
- `netlify.toml` - Netlify static hosting config

## Deployment Flow

1. Test locally.
2. Upload to GitHub Pages for testing.
3. Deploy final version to Netlify.
4. Generate student links only from the final hosted URL.
