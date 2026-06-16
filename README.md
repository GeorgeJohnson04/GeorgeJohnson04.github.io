# George Johnson — Professional Portfolio

A clean, responsive, multi-page portfolio site built with plain **HTML, CSS, and JavaScript** —
no frameworks, no build step. Designed to be hosted free on **GitHub Pages**.

## Pages

| Page | File | What's on it |
|---|---|---|
| Home / About | `index.html` | Intro, summary, core skills, quick links |
| Work Experience | `experience.html` | Timeline of all six analyst/research roles |
| Projects | `projects.html` | 14 projects, filterable by category, with live thumbnails |
| Certifications | `certifications.html` | UH Quantitative Economics + verifiable DataCamp / CITI / Anthropic certs |
| Résumé | `resume.html` | Embeds a PDF résumé (and a live, printable HTML fallback) |

## Project structure

```
.
├── index.html
├── experience.html
├── projects.html
├── certifications.html
├── resume.html
├── assets/
│   ├── css/style.css        # all styling
│   ├── js/main.js           # nav, filtering, image fallbacks, reveal
│   └── img/
│       ├── projects/        # project thumbnails (real charts/screenshots from your repos)
│       └── logos/           # company logos (drop real ones here to replace placeholders)
├── Certifications/          # the PDF certificates linked from certifications.html
├── .nojekyll                # tells GitHub Pages to serve files as-is
└── .gitignore
```

## How to deploy on GitHub Pages

You already have a repo named **`GeorgeJohnson04.github.io`** — that's a *user site*, so whatever
lands on its default branch is published at **https://georgejohnson04.github.io**.

1. Copy the contents of this folder into that repository (keep the folder structure).
2. Commit and push to the `main` branch.
3. In the repo: **Settings → Pages → Build and deployment → Source: Deploy from a branch**,
   branch **`main`**, folder **`/ (root)`**. Save.
4. Wait ~1 minute, then visit **https://georgejohnson04.github.io**.

> Prefer a project site instead? Put these files in any repo, enable Pages the same way,
> and the URL becomes `https://georgejohnson04.github.io/<repo-name>/`.

## Notes

- Project thumbnails/descriptions, company logos, and work samples were pulled from your public
  repos, the web, and the files you provided.
- The only logo still on a monogram placeholder is **Southern Esports** — save the fox image as
  `assets/img/logos/sec.png` to use it.
- Contact links point to `george.johnson022004@gmail.com`.
- `Johnson_George_Transcripts.pdf` is git-ignored so it won't be published.
