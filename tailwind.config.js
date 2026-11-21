// Builds assets/app.css from the classes used in index.html (replaces the Tailwind CDN script,
// which compiled the styles in every visitor's browser and was the slowest part of opening the page).
// Vercel re-runs this on every deploy (see vercel.json), so new classes added to index.html still work.
module.exports = {
  content: ['./index.html'],
  theme: { extend: {} },
  plugins: []
};
