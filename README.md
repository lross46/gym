# Gym Log

A phone web app for following a training programme: workout mode with timers, a log of weights and notes, and progress charts.

This repository holds only the app. No training data is stored here: the app keeps the log on the phone and syncs it to a separate private repository using an access key that stays on the device.

- `site/` is the app. Pushing to `main` copies it to the `gh-pages` branch, which GitHub Pages serves.
- `site/import.js` reads programme files (.doc / .docx) on the device.
- `site/vendor/xlsx.mini.min.js` is SheetJS (Apache 2.0), used for Excel export and reading Word files.
