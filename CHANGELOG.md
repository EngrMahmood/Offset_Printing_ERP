# Changelog

All notable changes to Offset Printing ERP.

## [2026.04.22.1] - 2026-04-22
### Added
- Merged the `planning` branch into `main` for server deployment.
- Added manual PO intake entry for direct PO creation without PDF upload.
- Added PO review manual entry support for manual PO number edits and PO line additions.
- Added planning job archive and hold workflows, plus archived jobs management page.
- Added a `Manual PO Entry` option from the PO upload and intake pages.
- Updated ERP software version to `2026.04.22.1`.

### Changed
- Redirect PO PDF upload success to the PO review page instead of PO intake queue.
- Updated planning workflow to track merges, archive/restore, and change tracking more clearly.

### Notes
- Use `main` branch for server pulls after this merge.
