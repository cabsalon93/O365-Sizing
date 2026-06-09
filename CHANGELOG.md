# Changelog

All notable changes to this project are documented in this file.

## [4.5-HYCU] - 2026-06-09

### Added
- **Licensed vs unlicensed mailbox breakdown** for Exchange Online. The report
  now reports licensed (user) mailboxes separately from unlicensed ones
  (shared / room / equipment), with counts and storage for each. New keys on the
  returned object: `LicensedMailboxes`, `UnlicensedMailboxes`, `LicensedSizeGB`,
  `UnlicensedSizeGB`.
- `.gitignore`, `LICENSE`, and this `CHANGELOG.md`.

### Changed
- **Growth calculation reworked.** Annual growth is now a proper compound
  (CAGR-style) projection from the first to the last data point over the actual
  number of days, instead of the previous `average daily delta x 2` heuristic,
  which was not mathematically meaningful.
- SharePoint per-site average is now exposed as `SizePerSiteGB` (previously
  mislabeled `SizePerUserGB`).
- All status/progress text now goes to the host stream, so `-OutputObject`
  returns a clean sizing object with no banner strings mixed in.

### Fixed
- `-OutputObject` is now a proper `[switch]` parameter.
- Single-mailbox archive runs no longer fail on `.Count` (results are coerced to
  arrays).
- Progress bars in `Start-SleepWithProgress` now display (previously suppressed
  by the global `$ProgressPreference`).
- Azure AD group names containing an apostrophe no longer break the Graph OData
  filter.
- Microsoft Graph session is now disconnected on normal completion.

## [4.4-HYCU] - 2026-01

- Complete HYCU branding with deep purple color scheme.
- Modern, responsive HTML report design.
- Optimized code structure and error handling.
- Enhanced progress indicators.
- Comprehensive inline documentation.
