# Changelog

All notable changes to this project will be documented in this file.

## [v5.15.7] - 2026-01-08
### Fixed
- **PDF Merge Compatibility**: Fixed `AttributeError: 'PageObject' object has no attribute 'indirect_ref'` by adding fallback support for older `pypdf` versions (using `indirect_reference` or skipping links safely).

## [v5.15.6] - 2026-01-08
### Fixed
- **Merge Output**: Added post-write validation and printed output paths to avoid “silent no-output” situations.
- **UX**: After merge, opens the `_MERGED` directory when outputs exist.

## [v5.15.5] - 2026-01-08
### Changed
- **PDF Index Hotspots**: Made all index entries clickable across multiple index pages.

## [v5.15.4] - 2026-01-08
### Fixed
- **Merge Logic**: Completely removed duplicated legacy merge logic that caused incorrect Excel output and double processing.
- **Excel Listing**: Verified that the merge list now correctly outputs one row per component file.

## [v5.15.3] - 2026-01-08
### Fixed
- **Runtime Error**: Fixed missing `traceback` import causing crash when logging merge errors.

## [v5.15.2] - 2026-01-08
### Added
- **PDF Index Hotspots**: The generated index page now includes clickable links jumping to the corresponding document start page.
- **Enhanced Excel Listing**: The merge list Excel now records one row per component file (Output Filename, Source Filename) for easier tracking.

### Changed
- **Merge Logic**: Refactored merge task generation for better stability.
- **Index Layout**: Improved filename truncation on index page to prevent layout overflow.

## [v5.15.1] - 2026-01-08
### Added
- Date-based file filtering (before/after specific date).
- Merge options: Index page generation and Excel file listing.
- UI controls for new features.
