# Changelog
All notable changes to this project will be documented in this file.  
Format based on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to [Semantic Versioning](https://semver.org/).

---

## [8.25.25] - 2025-08-25
### Added
- Duplicate checking now looks at both **URL (Col B)** and **Article Title (Col C)**.
- New labels supported: `Dupe-Title`, `Dupe-URL`, `Dupe-Both`, and `Original`.

---

## [6.4.25] - 2025-06-04
### Changed
- Dupe search limited to **Article Title only**.
- Script now aborts if Google Doc is not successfully created.

---

## [5.24.25] - 2025-05-24
### Changed
- Adjusted dupe logic so **only later duplicate records are labeled**.  
- Earliest record is marked as `Original`.  
- `createDocFromSidebar` no longer changes Col B or C.

---

## [5.23.25] - 2025-05-23
### Fixed
- Prevented script from modifying **Col C (Article Title)**.  
- Ensured Col C title is used in **Col L** for Google Doc title.

---

## [5.22.25] - 2025-05-22
### Fixed
- Improved duplicate checking logic.  
- Ensured function skips empty rows and writes to the correct sheet.

---

## [5.17.25] - 2025-05-17
### Added
- **Sort and Report** function added.  
- Fixed dupe count reporting for duplicate check.

---

## [5.14.25] - 2025-05-14
### Added
- Expanded duplicate checking to include both **URL and Title**.  
- Added **Sort and Report** function for record counts.

---
