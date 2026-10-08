# Project 03 Test Plan

## UI / Transform Preview Fix v1.2.0
- Transform & Export preview table has increased usable height.
- Preview table uses readable row height and explicit row foreground/background tags.
- Vertical and horizontal scrolling remain enabled.
- Request Builder is compacted to preserve workspace space on 1366x768 screens.
- GET/DELETE request body remains disabled; POST/PUT enables the body.

## Core tests
- GET demo: HTTP 200
- POST demo: HTTP 201
- PUT demo: HTTP 200
- DELETE demo: HTTP 200
- JSON array to table: 6 rows / 5 columns
- CSV export
- Excel export
- Error status handling
