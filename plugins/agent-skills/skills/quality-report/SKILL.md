---
name: quality-report
description: Generate a maintenance quality report for a Go project listed in awesome-go. Use when the user asks whether a listed project still meets the quality standards or should be flagged for removal.
---

# Quality Report

Assess whether a project already listed in awesome-go still meets the
ongoing quality standards described in `CONTRIBUTING.md`.

## Steps

1. Locate the project's entry in `README.md` and note its category, link,
   and description.
2. Evaluate the ongoing maintenance criteria:
   - Development is ongoing, with an official release at least once a year,
     or the project is mature and stable with no bug reports older than
     6 months in its issue tracker.
   - The documentation still links to current quality reports (Go Report
     Card and code coverage) for the most recent release or ongoing
     development.
   - The project remains open source under an approved license.
   - The project is compatible with a Go version released within the last
     year.
3. Produce a report with one section per criterion, stating whether it is
   met, unmet, or unverifiable from the available context, with supporting
   evidence for each conclusion.
4. If two or more criteria are unmet, recommend following the removal
   process from `CONTRIBUTING.md`: open an issue at the project's repository
   notifying the author, and prepare a removal PR that describes which
   criteria are not being met.

## Notes

- Be conservative: prefer marking a criterion as unverifiable over guessing.
- Never open removal issues or PRs yourself; present the prepared text to
  the user for review first.
