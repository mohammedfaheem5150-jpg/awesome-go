---
name: go-entry-review
description: Review a proposed awesome-go entry against the contribution guidelines. Use when the user asks to check a package addition, a PR entry, or a new link before submitting it to awesome-go.
---

# Go Entry Review

Review a proposed entry for the awesome-go list against the repository's
contribution guidelines in `CONTRIBUTING.md`.

## Steps

1. Identify the package being proposed: its repository URL, the target
   category in `README.md`, and the one-line description.
2. Check the entry formatting rules:
   - The entry must use the format `[package](https://link) - Description.`
   - The description starts with a capital letter and ends with a period.
   - The entry is placed in the correct category, in alphabetical order.
   - The link is not already present anywhere in `README.md`.
3. Check the quality standards from `CONTRIBUTING.md`:
   - The project has at least 5 months of history since its first commit.
   - The project has an approved open source license.
   - The project functions as documented and expected.
   - The project's documentation links to a Go Report Card and a code
     coverage report for the latest release or ongoing development.
4. Report the result as a checklist, marking each requirement as passing,
   failing, or impossible to verify from the available context. For any
   failure, quote the relevant guideline and explain what needs to change.

## Notes

- Do not approve an entry outright; the final decision belongs to the
  maintainers. Your job is to surface problems before review.
- If the target category does not exist in `README.md`, suggest the closest
  existing category instead of inventing a new one.
