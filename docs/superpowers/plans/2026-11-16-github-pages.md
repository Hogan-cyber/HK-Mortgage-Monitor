# GitHub Pages Implementation Plan

> **For agentic workers:** Use superpowers:executing-plans for this bounded publication change.

**Goal:** Publish the existing dashboard at a stable GitHub Pages URL, updated independently in GitHub Actions.

**Architecture:** Preserve the current workbook and collector at the repository root. Validate and package the generated HTML and workbook into an ignored site directory. A scheduled workflow collects, commits, and deploys; a separate compatibility workflow tests Linux and Windows.

**Tech Stack:** Python 3.11, existing spreadsheet libraries, self-contained Chart.js HTML, GitHub Actions and Pages.

**Spec:** ../specs/2026-11-16-github-pages-design.md

## Global constraints
- Preserve the approved Excel calculation and layout; no imputation.
- Keep the public interface in English, with Overview and Mortgage.
- Daily schedule is 10:00 UTC / 18:00 Asia/Hong_Kong.
- Publish only an allowlist; no credentials, backups, or local runtime caches.

## Review focus
- A mismatched workbook/HTML must fail before deployment: test publisher rejection.
- A moved checkout must work: run existing relocation regression tests.
- A failed refresh must not deploy broken data: workflow failure gate and retained last successful Pages deployment.
- A bot commit must still publish: deployment is in the same workflow.
- Windows differences: run regression suite on windows-latest.

## Tasks
- [ ] Add publisher rejection/round-trip tests and port existing updater regression tests.
- [ ] Implement scripts/build_site.py and the deployment/compatibility workflows, document viewer and maintenance steps.
- [ ] Run local checks and a focused code review, then upload the reviewed file allowlist.
- [ ] Enable Pages Actions, verify deployment and cloud refresh, verify Windows test result, inspect live charts and return the verified link.
