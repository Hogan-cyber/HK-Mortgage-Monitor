# Hong Kong Mortgage Monitor

[**Open the dashboard**](https://hogan-cyber.github.io/HK-Mortgage-Monitor/) · [Download the Excel workbook](https://hogan-cyber.github.io/HK-Mortgage-Monitor/hk_mortgage_data.xlsx)

Share the dashboard link with colleagues. Viewing requires only a browser: no account, installation, or local files. The English interface has two sections:

| Section | Content |
| --- | --- |
| Overview | HSBC prime rate, daily HIBOR by tenor, and M3 money supply |
| Mortgage | Centaline registrations, HKMA bank mortgage activity, Land Registry documents, and the workbook's non-bank residual estimate |

## Automatic updates

GitHub Actions retrieves source releases **daily at 18:00 Hong Kong time (10:00 UTC)**, updates the workbook, regenerates the dashboard, verifies every embedded data group against the saved workbook, and publishes to GitHub Pages. Your computer does not need to be running. The collection and publication code runs independently without an AI service.

The schedule is a target start time, not a guarantee of publication at exactly 18:00. GitHub can delay scheduled runs. Observation dates shown in the charts describe the data, while the snapshot time records when the dashboard was generated. A Git checkout does not preserve the original file save time. Missing source releases remain missing; historical values and negative estimates are preserved.

For an immediate update, open [Actions → Update data and publish dashboard](https://github.com/Hogan-cyber/HK-Mortgage-Monitor/actions/workflows/publish.yml), choose **Run workflow**, leave **Fetch new source data** checked, and run it on `main`. Only maintainers need to sign in. Uncheck it to publish an uploaded, saved Excel workbook without retrieving sources.

An unsuccessful run retains the previously published website. Check `update_status.txt`, `collection_report.json`, and the run's diagnostic artifact for missing observations or source errors. GitHub's workflow notifications follow each maintainer's notification settings. This repository commits refreshed snapshots as part of successful runs; GitHub may disable schedules after 60 days without repository activity.

## Excel and the estimate

`hk_mortgage_data.xlsx` is the baseline. Updates merge source observations into that workbook and retain its sheets, formatting, and established calculation. The dashboard reads saved workbook values.

- Estimated bank purchase drawdowns for month *t* = (primary + secondary approval count) / total approval count × total new drawdown count.
- Non-bank residual for registration month *t* = Centaline completed + off-plan registrations in *t* − estimated bank purchase drawdowns in the preceding calendar month.

This is the workbook's existing **count estimate**, not an independently observed count of non-bank lenders. Source coverage and timing differ. It does not establish a non-bank lending amount.

| Workbook data | Source |
| --- | --- |
| Completed/off-plan and housing-type mortgage registrations | Centaline Mortgage API and dated reports/attachments |
| Bank mortgage applications, approvals, drawdowns and balances | HKMA Monthly Statistical Bulletin T0307 |
| Mortgage and release document registrations | Land Registry monthly statistics |
| M3 | HKMA T020201 |
| HSBC prime | HKMA historical table and HSBC current effective rate |
| Daily HIBOR | HKMA historical table and Hong Kong Association of Banks daily fixing |

Exact Centaline source references are retained in `source_evidence.json`. The collector preserves distinct series definitions and dates; bank statistics and registration documents are not presented as identical measures.

## Local use

Download this repository with **Code → Download ZIP**, then extract the entire folder. Double-click `HK_Mortgage_Dashboard.html` to view the included snapshot offline.

For updates, install Python 3.10 or newer (Anaconda is also supported), connect to the internet, and close Excel. Double-click **Update Data.bat** on Windows or **Update Data.command** on macOS. Missing or unsupported Python libraries are prepared automatically in a local environment. The script creates `hk_mortgage_data.previous.xlsx` before replacing the workbook.

```bash
python update_mortgage_data.py                     # Sources → Excel → HTML
python update_mortgage_data.py --dashboard-only    # Saved Excel → HTML
python update_mortgage_data.py --recalculate-only  # Recalculate existing formula → HTML
python update_mortgage_data.py --no-dashboard      # Update Excel only
```

Use `py -3` on Windows or `python3` on macOS/Linux if `python` is unavailable. Local edits do not automatically upload: commit the saved workbook to `main` to publish them, or use the GitHub workflow for cloud collection.

## Maintenance

| File | Purpose |
| --- | --- |
| `update_mortgage_data.py` | Source collection, existing-workbook updates, formula caching, dashboard generation |
| `HK_Mortgage_Dashboard.html` | Self-contained interface and embedded workbook snapshot |
| `scripts/build_site.py` | Verify workbook/HTML equality and prepare only approved public files |
| `.github/workflows/publish.yml` | Daily/manual updates, snapshot commits, Pages deployment |
| `.github/workflows/compatibility.yml` | Python and workbook checks on Windows and Linux |
| `tests/` | Relocation, failure preservation, formula, dependency and publication checks |

The Pages publishing source is **GitHub Actions**. Only the verified `site/` output is deployed; credentials, local environments and backups are excluded. Run `python -m unittest discover -s tests -v` before changing collection or publication code. Keep the Chart.js license with redistributed dashboard files.
