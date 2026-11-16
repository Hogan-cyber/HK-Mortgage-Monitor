# GitHub dashboard delivery

The existing Excel-backed Overview and Mortgage dashboard will be published from the independent public repository Hogan-cyber/HK-Mortgage-Monitor. Its public Pages link is the shareable viewer. Reuse the approved workbook, deterministic Python collector, calculation, and self-contained HTML.

GitHub Actions runs collection daily at 18:00 Hong Kong time (10:00 UTC), and supports manual refresh. Source updates are committed to the repository; a verified snapshot is deployed to Pages in the same workflow. Failed runs retain the previously deployed site and expose their diagnostics. A push can publish an updated saved workbook without fetching sources. No AI service, account for viewers, personal credentials, or other hosting platform is required.

Keep local launchers working. The publisher copies only the verified HTML and matching Excel to the public site. CI checks the workbook-to-HTML match and runs update/backup regression tests on Linux and Windows. Source periods remain visible; existing gaps and signed estimates are preserved. GitHub schedules may be delayed. The new setup does not edit the old property repository or retire previously published sites.
