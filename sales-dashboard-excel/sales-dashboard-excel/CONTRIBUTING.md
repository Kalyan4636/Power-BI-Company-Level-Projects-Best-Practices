# Contributing

1. Fork, create a branch: `git checkout -b feat/short-description`.
2. `make install && make test` must pass before you open a PR.
3. Every new metric or filter needs a test in `tests/` that compares the workbook result to a pandas calculation.
4. If you change the layout, regenerate the images with `make png` and commit them.
5. Commit style: [Conventional Commits](https://www.conventionalcommits.org/) (`feat:`, `fix:`, `docs:`, `test:`, `chore:`).

## Ground rules for the workbook
- No hard-coded numbers on the Dashboard sheet: every value is a formula (enforced by `test_no_hardcoded_kpis`).
- Calculations live on `Analysis`; the `Dashboard` sheet only references it.
- Add a dimension by (a) adding a `Flag excl. <Dim>` column on `Data`, (b) adding a `block(...)` on `Analysis`, (c) wiring a chart in `build_dashboard.py`.
