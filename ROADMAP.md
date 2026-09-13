# vba-ng Roadmap

One row per version, from the first testing release to the last scheduled item. The detail behind a
row is [docs/build-log.md](docs/build-log.md)'s, and `vbang --version` reports the build in hand.

| Version | Delivers |
|---|---|
| 1.0.0-alpha1 | first general testing release. See [docs/build-log.md](docs/build-log.md), [ARCHITECTURE.md](ARCHITECTURE.md) for background. |
| 1.0.0-alpha2 | fixes from alpha testing; the deferred host items: cross-workbook `Application.Run`, ribbon `onAction`, MSForms argument order, `'@ThreadSafe` |
| 1.0.0-rc1 | alpha feedback closed or decided; the real workbooks end to end, output matching VBA's |
| 1.0.0 | first public release |
| 1.1.0 | VS Code extension: grammar, completions, diagnostics |
| 1.2.0 | UserForms |
| 1.3.0 | 32-bit Excel |
| 1.4.0 | vba-mp, opt-in per project |

After the 1.0.0-rc1 release, GitHub issues and pull requests take over from this table for planned
work.
