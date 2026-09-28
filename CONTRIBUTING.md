# Contributing

Keep changes focused and follow the surrounding code. Search for existing helpers before adding one. Describe the problem, the resulting behavior, and how you checked it in your pull request.

## Build

Use Node.js 22 or newer for development. On Linux/macOS, run `bash scripts/build.sh` (requires `zip`). Alternatively, including on Windows, run `node scripts/build-xpi.cjs`. Both produce `dist/thunderbird-mcp.xpi`. Install it from Thunderbird's Add-ons manager and restart Thunderbird to load the changes.

## Test and lint

The Node test suite needs no installed dependencies. **When a live Thunderbird is running, tests must use an isolated `TMPDIR`** so their connection-file fixtures cannot disturb it. From the repository root, use a dedicated test directory:

```sh
mkdir -p /tmp/thunderbird-mcp-tests && TMPDIR=/tmp/thunderbird-mcp-tests npm test
```

Keep that directory separate from Thunderbird's temporary directory. For linting, install only the development dependencies without lifecycle scripts, then run ESLint:

```sh
npm install --no-package-lock --ignore-scripts
npx eslint .
```

For Experiment API changes, follow the existing marker + `node:vm` tests that execute production code with Thunderbird services stubbed. Note any manual Thunderbird checks and the version/channel used in your PR.

## Bugs and security

Use the bug report form for reproducible failures; remove tokens and private mail data from logs. Report security issues through [GitHub private vulnerability reporting](https://github.com/TKasperczyk/thunderbird-mcp/security/advisories/new), as described in [SECURITY.md](SECURITY.md), rather than a public issue.
