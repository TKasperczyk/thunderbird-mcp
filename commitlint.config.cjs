// Conventional Commits enforcement.
//
// Why this exists:
//   - PRs get split by scope. A clean `type(scope): subject` header is what
//     makes that splitting cheap — grep by scope, cherry-pick by type.
//   - Dependabot is configured (in .github/dependabot.yml) to write
//     `chore(deps):` / `ci(actions):` prefixes. Commitlint enforces the
//     contributor side of the same contract.
//
// Rules left at defaults from @commitlint/config-conventional:
//   - type-enum: feat | fix | docs | style | refactor | perf | test | chore | ci | build | revert
//   - subject-case: lower / sentence-case / start-case / pascal-case / upper-case
//   - scope-case: lower-case
//
// Local overrides:
//   - header-max-length 100 (default 72) — Thunderbird-MCP scopes get long
//     ("listEvents", "permission-engine"), 72 is too tight in practice.
//   - body-max-line-length warn-only at 120 — bodies often paste log lines
//     or trace IDs that exceed any reasonable wrap.

module.exports = {
  extends: ["@commitlint/config-conventional"],
  rules: {
    "header-max-length": [2, "always", 100],
    "body-max-line-length": [1, "always", 120],
  },
};
