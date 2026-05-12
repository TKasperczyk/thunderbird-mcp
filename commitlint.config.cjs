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
//   - header-max-length 150 (default 72) — Thunderbird-MCP scopes get
//     long ("listEvents", "permission-engine", "post-#NN follow-ups"),
//     and the existing refactor branch history has headers in the
//     100-125 char range. 150 covers existing + leaves room before
//     something is genuinely absurd.
//   - body-max-line-length warn-only at 120 — bodies often paste log
//     lines or trace IDs that exceed any reasonable wrap.

module.exports = {
  extends: ["@commitlint/config-conventional"],
  rules: {
    "header-max-length": [2, "always", 150],
    "body-max-line-length": [1, "always", 120],
    // Downgraded from error (2) to warning (1).
    //
    // Default config-conventional bans sentence-case / start-case /
    // pascal-case / upper-case in the subject. That conflicts with
    // legitimate uses of tech acronyms (TB, EPIPE, URL, MCP, HTTP,
    // CI, etc) at subject start: commitlint sees the leading
    // uppercase word and classifies the whole subject as start-case.
    //
    // Project-wide we still WANT subjects to read like sentences
    // (lowercase first word where possible). Warning surfaces the
    // signal without blocking commits where the acronym is the
    // accurate term -- e.g. "fix(bridge): EPIPE-safe stdout writes"
    // is more descriptive than "fix(bridge): epipe-safe stdout writes".
    "subject-case": [
      1,
      "never",
      ["pascal-case", "upper-case"],
    ],
  },
};
