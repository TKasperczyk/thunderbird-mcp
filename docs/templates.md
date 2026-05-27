# Compose templates

The `listTemplates` and `renderTemplate` MCP tools read user-authored
templates from `<Thunderbird profile>/thunderbird-mcp/templates/`. The
directory does not exist by default — create it and drop `*.md` files
inside. The format is Jekyll-style YAML frontmatter followed by the
message body.

## Format

```
---
name: outreach-v1-approach
description: First-touch outreach. Minimal -- just approach + offer report.
subject: Security finding for {{program}}
isHtml: false
vars: [contact_name, program, my_name]
---
Hi {{contact_name}},

I'm {{my_name}}. I'm reaching out about a security finding affecting {{program}}.
...
```

### Frontmatter keys

| key           | type            | required | meaning                                                   |
| ---           | ---             | ---      | ---                                                       |
| `name`        | string          | yes      | Short ID. What you pass to `renderTemplate({ name, vars })`. Filename without `.md` is used if omitted. |
| `description` | string          | no       | One-line summary; shown in `listTemplates`.               |
| `subject`     | string          | no       | Subject line. Supports `{{var}}` substitution.            |
| `isHtml`      | boolean         | no       | Treat body as HTML. Default `false`.                      |
| `vars`        | array of string | no       | Names of `{{var}}` placeholders the caller must supply. `renderTemplate` errors if any are missing. |

### Variable substitution

Placeholders look like `{{name}}` and are replaced by the matching key
from the `vars` object passed to `renderTemplate`. Unknown placeholders
are left as literals so you notice rather than silently shipping an
empty value.

## Usage from an MCP client

```jsonc
// 1. Discover what's available
{ "method": "tools/call", "params": { "name": "listTemplates" } }

// 2. Render
{
  "method": "tools/call",
  "params": {
    "name": "renderTemplate",
    "arguments": {
      "name": "outreach-v1-approach",
      "vars": { "contact_name": "Alex", "program": "ExampleCorp", "my_name": "Jordan" }
    }
  }
}
// Returns { name, subject, body, isHtml, file }

// 3. Feed the rendered output into sendMail
{
  "method": "tools/call",
  "params": {
    "name": "sendMail",
    "arguments": {
      "to": "security@example.com",
      "subject": "<rendered subject>",
      "body": "<rendered body>",
      "isHtml": false,
      "skipReview": false,
      "idempotencyKey": "outreach-v1-examplecorp-2026-01"
    }
  }
}
```

Using `idempotencyKey` lets you re-run the same agent loop after a
crash without double-sending to the same target.

## Why this format

- One file per template, plain text. Editable in any tool, diff-able,
  versionable in your own private dotfiles repo.
- Lives in `ProfD`, not in the extension bundle — your templates are
  not shipped publicly when the extension is updated, and they survive
  an `.xpi` reinstall.
- No template engine dependency. The substitution is `{{name}}` only;
  no loops, no conditionals, no inline code. Keeps the surface
  predictable for an LLM.
- Variable list is declared up-front so `renderTemplate` can fail
  loudly when the caller forgot a placeholder, instead of producing
  an output with a literal `{{name}}` in it.

An example template ships in `docs/templates-example.md`.
