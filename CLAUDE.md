# Project Instructions for AI Agents

This file provides instructions and context for AI coding agents working on this project.

<!-- BEGIN BEADS INTEGRATION v:1 profile:minimal hash:7510c1e2 -->
## Beads Issue Tracker

This project uses **bd (beads)** for issue tracking. Run `bd prime` to see full workflow context and commands.

### Quick Reference

```bash
bd ready              # Find available work
bd show <id>          # View issue details
bd update <id> --claim  # Claim work
bd close <id>         # Complete work
```

### Rules

- Use `bd` for ALL task tracking — do NOT use TodoWrite, TaskCreate, or markdown TODO lists
- Run `bd prime` for detailed command reference and session close protocol
- Use `bd remember` for persistent knowledge — do NOT use MEMORY.md files

**Architecture in one line:** issues live in a local Dolt DB; sync uses `refs/dolt/data` on your git remote; `.beads/issues.jsonl` is a passive export. See https://github.com/gastownhall/beads/blob/main/docs/SYNC_CONCEPTS.md for details and anti-patterns.

## Session Completion

**When ending a work session**, you MUST complete ALL steps below. Work is NOT complete until `git push` succeeds.

**MANDATORY WORKFLOW:**

1. **File issues for remaining work** - Create issues for anything that needs follow-up
2. **Run quality gates** (if code changed) - Tests, linters, builds
3. **Update issue status** - Close finished work, update in-progress items
4. **PUSH TO REMOTE** - This is MANDATORY:
   ```bash
   git pull --rebase
   git push
   git status  # MUST show "up to date with origin"
   ```
5. **Clean up** - Clear stashes, prune remote branches
6. **Verify** - All changes committed AND pushed
7. **Hand off** - Provide context for next session

**CRITICAL RULES:**
- Work is NOT complete until `git push` succeeds
- NEVER stop before pushing - that leaves work stranded locally
- NEVER say "ready to push when you are" - YOU must push
- If push fails, resolve and retry until it succeeds
<!-- END BEADS INTEGRATION -->


## Build & Test

This project is **pnpm-only** (`only-allow pnpm`). Never use `npm` or `npx`.

```bash
pnpm install
pnpm test            # deterministic node suites — must pass outright
pnpm run test:live-sit   # live smoke tests against the SIT deployment
```

## Deploying

**The deploy pipeline lives in the shared `gas-deploy` package** (GAS-Core
`packages/gas-deploy/`, pinned by tag in `package.json`), not in this repo.
`tools/manage-deployments.js` here is pure config — the three targets, the stamper, and the
ordered post-deploy hooks — and `tools/callWebapp.js` is a thin wrapper over the package's one
HTTP client. For deploy *internals* (auth, deployment-ID resolution, stamping, verification,
summary, hook semantics) read that package's README; changing behaviour means changing the
package and cutting a new `gas-deploy-vX.Y.Z` tag, not editing these two files. Background:
`GAS-Core/best-practices/gas-deployment/RECOMMENDATION.md`.

```bash
pnpm run deploy:sit    # bump build + stamp SIT  + push to sitScriptId
pnpm run deploy:prod   # bump patch + stamp PROD + push to prodScriptId
pnpm run deploy:nuuc   # bump patch + stamp NUUC + push to nuucScriptId
```

### Deploy verification (`cmd=version`)

`clasp deploy` exiting 0 only proves a version was *created*, not that the `/exec` URL is
serving it. Every deploy therefore ends with `assertDeployedVersion`
(gas-deploy's `lib/verify.js`), which polls the deployment's `?cmd=version` route until the
webapp itself reports the exact version **and** target just stamped:

```jsonc
// GET or POST ?cmd=version  →
{ "ok": true, "version": "0.1.6.2", "versionDate": "2026-08-22T01:19:41.747Z",
  "target": "SIT", "deploymentId": "AKfycbwRGVyw…" }
```

- **No secret required.** `cmd=version` is routed ahead of the `cmd=admin` branch in both
  `doGet` and `doPost`, so it answers on an `ANYONE_ANONYMOUS` deployment and before
  `ADMIN_SHARED_SECRET` is ever bootstrapped. Do not move it behind the secret gate.
- The **target** check is what catches a deploy landing in the wrong environment — SIT, PROD and
  NUUC share one version counter, so a version match alone would not.
- On mismatch the deploy fails (non-zero exit) with expected-vs-actual, but still prints the
  summary so you can see what *is* deployed.
- The summary's version row is the **server-confirmed** value, not the locally stamped one.

Query it by hand with:

```bash
node -e "require('./tools/callWebapp.js').post('https://script.google.com/macros/s/'+require('./local.settings.json').sitDeploymentId+'/exec?cmd=version',{action:'version'}).then(r=>console.log(r))"
```

The route is `script/WebApp.js`'s `handleVersionRequest_`, reading `script/version.js`'s stamped
`APP_VERSION` / `APP_VERSION_DATE` / `APP_DEPLOY_TARGET`. It is GAS-side, per-project code by
necessity — only this project knows where its stamper wrote — so it stays here while the Node
side of the verification lives in the package.

`node tools/manage-deployments.js --summary --env sit` prints the same summary for what is
*currently* deployed without deploying anything, and flags any divergence between the live
version and the local `script/version.js`.

## Architecture Overview

_Add a brief overview of your project architecture_

## Conventions & Patterns

_Add your project-specific conventions here_
