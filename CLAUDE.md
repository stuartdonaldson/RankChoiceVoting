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
pnpm run deploy:sit    # bump build + stamp SIT  + push to envs.sit.scriptId
pnpm run deploy:prod   # bump patch + stamp PROD + push to envs.prod.scriptId
pnpm run deploy:nuuc   # bump patch + stamp NUUC + push to envs.nuuc.scriptId
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

`pnpm run deployments` (`tools/show-deployments.js`) lists every env at once — spreadsheet, admin
page and webapp URLs, plus the version the backend (`cmd=version`) and the static page
(`version.json`) each report, flagging a wrong target or backend/page skew. It is read-only and
never calls clasp; the webapp URL comes from the `*DeploymentId` cache in `local.settings.json`.

## Configuration lives in two files

GAS-Core `adr/0004`. Know which half a fact belongs in before you add one:

| | file | holds |
|---|---|---|
| **project truth** | `gas-project.json`, **committed** | Per-env identifiers of external resources: `scriptId`, `sheetId`, and the static host (`staticHosts.<host>.repo` / `.pagesUrl`, plus each env's `static.dest`). Identical for every developer, so a typo is caught in review once instead of on every machine. |
| **machine truth & secrets** | `local.settings.json`, **gitignored** | `claspAuth` / `nuucAuth`, the three `*AdminSecret`s, the Axiom tokens, `staticRepoPath` / `nuucStaticRepoPath` (this machine's checkouts of the host repos), and the `*DeploymentId` **cache** the deploy writes back. |

None of the IDs in the committed half is a secret — a `scriptId` appears in every Apps Script
editor URL. Access is controlled by Drive sharing, not by their obscurity.

**Env *declarations* stay in JS**, not JSON: labels, counters, emoji and the post-deploy hook order
are `tools/manage-deployments.js`'s; the themes and font links are `tools/static-pages.js`'s. They
were always committed and reviewable, so moving them would add a second declaration mechanism
without retiring the first.

A disagreement between the two files is never silent: an env declared in `gas-project.json` with no
matching secret throws before anything shells out, and a fact left behind in both files takes the
committed value and warns which stale key to delete.

## The static front end

The UI is a static page on GitHub Pages, not `HtmlService` — GAS serves only `?cmd=api` JSON plus a
one-tap redirect (`script/ApiBridge.js`). **The build/publish/verify pipeline lives in the shared
`gas-static` package**, pinned by tag; `tools/static-pages.js` here is pure config. It replaced
`tools/build-static-pages.js` and `tools/publish-static-pages.js`, both deleted.

It is never run on its own: `manage-deployments.js` chains **build → publish →
assertPublishedBuild** as *required* post-deploy hooks, so the page and the backend it calls ship as
one release and a disagreement fails the deploy.

Three env keys (`sit`, `prod`, `nuuc`) across **two** host repos — `f3go30/static-pages` at
`ballot/{sit,prod}`, `nuuc-it/Static` at `pub/ballot`. Both carry a `PUBLISHERS.md` ownership map,
and a publish to a folder registered to another project is refused (GAS-Core `adr/0003`).
Registering a new folder is a reviewed edit **in the host repo**, first.

Two things worth knowing before touching it:

- **The backend URL comes from `BUILD_INFO`, never from a deployment ID.** `script/version.js`
  carries a generated `BUILD_INFO` literal whose `webappUrl` is the `/exec` URL of the deployment
  *this* deploy just landed in — which is why `resolveBeforeStamp: true` is set. Building an env
  whose declared `deployTarget` disagrees with the stamped `env` throws before writing anything.
  GAS-Core `adr/0001`; the old build constructed that URL from `local.settings.json` and reconciled
  it against nothing.
- **`BUILD_INFO.staticUrl` is how GAS-side code knows the page URL.** `_staticPagesBaseUrl_` reads
  it rather than hardcoding three URLs as it used to. A GAS script file cannot `require()`
  `gas-project.json`, so the composed URL is stamped in at deploy time.

Adding a page-level build-time value means adding a token to `placeholders` in
`tools/static-pages.js` — not changing the package. A missing placeholder is a hard failure that
writes nothing.

## Architecture Overview

_Add a brief overview of your project architecture_

## Conventions & Patterns

_Add your project-specific conventions here_
