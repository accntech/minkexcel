# Releasing MinkExcel

The `Stage npm release` workflow in `.github/workflows/publish.yml` runs when a
GitHub release is published. It checks out the release tag, sets the package
version from that tag, builds JavaScript and TypeScript declarations, and runs
unit, tree-shaking, ExcelJS interoperability, Chromium, Web Worker and package
checks before staging the package on npm with provenance for your approval.
The same checks run on pushes and pull requests. The npm package includes
`dist/`, the README, the manifest, the MIT license and the icon. TypeScript and
the other development tools stay in the separate `tools/` project.

## npm authentication

For an existing package, configure a GitHub Actions
[trusted publisher](https://docs.npmjs.com/trusted-publishers/) in its npm settings:

- Organization or user: `accntech`
- Repository: `minkexcel`
- Workflow filename: `publish.yml`
- Environment name: leave blank
- Allowed actions: allow `npm stage publish`; direct publishing is unnecessary

With trusted staging configured, the workflow uses GitHub's OIDC token and
needs no npm secret. A direct publish can fail with `OIDC permission denied for
this action` when the publisher permits staging only.

For initial staging before the package exists, or when trusted staging is not
configured, create a granular npm access token with **Read and write (stage
only)** package permission and access to `minkexcel` (or **All packages** for
its first staging). Save it as the repository's Actions secret `NPM_TOKEN`.
Staging does not require 2FA bypass or organization management permissions.
Once trusted staging works, remove the secret. See
[npm's token setup](https://docs.npmjs.com/creating-and-viewing-access-tokens/).

## Stage and approve a release

Publish a GitHub release using a version tag such as `v0.1.1` or
`v0.2.0-beta.1` (tags without the `v` prefix also work). The workflow sets
`package.json` to the tag's version in the runner; no separate version commit
is required. Stable releases use npm's `latest` tag. GitHub prereleases and
versions with a prerelease suffix use `next`. Draft releases and tag or branch
pushes alone do not stage a package.

After the workflow succeeds, open **Staged Packages** in your npm account,
review the `minkexcel` version, and click **Approve**. npm requires interactive
2FA verification before publishing the staged version. For a new package, npm
creates a public `0.0.0-stage` placeholder; the release contents remain staged
until approval. See [npm staged publishing](https://docs.npmjs.com/staged-publishing/).

## Retry an existing release

After this workflow change reaches `main`, open GitHub Actions → **Stage npm
release** → **Run workflow**, select `main`, and enter the existing release tag
(for example, `v0.1.1`). This runs the updated workflow against that tag's source,
so an older tag does not need to contain the staging workflow. Re-running the
original failed run uses its original workflow instead.

Each staged or published version must be unique; staging the same version again
fails. The workflow pins npm 11.15.0 for staging support, with Node.js 24 and
Bun 1.4.0. For a local package preview, install the tools and run
`npm pack --dry-run`.
