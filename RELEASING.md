# Releasing MinkExcel

`.github/workflows/publish.yml` publishes to npm when a GitHub release is
published. It checks out the release tag, verifies that the tag matches the
version in `package.json`, builds JavaScript and TypeScript declarations, and
runs unit, tree-shaking, ExcelJS interoperability, Chromium and Web Worker checks
before publishing. The same checks run on pushes and pull requests. The npm package includes `dist/`,
the README, the manifest, the MIT license and the icon. The library remains
dependency-free; TypeScript is installed only in the separate `tools/` project.

## One-time npm setup

The package must exist on npm before configuring
[trusted publishing](https://docs.npmjs.com/trusted-publishers/). Publish the
initial version from a local checkout with an npm account that owns the package:

```sh
bun install --cwd tools --frozen-lockfile --ignore-scripts
bun run build
bun run test
bun run test:treeshaking
bun run test:interop
bun tools/node_modules/@playwright/test/cli.js install chromium
bun run test:browser
npm login
npm publish
```

`npm publish` builds the package through the `prepack` script. Then open the
package's settings on npmjs.com and add a GitHub Actions trusted publisher:

- Organization or user: `accntech`
- Repository: `minkexcel`
- Workflow filename: `publish.yml`
- Environment name: leave blank
- Allowed actions: enable direct publishing with `npm publish`

The workflow authenticates with OIDC, so no npm token or GitHub secret is needed.
Complete the first workflow publish within two days of adding the trusted
publisher; npm expires new configurations that have not yet published.

## Subsequent releases

Update `package.json` to a new, unpublished version and commit it with the
release changes. Publish a GitHub release for that commit using a matching tag,
such as `v0.1.1` for version `0.1.1` (tags without the `v` prefix also work).
The tag must include this workflow and the packaging configuration.

Stable releases publish to npm's `latest` tag. GitHub prereleases and versions
such as `0.2.0-beta.1` publish to `next`. Draft releases and branch pushes do
not publish. For a local package preview, run `npm pack --dry-run` after
installing the tools.
