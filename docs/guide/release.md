# Release Checklist

Use this checklist when publishing a new `excel-kit` release.

## Before tagging

- Update the project version in `build.gradle.kts`.
- Move completed `CHANGELOG.md` entries from `Unreleased` into a dated version section.
- Update README installation snippets if the published version changed.
- Run `./gradlew check --no-daemon`.
- Commit the release preparation changes.

## Publish

- Create and push an annotated or lightweight tag named `vX.Y.Z`.
- Watch the tag-triggered workflows:
  - `Release`
  - `Javadoc`
  - `maven-publish.yml`
- Do not rerun Maven publishing after Central upload succeeds. If only visibility
  polling failed, use the manual Maven Central verification workflow instead.

## Verify

- Confirm the GitHub Release exists and is not a draft.
- Confirm Javadocs deployed successfully.
- Confirm Maven Central visibility with:

```bash
scripts/verify-maven-central.sh X.Y.Z
```

After verification, start the next development cycle by restoring an empty
`Unreleased` section in `CHANGELOG.md`.
