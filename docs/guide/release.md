# Release Checklist

Use this checklist when publishing a new `excel-kit` release.

## Before tagging

- Update the project version in `build.gradle.kts`.
- Move completed `CHANGELOG.md` entries from `Unreleased` into a dated version section.
- Update README installation snippets if the published version changed.
- Run `./gradlew build :kit:poiCompatibilityTest --no-daemon`.
- Run `scripts/verify-consumers.sh` to publish unsigned artifacts to the isolated
  `build/consumer-repository` and test Excel-only, CSV-only, optional Validation,
  and Spring consumers. Both Gradle module metadata and Maven POM resolution are
  exercised; neither mode substitutes local project dependencies.
- CI must pass on actual test JVMs 17, 21, and 25. Compilation uses JDK 21 with
  `--release 17`; `-PtestJavaVersion=17` (or 21/25) selects the test JVM explicitly.
- Run `./gradlew verifyReleaseVersion -PreleaseTag=vX.Y.Z --no-daemon`.
- Commit the release preparation changes.

## Publish

- Create and push an annotated or lightweight tag named `vX.Y.Z`.
- Watch the tag-triggered workflows:
  - `Maven Central release` (calls `Release` after publishing and consumer verification)
  - `Javadoc`
- The publishing workflow rejects tag/version mismatches before loading signing
  credentials. Both modules are published in one Gradle invocation so the shared
  Central publishing service uploads a single deployment bundle.
- GitHub Release creation waits for both artifacts to become visible and for
  independent consumers to compile and run against the released artifacts.
- Do not rerun Maven publishing after Central upload succeeds. If only visibility
  polling failed, use the manual Maven Central verification workflow instead.
  If GitHub Release creation or the post-publication consumer check fails, rerun
  only the `Release` workflow against the existing release tag (for example,
  `gh workflow run release.yml --ref vX.Y.Z`), without uploading artifacts again.

## Verify

- Confirm the GitHub Release exists and is not a draft.
- Confirm Javadocs deployed successfully.
- Confirm Maven Central visibility for both published artifacts with:

```bash
scripts/verify-maven-central.sh X.Y.Z
scripts/verify-maven-central.sh X.Y.Z --artifact-id excel-kit-spring
scripts/verify-consumers.sh X.Y.Z https://repo1.maven.org/maven2
```

After verification, start the next development cycle by restoring an empty
`Unreleased` section in `CHANGELOG.md`.
