#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "$0")/.." && pwd)"
cd "$repo_root"
version="${1:-$(./gradlew -q printVersion --no-daemon)}"
repository="${2:-$repo_root/build/consumer-repository}"

if [[ $# -lt 2 ]]; then
  ./gradlew :kit:publishMavenPublicationToConsumerTestRepository \
    :excel-kit-spring:publishMavenPublicationToConsumerTestRepository \
    --init-script verification/publish-consumers.init.gradle --no-daemon
fi

# Exercise Gradle module metadata and Maven POM dependency resolution separately.
for pom_only in false true; do
  ./gradlew -p verification/consumers verifyConsumers --rerun-tasks --no-daemon \
    "-PkitVersion=$version" "-PconsumerRepository=$repository" "-PpomOnly=$pom_only"
done
