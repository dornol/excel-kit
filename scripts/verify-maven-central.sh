#!/usr/bin/env bash
set -euo pipefail

usage() {
  echo "Usage: $0 <version> [--attempts N] [--sleep-seconds N]" >&2
}

version="${1:-}"
attempts=60
sleep_seconds=30

if [[ -z "${version}" ]]; then
  usage
  exit 2
fi

shift || true
while [[ $# -gt 0 ]]; do
  case "$1" in
    --attempts)
      attempts="${2:-}"
      shift 2
      ;;
    --sleep-seconds)
      sleep_seconds="${2:-}"
      shift 2
      ;;
    *)
      usage
      exit 2
      ;;
  esac
done

if ! [[ "${attempts}" =~ ^[0-9]+$ ]] || [[ "${attempts}" -lt 1 ]]; then
  echo "--attempts must be a positive integer" >&2
  exit 2
fi

if ! [[ "${sleep_seconds}" =~ ^[0-9]+$ ]]; then
  echo "--sleep-seconds must be a non-negative integer" >&2
  exit 2
fi

artifact_path="io/github/dornol/excel-kit"
base_url="https://repo1.maven.org/maven2/${artifact_path}"
pom_url="${base_url}/${version}/excel-kit-${version}.pom"
metadata_url="${base_url}/maven-metadata.xml"

for i in $(seq 1 "${attempts}"); do
  if curl -fsI "${pom_url}" >/dev/null; then
    if curl -fsSL "${metadata_url}" | grep -F "<version>${version}</version>" >/dev/null; then
      echo "Maven Central metadata contains ${version}"
    else
      echo "Maven Central POM is visible, but metadata has not caught up yet"
    fi
    echo "Maven Central POM is visible for ${version}: ${pom_url}"
    exit 0
  fi

  echo "Waiting for Maven Central visibility (${i}/${attempts}): ${pom_url}"
  if [[ "${i}" -lt "${attempts}" && "${sleep_seconds}" -gt 0 ]]; then
    sleep "${sleep_seconds}"
  fi
done

echo "Maven Central did not expose ${version} after ${attempts} attempts"
exit 1
