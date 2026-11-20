#!/usr/bin/env bash
# Checks that every API named in a CHANGELOG "### Added" section appears in each documentation set
# (README, docs/guide, AI.md, META-INF/excel-kit/*.md), as required by the release checklist.
#
# Usage: scripts/verify-release-docs.sh [version]   (default: version from build.gradle.kts)
# Names that are not excel-kit APIs (JDK types, etc.) go in scripts/release-docs-ignore.txt.
set -euo pipefail

repo_root="$(cd "$(dirname "$0")/.." && pwd)"
cd "$repo_root"

version="${1:-$(sed -nE 's/^[[:space:]]*version = "([^"]+)".*/\1/p' build.gradle.kts | head -1)}"
[[ -n "$version" ]] || { echo "Cannot determine version" >&2; exit 2; }

# The "### Added" body of "## [version]" up to the next heading.
section="$(awk -v v="$version" '
  /^## \[/ { in_version = (index($0, "[" v "]") > 0); in_added = 0; next }
  in_version && /^### / { in_added = ($0 == "### Added"); next }
  in_version && in_added { print }
' CHANGELOG.md)"

if [[ -z "$section" ]]; then
  echo "No '### Added' section for $version in CHANGELOG.md; nothing to verify."
  exit 0
fi

ignore_file="scripts/release-docs-ignore.txt"
# `Type.method(args)` -> method; keep only plain identifiers.
apis="$(printf '%s\n' "$section" \
  | grep -oE '`[^`]+`' | tr -d '`' \
  | sed -E 's/\(.*$//; s/.*\.//' \
  | grep -E '^[A-Za-z_][A-Za-z0-9_]*$' | sort -u || true)"
if [[ -f "$ignore_file" ]]; then
  ignored="$(grep -vE '^\s*(#|$)' "$ignore_file" || true)"
  apis="$(printf '%s\n' "$apis" | grep -vxF -f <(printf '%s\n' "$ignored") || true)"
fi

# Each documentation group must mention the API at least once.
groups=(
  "README.md"
  "docs/guide"
  "kit/src/main/resources/META-INF/AI.md"
  "kit/src/main/resources/META-INF/excel-kit"
)

missing=0
checked=0
while IFS= read -r api; do
  [[ -n "$api" ]] || continue
  checked=$((checked + 1))
  absent=()
  for group in "${groups[@]}"; do
    grep -rqw -- "$api" "$group" || absent+=("$group")
  done
  if (( ${#absent[@]} > 0 )); then
    echo "MISSING: $api -> ${absent[*]}"
    missing=$((missing + 1))
  fi
done <<< "$apis"

echo "Checked $checked API name(s) for $version; $missing missing."
(( missing == 0 ))
