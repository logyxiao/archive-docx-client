#!/usr/bin/env sh
set -eu

if [ "$#" -ne 1 ]; then
  echo "Usage: sh scripts/release-tag.sh v0.1.1"
  exit 1
fi

tag="$1"

case "$tag" in
  v[0-9]*.[0-9]*.[0-9]*) ;;
  *)
    echo "Tag must look like v0.1.1"
    exit 1
    ;;
esac

if git rev-parse "$tag" >/dev/null 2>&1; then
  echo "Tag already exists locally: $tag"
  exit 1
fi

if git ls-remote --tags origin "refs/tags/$tag" | grep -q "$tag"; then
  echo "Tag already exists on origin: $tag"
  exit 1
fi

branch="$(git rev-parse --abbrev-ref HEAD)"
git fetch origin "$branch"

if [ "$(git rev-parse HEAD)" != "$(git rev-parse "origin/$branch")" ]; then
  echo "Current branch is not aligned with origin/$branch. Please pull or push your commits first."
  exit 1
fi

git tag -a "$tag" -m "Release $tag"
git push origin "$tag"

echo "Pushed $tag. GitHub Actions will build and publish release bundles."
