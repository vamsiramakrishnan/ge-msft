#!/usr/bin/env bash
# Check that the built web app carries the brand it was built for (GE_BRAND → brands/<id>/), so an
# upstream update can't quietly bring back the upstream name or icons. Run after the web build.
#
#   scripts/check-brand.sh              # brand from GE_BRAND, else brands/current, else default
set -euo pipefail
cd "$(dirname "$0")/.."

brand="${GE_BRAND:-$(cat brands/current 2>/dev/null || echo default)}" # same order as tools/brand/brand.mjs
out=packages/web-shell/dist-web
dir="brands/$brand"
fail() { echo "brand check failed: $*" >&2; exit 1; }

[ -f "$out/taskpane.html" ] || fail "no build in $out; build first"
[ -f "$dir/brand.json" ] || fail "no brand at $dir"
name="$(node -e "console.log(JSON.parse(require('fs').readFileSync('$dir/brand.json','utf8')).name)")"

# Every page title names this brand, and none still shows a placeholder.
for page in "$out"/*.html; do
  title="$(grep -o '<title>[^<]*</title>' "$page" || true)"
  [ -z "$title" ] || grep -qF "$name" <<<"$title" || fail "$(basename "$page") title is $title, expected $name"
done
if grep -l '%GE_BRAND_' "$out"/*.html >/dev/null; then fail "an HTML page still has a %GE_BRAND_*% placeholder"; fi

# A customer brand must not show the upstream product name on any page.
if [ "$brand" != default ] && grep -l 'Gemini Enterprise' "$out"/*.html >/dev/null; then
  fail "a page still says Gemini Enterprise: $(grep -l 'Gemini Enterprise' "$out"/*.html | xargs -n1 basename | tr '\n' ' ')"
fi

# The icons Office loads are exactly the brand's own files.
for icon in "$dir"/icons/*.png; do
  cmp -s "$icon" "$out/$(basename "$icon")" || fail "$(basename "$icon") is not the $brand icon"
done

echo "brand check passed: $brand ($name)"
