#!/usr/bin/env bash
set -euo pipefail

repo="${1:-PeachGumi/PhotoOrganizer}"

command -v gh >/dev/null 2>&1 || { echo "GitHub CLI (gh) is required." >&2; exit 1; }
gh auth status >/dev/null

read -r -p "Windows code-signing PFX path: " windows_pfx
read -r -s -p "Windows PFX password: " windows_password
echo
read -r -p "macOS Developer ID P12 path: " mac_p12
read -r -s -p "macOS P12 password: " mac_password
echo
read -r -p "Developer ID Application identity (exact security find-identity text): " developer_id
read -r -p "Apple ID used for notarization: " apple_id
read -r -p "Apple Team ID: " apple_team_id
read -r -s -p "Apple app-specific password: " apple_app_password
echo

[[ -f "$windows_pfx" ]] || { echo "Windows PFX not found." >&2; exit 1; }
[[ -f "$mac_p12" ]] || { echo "macOS P12 not found." >&2; exit 1; }
[[ -n "$windows_password" && -n "$mac_password" && -n "$developer_id" && -n "$apple_id" && -n "$apple_team_id" && -n "$apple_app_password" ]] || {
  echo "All release credentials are required." >&2
  exit 1
}

set_secret() {
  local name="$1"
  gh secret set "$name" --repo "$repo"
  echo "Configured $name"
}

base64 < "$windows_pfx" | tr -d '\n' | set_secret WINDOWS_CERTIFICATE
printf '%s' "$windows_password" | set_secret WINDOWS_CERTIFICATE_PASSWORD
base64 < "$mac_p12" | tr -d '\n' | set_secret MACOS_CERTIFICATE
printf '%s' "$mac_password" | set_secret MACOS_CERTIFICATE_PASSWORD
printf '%s' "$developer_id" | set_secret DEVELOPER_ID_APPLICATION
printf '%s' "$apple_id" | set_secret APPLE_ID
printf '%s' "$apple_team_id" | set_secret APPLE_TEAM_ID
printf '%s' "$apple_app_password" | set_secret APPLE_APP_SPECIFIC_PASSWORD

unset windows_password mac_password apple_app_password

echo "Release secrets configured for $repo. Values were sent to GitHub through stdin and were not printed by this script."
