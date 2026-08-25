#!/usr/bin/env bash
set -euo pipefail

repo="${1:-PeachGumi/PhotoOrganizer}"
environment="${2:-production-signing}"

command -v gh >/dev/null 2>&1 || { echo "GitHub CLI (gh) is required." >&2; exit 1; }
gh auth status >/dev/null

# Refuse to place production credentials into an environment until its server-side
# deployment policy has been verified. Run configure_repository.sh first.
gh api "repos/${repo}/environments/${environment}" >/dev/null 2>&1 || {
  echo "Protected environment '${environment}' does not exist. Run Scripts/configure_repository.sh first." >&2
  exit 1
}
[[ "$(gh api "repos/${repo}/environments/${environment}" --jq '.deployment_branch_policy.protected_branches')" == "false" ]]
[[ "$(gh api "repos/${repo}/environments/${environment}" --jq '.deployment_branch_policy.custom_branch_policies')" == "true" ]]
[[ "$(gh api "repos/${repo}/environments/${environment}/deployment-branch-policies?per_page=100" --jq '.total_count')" == "1" ]]
[[ "$(gh api "repos/${repo}/environments/${environment}/deployment-branch-policies?per_page=100" --jq '[.branch_policies[] | select(.name == "main" and ((.type // "branch") == "branch"))] | length')" == "1" ]] || {
  echo "Environment '${environment}' is not restricted to the main branch. Refusing to configure signing credentials." >&2
  exit 1
}

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
  gh secret set "$name" --repo "$repo" --env "$environment"
  echo "Configured environment secret $name"
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

required=(
  WINDOWS_CERTIFICATE WINDOWS_CERTIFICATE_PASSWORD
  MACOS_CERTIFICATE MACOS_CERTIFICATE_PASSWORD DEVELOPER_ID_APPLICATION
  APPLE_ID APPLE_TEAM_ID APPLE_APP_SPECIFIC_PASSWORD
)
for name in "${required[@]}"; do
  gh secret list --repo "$repo" --env "$environment" | awk '{print $1}' | grep -Fxq "$name" || {
    echo "Environment secret verification failed: $name" >&2
    exit 1
  }
done

# Remove legacy repository-level copies only after all protected environment secrets
# have been confirmed. This closes the old repository-wide credential path.
repo_secret_names="$(gh secret list --repo "$repo" | awk '{print $1}')"
for name in "${required[@]}"; do
  if grep -Fxq "$name" <<< "$repo_secret_names"; then
    gh secret delete "$name" --repo "$repo"
    echo "Removed legacy repository secret $name"
  fi
done

echo "Release secrets configured and verified in ${environment} for ${repo}. Values were sent through stdin and were not printed by this script."
