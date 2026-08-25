#!/usr/bin/env bash
set -euo pipefail

REPO="${1:-PeachGumi/PhotoOrganizer}"
SIGNING_ENVIRONMENT="production-signing"
RELEASE_ENVIRONMENT="production-release"

command -v gh >/dev/null 2>&1 || { echo "GitHub CLI (gh) is required." >&2; exit 1; }
gh auth status >/dev/null

echo "Configuring repository settings for ${REPO}..."

gh api --method PATCH "repos/${REPO}" \
  -f description='Cross-platform photo/video organizer for Windows and macOS with fail-safe SD-card import verification.' \
  -F has_issues=true \
  -F has_projects=false \
  -F has_wiki=false \
  -F allow_squash_merge=true \
  -F allow_merge_commit=false \
  -F allow_rebase_merge=false \
  -F allow_auto_merge=true \
  -F delete_branch_on_merge=true \
  -F allow_update_branch=true >/dev/null

gh api --method PUT "repos/${REPO}/topics" --input - >/dev/null <<'JSON'
{
  "names": ["photography", "photo-organizer", "windows", "macos", "dotnet", "avalonia"]
}
JSON

# Keep the default GITHUB_TOKEN read-only. Workflows that genuinely publish must
# request their narrow write permission explicitly.
gh api --method PUT "repos/${REPO}/actions/permissions/workflow" --input - >/dev/null <<'JSON'
{
  "default_workflow_permissions": "read",
  "can_approve_pull_request_reviews": false
}
JSON

# Public repositories can use Dependabot/vulnerability alerts. These calls may be
# policy-dependent, so a failure is reported but does not hide the successful core settings.
gh api --method PUT "repos/${REPO}/vulnerability-alerts" >/dev/null || echo "warning: could not enable vulnerability alerts"
gh api --method PUT "repos/${REPO}/automated-security-fixes" >/dev/null || echo "warning: could not enable automated security fixes"

cat <<'JSON' | gh api --method PUT "repos/${REPO}/branches/main/protection" --input - >/dev/null
{
  "required_status_checks": {
    "strict": true,
    "contexts": ["required"]
  },
  "enforce_admins": true,
  "required_pull_request_reviews": {
    "dismiss_stale_reviews": true,
    "require_code_owner_reviews": false,
    "required_approving_review_count": 0
  },
  "restrictions": null,
  "required_linear_history": true,
  "allow_force_pushes": false,
  "allow_deletions": false,
  "block_creations": false,
  "required_conversation_resolution": true
}
JSON

configure_main_only_environment() {
  local environment="$1"

  cat <<'JSON' | gh api --method PUT "repos/${REPO}/environments/${environment}" --input - >/dev/null
{
  "deployment_branch_policy": {
    "protected_branches": false,
    "custom_branch_policies": true
  }
}
JSON

  while IFS= read -r policy_id; do
    [[ -n "$policy_id" ]] || continue
    gh api --method DELETE \
      "repos/${REPO}/environments/${environment}/deployment-branch-policies/${policy_id}" >/dev/null
  done < <(gh api "repos/${REPO}/environments/${environment}/deployment-branch-policies?per_page=100" \
    --jq '.branch_policies[].id')

  gh api --method POST \
    "repos/${REPO}/environments/${environment}/deployment-branch-policies" \
    -f name='main' -f type='branch' >/dev/null
}

verify_main_only_environment() {
  local environment="$1"
  [[ "$(gh api "repos/${REPO}/environments/${environment}" --jq '.deployment_branch_policy.protected_branches')" == "false" ]]
  [[ "$(gh api "repos/${REPO}/environments/${environment}" --jq '.deployment_branch_policy.custom_branch_policies')" == "true" ]]
  [[ "$(gh api "repos/${REPO}/environments/${environment}/deployment-branch-policies?per_page=100" --jq '.total_count')" == "1" ]]
  [[ "$(gh api "repos/${REPO}/environments/${environment}/deployment-branch-policies?per_page=100" --jq '[.branch_policies[] | select(.name == "main" and ((.type // "branch") == "branch"))] | length')" == "1" ]]
}

# Both credential access and stable-release mutation are protected server-side.
# Reset custom policies to exactly one main-branch rule, making the helper idempotent.
configure_main_only_environment "$SIGNING_ENVIRONMENT"
configure_main_only_environment "$RELEASE_ENVIRONMENT"

# Post-configuration verification is intentionally fatal. Do not print a success
# message when GitHub rejected or partially applied a hardening setting.
[[ "$(gh api "repos/${REPO}" --jq '.allow_squash_merge')" == "true" ]]
[[ "$(gh api "repos/${REPO}" --jq '.allow_merge_commit')" == "false" ]]
[[ "$(gh api "repos/${REPO}" --jq '.allow_rebase_merge')" == "false" ]]
[[ "$(gh api "repos/${REPO}" --jq '.allow_auto_merge')" == "true" ]]
[[ "$(gh api "repos/${REPO}" --jq '.allow_update_branch')" == "true" ]]

[[ "$(gh api "repos/${REPO}/actions/permissions/workflow" --jq '.default_workflow_permissions')" == "read" ]]
[[ "$(gh api "repos/${REPO}/branches/main/protection" --jq '.enforce_admins.enabled')" == "true" ]]
[[ "$(gh api "repos/${REPO}/branches/main/protection" --jq '.required_status_checks.strict')" == "true" ]]
[[ "$(gh api "repos/${REPO}/branches/main/protection" --jq '[.required_status_checks.contexts[] | select(. == "required")] | length')" == "1" ]]

verify_main_only_environment "$SIGNING_ENVIRONMENT"
verify_main_only_environment "$RELEASE_ENVIRONMENT"

echo "Repository settings, main protection, ${SIGNING_ENVIRONMENT}, and ${RELEASE_ENVIRONMENT} main-only policies verified."
