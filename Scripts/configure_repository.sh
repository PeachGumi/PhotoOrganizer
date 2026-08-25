#!/usr/bin/env bash
set -euo pipefail

REPO="PeachGumi/PhotoOrganizer"

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

# Public repositories can use Dependabot/vulnerability alerts. These calls may be
# policy-dependent, so a failure is reported but does not hide the successful core settings.
gh api --method PUT "repos/${REPO}/vulnerability-alerts" >/dev/null || echo "warning: could not enable vulnerability alerts"
gh api --method PUT "repos/${REPO}/automated-security-fixes" >/dev/null || echo "warning: could not enable automated security fixes"

cat <<'JSON' | gh api --method PUT "repos/PeachGumi/PhotoOrganizer/branches/main/protection" --input - >/dev/null
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

echo "Repository settings and main branch protection configured."
