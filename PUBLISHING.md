# Maven Central Publication Guide

This document describes how to publish `org.aspose:aspose-cells-foss` to Maven Central
and serves as the template for all Aspose FOSS Java packages.

## Quick Reference

- **Coordinate**: `org.aspose:aspose-cells-foss:{version}`
- **Version**: Always matches the commercial Aspose.Cells release (e.g. 26.5.0)
- **GPG Key ID**: `E176D5CBCA1DCC62` (shared across all Aspose FOSS Java packages)
- **Portal**: https://central.sonatype.com

---

## Workflow Overview

The workflow (`.github/workflows/maven-central-release.yml`) has 3 jobs:

| Job | Trigger | Purpose |
|-----|---------|---------|
| `dry-run-ubuntu` | `workflow_dispatch -f target=ubuntu` | Pre-release validation, no deploy |
| `dry-run-windows` | `workflow_dispatch -f target=windows` | Windows runner validation, no deploy |
| `release-deploy` | GitHub Release published | Full sign + deploy to Central Portal |

**Deploy ONLY fires on `release.published`** — `workflow_dispatch` never deploys.

---

## Publishing a New Version

### Step 1: Update version in pom.xml
```xml
<version>26.X.0</version>  <!-- must match the commercial Aspose.Cells release -->
```

### Step 2: Commit and push
```bash
git add pom.xml
git commit -m "chore(version): bump to 26.X.0"
git push origin master
```

### Step 3: Run remote dry-run (validates on GitHub-hosted ubuntu)
```bash
gh workflow run maven-central-release.yml --repo aspose-cells-foss/Aspose.Cells-FOSS-for-Java -f target=ubuntu
```
Check: `gh run list --repo aspose-cells-foss/Aspose.Cells-FOSS-for-Java --workflow maven-central-release.yml`

### Step 4: Create GitHub Release (triggers actual deploy)
```bash
gh release create V26.X.0   --repo aspose-cells-foss/Aspose.Cells-FOSS-for-Java   --title "V26.X.0"   --notes "Aspose.Cells FOSS for Java 26.X.0"   --target master
```

### Step 5: Monitor the release-deploy job
```bash
gh run list --repo aspose-cells-foss/Aspose.Cells-FOSS-for-Java --workflow maven-central-release.yml
gh run view <RUN_ID> --repo aspose-cells-foss/Aspose.Cells-FOSS-for-Java
```

### Step 6: Publish on Central Portal
1. Go to https://central.sonatype.com
2. Find the deployment for `org.aspose:aspose-cells-foss:26.X.0`
3. Wait for **VALIDATED** status
4. Click **Publish**

---

## GitHub Secrets (must be set once per repo)

| Secret | Description |
|--------|-------------|
| `GPG_PRIVATE_KEY` | Armored private key for `E176D5CBCA1DCC62` |
| `GPG_PASSPHRASE` | Passphrase for the GPG key |
| `MAVEN_CENTRAL_USERNAME` | Sonatype Central Portal token username |
| `MAVEN_CENTRAL_PASSWORD` | Sonatype Central Portal token password |

Set from Git Bash:
```bash
gpg --armor --export-secret-keys E176D5CBCA1DCC62 | gh secret set GPG_PRIVATE_KEY --repo aspose-cells-foss/Aspose.Cells-FOSS-for-Java
gh secret set GPG_PASSPHRASE         --repo aspose-cells-foss/Aspose.Cells-FOSS-for-Java
gh secret set MAVEN_CENTRAL_USERNAME --repo aspose-cells-foss/Aspose.Cells-FOSS-for-Java
gh secret set MAVEN_CENTRAL_PASSWORD --repo aspose-cells-foss/Aspose.Cells-FOSS-for-Java
```

> Note: Use Git Bash (not PowerShell) — `gpg` is at `/usr/bin/gpg` in Git Bash.

---

## Self-Hosted Runner (Windows)

Location: `C:	oolsctions-runner-cells-foss\`

Start: `cd C:	oolsctions-runner-cells-foss && .un.cmd`

The runner must have these labels: `self-hosted,Windows,maven-pilot`

If re-registering:
```powershell
$token = (gh api repos/aspose-cells-foss/Aspose.Cells-FOSS-for-Java/actions/runners/registration-token --method POST -q .token)
.\config.cmd --url https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-Java --token $token --name cells-foss-win --labels "self-hosted,Windows,maven-pilot" --work _work --replace
$token = $null
```

---

## Known Issues and Fixes

### mvn help:evaluate fails on Windows runner
`actions/setup-java` with `server-id: central` + `central-publishing-maven-plugin` extension
causes `LifecyclePhaseNotFoundException`. The workflow uses PowerShell XML parsing instead:
```powershell
[xml]$pom = Get-Content pom.xml
$VER = $pom.project.version
```

### PowerShell em dash encoding
Em dash `—` in double-quoted PS strings breaks PS 5.1 (byte 0x94 = smart quote in Win-1252).
Use `--` instead of `—` in double-quoted strings.

### git grep exit code
`git grep` exits 1 when no matches found. Add `exit 0` after clean-case `Write-Host` in scan steps.

### Release tag case
Both `V26.5.0` and `v26.5.0` are accepted. The workflow strips the prefix before comparing.
