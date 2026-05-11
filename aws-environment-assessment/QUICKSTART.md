# Quick Start Guide

> **Disclaimer:** This is a community sample script provided without support guarantees. It is not an official product and is not covered by any support agreement. Use at your own risk.

This guide gets you from zero to a completed AWS assessment in under 10 minutes.

---

## Step 1 — Install Python dependencies

You need Python 3.10 or later. Check with `python --version`.

```bash
pip install -r requirements.txt
```

This installs three packages: `boto3` (AWS SDK), `openpyxl` (Excel writer), and `tqdm` (progress bars).

---

## Step 2 — Set up AWS credentials

The tool uses the same credentials as the AWS CLI. Pick one of the following:

### Option A: You already use the AWS CLI

```bash
# Verify your credentials are active
aws sts get-caller-identity
```

If that returns your Account ID and ARN, skip to Step 3.

### Option B: First-time setup

```bash
aws configure
```

You'll be prompted for:
- **AWS Access Key ID** — from IAM → Users → Security credentials → Create access key
- **AWS Secret Access Key** — shown once at creation time
- **Default region** — e.g. `us-east-1`
- **Output format** — type `json`

### Option C: Customer account using a named profile

```bash
aws configure --profile customer-name
aws sts get-caller-identity --profile customer-name   # verify it works
```

---

## Step 3 — Run the assessment

```bash
# Scan your default region
python aws_assessment.py

# Scan specific regions
python aws_assessment.py --regions us-east-1 us-west-2 eu-west-1

# Scan every region in the account (most thorough, 5–20 min depending on size)
python aws_assessment.py --all-regions

# Customer account with a named profile
python aws_assessment.py --profile customer-name --all-regions

# Custom output filename
python aws_assessment.py --all-regions --output "CustomerName_$(date +%Y%m%d).xlsx"
```

The `.xlsx` file is saved in the directory you run the script from.

---

## Step 4 — Open the workbook

Open the file in Excel or Google Sheets. Start with the **Summary** sheet — it gives you the full picture without needing to look at individual tabs.

---

## All command-line options

```
python aws_assessment.py [OPTIONS]

Options:
  --regions REGION [REGION ...]   Regions to scan (default: current region)
  --all-regions                   Scan all enabled regions in the account
  --profile PROFILE               AWS CLI profile name
  --output FILENAME               Output .xlsx filename
  --workers N                     Parallel region workers, default 4
  --skip-snapshots                Skip EBS snapshot enumeration (faster)
  --verbose                       Show detailed logging
```

---

## Example commands

### First run — single region, see what happens

```bash
python aws_assessment.py --regions us-east-1 --verbose
```

### Full account scan with a date-stamped file

```bash
# macOS / Linux
python aws_assessment.py --all-regions --output "Assessment_$(date +%Y%m%d).xlsx"

# Windows PowerShell
$date = Get-Date -Format "yyyyMMdd"
python aws_assessment.py --all-regions --output "Assessment_$date.xlsx"

# Windows Command Prompt
python aws_assessment.py --all-regions --output "Assessment.xlsx"
```

### Customer account — complete workflow

```bash
# 1. Add their credentials as a named profile
aws configure --profile acme-corp
#    AWS Access Key ID:     AKIA...
#    AWS Secret Access Key: xxxxxxxx
#    Default region:        us-east-1
#    Default output format: json

# 2. Verify you can authenticate as them
aws sts get-caller-identity --profile acme-corp
#    {
#        "UserId": "AIDA...",
#        "Account": "123456789012",
#        "Arn": "arn:aws:iam::123456789012:user/assessment-readonly"
#    }

# 3. Run the assessment
python aws_assessment.py \
  --profile acme-corp \
  --all-regions \
  --output "AcmeCorp_Assessment_$(date +%Y%m%d).xlsx"
```

### Customer using AWS SSO

```bash
aws sso login --profile acme-sso
python aws_assessment.py \
  --profile acme-sso \
  --all-regions \
  --output "AcmeCorp_$(date +%Y%m%d).xlsx"
```

### Customer using a cross-account IAM role

```bash
# Add this to ~/.aws/config:
#
# [profile acme-readonly]
# role_arn     = arn:aws:iam::123456789012:role/ReadOnlyAssessmentRole
# source_profile = default
# region       = us-east-1

aws sts get-caller-identity --profile acme-readonly   # confirm role assumption works
python aws_assessment.py --profile acme-readonly --all-regions
```

### Large account — fastest scan

```bash
python aws_assessment.py \
  --all-regions \
  --skip-snapshots \
  --workers 10 \
  --output "LargeAccount_$(date +%Y%m%d).xlsx"
```

### Regional scans

```bash
# US only
python aws_assessment.py \
  --regions us-east-1 us-east-2 us-west-1 us-west-2

# Europe only
python aws_assessment.py \
  --regions eu-west-1 eu-west-2 eu-west-3 eu-central-1 eu-north-1

# APAC only
python aws_assessment.py \
  --regions ap-southeast-1 ap-southeast-2 ap-northeast-1 ap-northeast-2 ap-south-1
```

---

## IAM permissions

The tool is read-only. It never creates, modifies, or deletes anything.

**Quickest option** — attach the AWS managed policy:
```
arn:aws:iam::aws:policy/ReadOnlyAccess
```

**Tighter scope** — create a custom policy with only these actions:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": [
        "ec2:Describe*",
        "rds:Describe*",
        "s3:ListAllMyBuckets",
        "s3:GetBucket*",
        "s3:ListBucket",
        "efs:Describe*",
        "fsx:Describe*",
        "dynamodb:List*",
        "dynamodb:Describe*",
        "redshift:Describe*",
        "redshift-serverless:List*",
        "eks:List*",
        "eks:Describe*",
        "ecs:List*",
        "ecs:Describe*",
        "lambda:ListFunctions",
        "workspaces:Describe*",
        "docdb:Describe*",
        "elasticache:Describe*",
        "backup:List*",
        "backup:Get*",
        "cloudwatch:GetMetricStatistics",
        "sts:GetCallerIdentity"
      ],
      "Resource": "*"
    }
  ]
}
```

To create this policy and attach it to an IAM user via CLI:

```bash
# Save the policy above to a file
cat > assessment-policy.json << 'EOF'
{ ...paste policy here... }
EOF

# Create the policy
aws iam create-policy \
  --policy-name AWSAssessmentReadOnly \
  --policy-document file://assessment-policy.json

# Attach to a user
aws iam attach-user-policy \
  --user-name your-iam-user \
  --policy-arn arn:aws:iam::<account-id>:policy/AWSAssessmentReadOnly
```

---

## Tips for large accounts

| Situation | Recommended flags |
|---|---|
| Account with 1,000+ snapshots | `--skip-snapshots` |
| Scanning 10+ regions | `--workers 8` |
| Both of the above | `--all-regions --skip-snapshots --workers 8` |
| Want to watch progress | `--verbose` |

---

## Understanding the output

### Summary sheet layout

```
┌─────────────────────────────────────────────────┐
│           AWS Environment Assessment            │  ← Title + account/date
├──────────┬──────────┬──────────┬──────────┬─────┤
│ TOTAL    │ STORAGE  │  EC2     │  RDS     │ S3  │  ← KPI tiles
│ RESOURCES│  (TiB)   │ running  │available │ etc │
├──────────┴──────────┴──────────┴──────────┴─────┤
│                                                  │
│  Workload Inventory    │  Risk & Findings        │
│  (left column)         │  (right column)         │
│                        │                         │
│  EC2 Instances    45   │  CRITICAL  Unencrypted  │
│  EBS Volumes      80   │  HIGH      No backup tag│
│  RDS / Aurora     12   │  MEDIUM    No versioning│
│  S3 Buckets       28   │                         │
│  ...                   │  Backup Infrastructure  │
│                        │  Region Distribution    │
│  EC2 Breakdown         │  Storage by Service     │
└────────────────────────┴─────────────────────────┘
```

### Colour coding

| Colour | Meaning | Examples |
|---|---|---|
| Red | Critical risk or gap | No backup, public S3 bucket, unencrypted EBS, public RDS |
| Yellow | Warning | Stopped EC2, single-AZ RDS, versioning off, unattached volume |
| Green | Protected | (used in Risk column when count = 0) |

### S3 size columns

S3 buckets show three size columns: **MiB**, **GiB**, and **TiB**. This is intentional — small buckets show a meaningful value in MiB while large buckets are easier to read in TiB.

---

## Troubleshooting

**`NoCredentialsError`**
Your credentials aren't configured. Run `aws configure` or check that your profile name matches what's in `~/.aws/credentials`.

**`AccessDenied` on a specific service**
Your IAM policy is missing permissions for that service. Add the relevant `Describe*` or `List*` action from the policy above.

**`SubscriptionRequiredException` or `OptInRequired` warnings**
Certain services (FSx, Redshift) require explicit opt-in before use. These warnings are normal on accounts that haven't enabled those services — the script skips them and continues.

**S3 sizes show `N/A` or `0`**
S3 size metrics in CloudWatch update once every 24 hours. For buckets with no CloudWatch history yet, the script falls back to direct object listing. On very large buckets (millions of objects) this fallback is capped at 100,000 objects — the CloudWatch metric will populate by the next day.

**Scan is slow**
Add `--skip-snapshots` (EBS snapshot enumeration is the slowest part on large accounts) and increase `--workers` to match the number of regions you're scanning.
