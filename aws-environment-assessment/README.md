# AWS Environment Assessment Tool

> **Disclaimer:** This is a community sample script provided without support guarantees. It is not an official product and is not covered by any support agreement. Use at your own risk. Review the code before running it in any environment.

A read-only AWS inventory tool that scans your account and produces a single Excel workbook covering every major workload type. Think of it as RVTools — but for AWS.

The output is a colour-coded, multi-sheet spreadsheet your team can use to understand what's running, what's at risk, and how much storage needs to be protected.

---

## How it works

1. You point it at an AWS account using your existing AWS credentials
2. It scans every enabled region in parallel (or just the ones you specify)
3. It writes a single `.xlsx` file with one sheet per service type plus a summary dashboard

The script is **100% read-only** — it only calls `Describe*`, `List*`, and `Get*` APIs. It makes no changes to your environment.

---

## Prerequisites

**Python 3.10 or later**

```bash
pip install -r requirements.txt
```

**AWS credentials configured**

The tool uses whatever credentials are already set up on your machine — the same ones the AWS CLI uses. If you can run `aws s3 ls`, you're ready.

```bash
# Check your credentials are working
aws sts get-caller-identity
```

You should see your Account ID, User ID, and ARN. If not, run `aws configure` first.

---

## Quickstart

```bash
# Scan your current region
python aws_assessment.py

# Scan specific regions
python aws_assessment.py --regions us-east-1 us-west-2 eu-west-1

# Scan every enabled region in the account
python aws_assessment.py --all-regions
```

The output file is saved in the current directory:
```
aws_assessment_<account-id>_<date>.xlsx
```

Open it in Excel or Google Sheets.

---

## All options

| Flag | Description | Default |
|---|---|---|
| `--regions` | One or more specific regions to scan | Current configured region |
| `--all-regions` | Scan every enabled region in the account | — |
| `--profile` | AWS CLI named profile | Default profile |
| `--output` | Output `.xlsx` filename | `aws_assessment_<account>_<date>.xlsx` |
| `--workers` | Number of regions scanned in parallel | `4` |
| `--skip-snapshots` | Skip EBS snapshot enumeration | — |
| `--verbose` | Print detailed per-service logging | — |

---

## Examples

### Basic scans

```bash
# Scan your current region with default credentials
python aws_assessment.py

# Scan a single specific region
python aws_assessment.py --regions us-east-1

# Scan multiple specific regions
python aws_assessment.py --regions us-east-1 us-west-2 eu-west-1 ap-southeast-1

# Scan every enabled region in the account
python aws_assessment.py --all-regions
```

### Customer accounts

```bash
# First: add the customer's credentials as a named profile
aws configure --profile acme-corp
# (enter their Access Key ID, Secret Access Key, region, output=json)

# Verify the credentials work
aws sts get-caller-identity --profile acme-corp

# Run a full assessment — all regions, date-stamped output file
python aws_assessment.py \
  --profile acme-corp \
  --all-regions \
  --output "AcmeCorp_Assessment_$(date +%Y%m%d).xlsx"
```

```bash
# Customer using AWS SSO
aws sso login --profile acme-sso
python aws_assessment.py \
  --profile acme-sso \
  --all-regions \
  --output "AcmeCorp_SSO_$(date +%Y%m%d).xlsx"
```

```bash
# Customer using a cross-account IAM role (assume role)
# First configure the profile in ~/.aws/config:
#
# [profile acme-readonly]
# role_arn = arn:aws:iam::123456789012:role/ReadOnlyAssessmentRole
# source_profile = default
# region = us-east-1
#
aws sts get-caller-identity --profile acme-readonly   # verify
python aws_assessment.py --profile acme-readonly --all-regions
```

### Large or complex accounts

```bash
# Skip EBS snapshots — biggest time saving on accounts with thousands of snapshots
python aws_assessment.py --all-regions --skip-snapshots

# Increase parallel workers for faster multi-region scans (one worker per region)
python aws_assessment.py --all-regions --workers 10

# Both together — fastest possible full-account scan
python aws_assessment.py \
  --all-regions \
  --skip-snapshots \
  --workers 10 \
  --output "LargeAccount_$(date +%Y%m%d).xlsx"
```

### Targeted scans

```bash
# US regions only
python aws_assessment.py --regions us-east-1 us-east-2 us-west-1 us-west-2

# Europe only
python aws_assessment.py --regions eu-west-1 eu-west-2 eu-west-3 eu-central-1 eu-north-1

# APAC only
python aws_assessment.py --regions ap-southeast-1 ap-southeast-2 ap-northeast-1 ap-south-1

# Single region with verbose logging (useful for debugging or first run)
python aws_assessment.py --regions us-east-1 --verbose
```

### Windows users (Command Prompt)

```cmd
rem Basic scan
python aws_assessment.py --regions us-east-1

rem Customer profile, all regions
python aws_assessment.py --profile acme-corp --all-regions --output "AcmeCorp_Assessment.xlsx"
```

### Windows users (PowerShell)

```powershell
# Basic scan
python aws_assessment.py --regions us-east-1

# Date-stamped output
$date = Get-Date -Format "yyyyMMdd"
python aws_assessment.py --profile acme-corp --all-regions --output "AcmeCorp_$date.xlsx"
```

---

## Setting up credentials for a customer account

```bash
# Step 1: Add a named profile for the customer
aws configure --profile customer-name
# Prompts for:
#   AWS Access Key ID:     (paste their key)
#   AWS Secret Access Key: (paste their secret)
#   Default region:        us-east-1
#   Default output format: json

# Step 2: Confirm it works — you should see their Account ID
aws sts get-caller-identity --profile customer-name

# Step 3: Run the assessment
python aws_assessment.py \
  --profile customer-name \
  --all-regions \
  --output "CustomerName_$(date +%Y%m%d).xlsx"
```

If the customer uses AWS SSO:
```bash
aws sso login --profile customer-name
python aws_assessment.py --profile customer-name --all-regions
```

---

## IAM permissions required

The tool only needs read-only access. The quickest option is to attach the AWS managed policy:

```
arn:aws:iam::aws:policy/ReadOnlyAccess
```

If the customer prefers a tighter scope, use this minimal custom policy:

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

---

## What's in the workbook

### Summary dashboard (first sheet)

The Summary sheet is a full dashboard — no need to click through every tab to get an overview.

| Section | What it shows |
|---|---|
| **KPI tiles** | Total resources, total storage (TiB), EC2 running/stopped, RDS available/stopped, S3 object count, snapshot count |
| **Workload inventory** | Every service type with resource count, storage in GiB and TiB, encryption %, and which regions it's deployed in |
| **Risk & Findings** | Colour-coded CRITICAL / HIGH / MEDIUM findings with counts — unencrypted volumes, publicly accessible resources, missing backup configuration, no PITR, unattached volumes |
| **AWS Backup infrastructure** | Existing backup vaults, locked (WORM) vaults, total recovery points, backup plans, cross-region copy rules |
| **Region distribution** | How many resources are in each region |
| **Storage by service** | Which services consume the most storage, ranked |
| **EC2 state breakdown** | Instance counts by state, top instance types, OS breakdown |

### Detail sheets (one per service)

| Sheet | What you get |
|---|---|
| **EC2 Instances** | Instance ID, name, type, state, OS, AZ, environment tag, owner tag, root disk, data disks, total storage, backup tag, VPC |
| **EBS Volumes** | Volume ID, type (gp3/io2/etc), size, IOPS, throughput, encrypted, attachment status |
| **EBS Snapshots** | All snapshots owned by the account — size, encryption, age |
| **RDS & Aurora** | Engine, version, instance class, storage, backup retention, Multi-AZ, encryption, public accessibility |
| **S3 Buckets** | Size in MiB/GiB/TiB, object count, versioning, replication, lifecycle rules, encryption, public access block status |
| **EFS** | File system size broken down by Standard and Infrequent Access tiers, throughput mode |
| **FSx** | All FSx types (Windows, Lustre, ONTAP, OpenZFS) with capacity and configuration |
| **DynamoDB** | Table size, item count, PITR status, billing mode (on-demand vs provisioned), global tables |
| **Redshift** | Clusters and Serverless namespaces, node type, node count, backup retention |
| **EKS** | Cluster version, node groups, total nodes, node instance types |
| **ECS** | Active services, running tasks, capacity providers (Fargate/EC2) |
| **Lambda** | Runtime, memory, timeout, package size, architecture |
| **WorkSpaces** | User, bundle, root and user volume sizes, running mode |
| **DocumentDB** | Cluster members, storage, backup retention, encryption |
| **ElastiCache** | Engine (Redis/Memcached), node type, node count, backup retention |
| **AWS Backup Vaults** | Recovery point count, immutability lock, retention limits |
| **AWS Backup Plans** | Schedule, target vault, retention period, cross-region copy destination |

---

## Colour coding

Every cell that represents a risk or gap is highlighted automatically:

| Colour | Meaning |
|---|---|
| 🔴 Red | Critical gap — no backup, publicly accessible resource, unencrypted storage, 0-day retention |
| 🟡 Yellow | Warning — stopped instance, unattached volume, versioning disabled, single-AZ database |
| 🟢 Green | Protected / compliant |

---

## Security

This section explains exactly what the tool does and does not do with your AWS account and data — so you can share it with security-conscious customers with confidence.

### The tool is 100% read-only

Every AWS API call made by this script is a read operation — `Describe*`, `List*`, or `Get*`. There are no `Create`, `Put`, `Update`, `Delete`, or `Modify` calls anywhere in the code. It is **not possible** for this tool to create, change, or delete any resource in your account.

You can verify this yourself:
```bash
# Search the source for any write operations
grep -E "(create_|delete_|put_|update_|modify_|terminate_|run_|start_|stop_)" aws_assessment.py
# Returns nothing
```

### Credentials stay on your machine

- AWS credentials are passed directly to the `boto3` SDK — the same library used by the official AWS CLI
- Credentials are **never** printed, logged, written to files, or transmitted anywhere other than to AWS API endpoints
- The script has no knowledge of your access keys; it only calls `boto3.Session(profile_name=...)` and boto3 handles everything else

### No data leaves your machine

- The only output is the `.xlsx` file written locally to your current directory
- The script makes **no HTTP calls** to any server other than official AWS API endpoints (`*.amazonaws.com`)
- There is no telemetry, no analytics, no call-home behaviour of any kind
- No third-party libraries with network capability are used — only `boto3` (AWS), `openpyxl` (local Excel writing), and `tqdm` (local progress bar)

You can verify the network behaviour:
```bash
# Confirm no unexpected imports
grep -E "^import|^from" aws_assessment.py | grep -vE "boto3|openpyxl|json|sys|argparse|datetime|logging|collections|concurrent"
# Returns nothing
```

### What's in the output file

The `.xlsx` file contains only resource **metadata** — the same information visible in the AWS Console:

- Resource IDs, names, types, and sizes
- Configuration flags (encrypted: yes/no, multi-AZ: yes/no, public: yes/no)
- Region and availability zone
- Tags you've applied to resources
- Counts and storage totals

It does **not** contain:
- AWS access keys or secret keys
- Database passwords or connection strings
- KMS key material or key ARNs
- Application secrets or environment variable values
- Any data stored inside your resources (no S3 object contents, no database rows)

### IAM permissions

The tool only needs read-only access. We recommend creating a dedicated IAM user or role with the minimal policy in this README — not your admin credentials.

The minimal policy grants access to 25 specific `Describe*`/`List*`/`Get*` actions and nothing else. You can revoke it immediately after the assessment is complete.

### Open source

The full source code is in this repository. There are no compiled binaries, no obfuscated code, and no external dependencies beyond the three packages in `requirements.txt`. You or your security team can review every line before running it.

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `NoCredentialsError` | No AWS credentials found | Run `aws configure` or set `AWS_PROFILE` |
| `AccessDenied` on a service | IAM policy missing that service | Add the relevant `Describe*` / `List*` actions |
| `SubscriptionRequiredException` | Service not enabled in that region | Normal — the script skips it automatically |
| `EndpointResolutionError` | Service unavailable in region | Normal — skipped automatically |
| S3 sizes show `0` or `N/A` | CloudWatch metrics update once daily | The script falls back to direct listing; give it a moment on very large buckets |
| Scan takes a long time | Large account or many regions | Add `--skip-snapshots` and increase `--workers` |

---

## Files

| File | Purpose |
|---|---|
| `aws_assessment.py` | The assessment script |
| `requirements.txt` | Python dependencies (`boto3`, `openpyxl`, `tqdm`) |
| `QUICKSTART.md` | Step-by-step setup guide including IAM policy |
