#!/usr/bin/env bash
# Retries creating an Ampere A1 (VM.Standard.A1.Flex) instance until Oracle
# has capacity, instead of you manually clicking "Create" over and over.
# Handles both errors you've been hitting: "Out of host capacity" and
# "TooManyRequests" (429) -- both just mean "try again shortly", not a real
# failure, so the script sleeps and retries instead of giving up.
#
# WHERE TO RUN THIS: Oracle Cloud Shell (the ">_" icon in the top bar of
# cloud.oracle.com), not your own PC. Cloud Shell already has the OCI CLI
# installed and authenticated to your account -- zero setup needed. Running
# it locally instead would require installing the OCI CLI and generating a
# separate API signing key, which Cloud Shell skips entirely.
#
# BEFORE RUNNING, gather these three things from the Console:
#   1. COMPARTMENT_ID  -- Console -> click your account/profile icon (top
#      right) -> "Tenancy: <name>" -> copy the OCID shown there (this is
#      your tenancy/root compartment -- fine to use directly unless you
#      specifically created a sub-compartment for this).
#   2. SUBNET_ID -- Console -> Networking -> Virtual Cloud Networks -> your
#      VCN -> Subnets -> click the subnet you want the instance on -> copy
#      its OCID. (If you don't have a VCN yet, use the "Create VM instance"
#      wizard once by hand to auto-create the default networking, then
#      cancel before the final Create click -- or just let this script's
#      first failed attempt guide you, the VCN it needs will already exist
#      from your earlier attempts.)
#   3. SSH_PUBLIC_KEY -- the *public* key (not the .key private file) that
#      will let you log into the new instance. If you don't have one yet:
#        ssh-keygen -t ed25519 -f ~/offset-erp-oracle-a1 -N ""
#      then use the contents of offset-erp-oracle-a1.pub here, and keep the
#      matching private key (offset-erp-oracle-a1, no extension) safe --
#      same as the existing offset-erp-oracle.key for the current VM.
#
# USAGE (paste into Cloud Shell):
#   export COMPARTMENT_ID="ocid1.tenancy.oc1..xxxx"
#   export SUBNET_ID="ocid1.subnet.oc1..xxxx"
#   export SSH_PUBLIC_KEY="ssh-ed25519 AAAA... you@host"
#   bash oracle_a1_retry.sh
#
# It tries the full 4 OCPU/24GB shape first, then falls back to 2 OCPU/12GB
# (Oracle halved the default Always-Free A1 allowance in mid-2026) if the
# bigger one won't go through after a while. Leave it running in a Cloud
# Shell tab -- it'll keep retrying every 60s until it succeeds or you stop
# it with Ctrl+C. Cloud Shell disconnects after ~20 min idle, so keep the
# tab focused/active, or run it under `screen`/`tmux` if Cloud Shell has it.

set -uo pipefail

: "${COMPARTMENT_ID:?Set COMPARTMENT_ID first (see script header for how to find it)}"
: "${SUBNET_ID:?Set SUBNET_ID first (see script header for how to find it)}"
: "${SSH_PUBLIC_KEY:?Set SSH_PUBLIC_KEY first (see script header for how to find it)}"

DISPLAY_NAME="${DISPLAY_NAME:-offset-erp-a1}"
RETRY_INTERVAL_SECONDS="${RETRY_INTERVAL_SECONDS:-60}"
# After this many failed attempts on the full-size shape, drop to the
# smaller one instead of waiting forever for the bigger one specifically.
ATTEMPTS_BEFORE_FALLBACK="${ATTEMPTS_BEFORE_FALLBACK:-30}"

echo "Looking up availability domain and a matching Ubuntu 24.04 ARM image..."

AD_NAME=$(oci iam availability-domain list \
  --compartment-id "$COMPARTMENT_ID" \
  --query "data[0].name" --raw-output)

if [ -z "$AD_NAME" ]; then
  echo "Could not find an availability domain in this compartment. Check COMPARTMENT_ID." >&2
  exit 1
fi
echo "Availability domain: $AD_NAME"

IMAGE_ID=$(oci compute image list \
  --compartment-id "$COMPARTMENT_ID" \
  --operating-system "Canonical Ubuntu" \
  --operating-system-version "24.04" \
  --shape "VM.Standard.A1.Flex" \
  --sort-by TIMECREATED --sort-order DESC \
  --query "data[0].id" --raw-output)

if [ -z "$IMAGE_ID" ] || [ "$IMAGE_ID" == "null" ]; then
  echo "Could not find an Ubuntu 24.04 ARM image automatically. Check the Console" >&2
  echo "under Compute -> Images and pass its OCID as IMAGE_ID instead." >&2
  exit 1
fi
echo "Image: $IMAGE_ID"

attempt=0
shape_ocpus=4
shape_memory=24

while true; do
  attempt=$((attempt + 1))

  if [ "$attempt" -eq "$ATTEMPTS_BEFORE_FALLBACK" ]; then
    echo ""
    echo "Still no capacity for 4 OCPU/24GB after $((attempt - 1)) attempts."
    echo "Falling back to 2 OCPU/12GB for the remaining retries."
    shape_ocpus=2
    shape_memory=12
  fi

  echo "[$(date '+%Y-%m-%d %H:%M:%S')] Attempt #$attempt -- trying ${shape_ocpus} OCPU / ${shape_memory}GB in $AD_NAME..."

  result=$(oci compute instance launch \
    --compartment-id "$COMPARTMENT_ID" \
    --availability-domain "$AD_NAME" \
    --display-name "$DISPLAY_NAME" \
    --shape "VM.Standard.A1.Flex" \
    --shape-config "{\"ocpus\": $shape_ocpus, \"memoryInGBs\": $shape_memory}" \
    --image-id "$IMAGE_ID" \
    --subnet-id "$SUBNET_ID" \
    --assign-public-ip true \
    --metadata "{\"ssh_authorized_keys\": \"$SSH_PUBLIC_KEY\"}" \
    --wait-for-state RUNNING \
    --max-wait-seconds 300 \
    2>&1)
  status=$?

  if [ "$status" -eq 0 ]; then
    echo ""
    echo "=================================================="
    echo " SUCCESS -- instance is up after $attempt attempt(s)."
    echo "=================================================="
    echo "$result" | grep -E '"id"|"display-name"|"lifecycle-state"' | head -5
    echo ""
    echo "Find its public IP: Console -> Compute -> Instances -> $DISPLAY_NAME"
    printf '\a'
    exit 0
  fi

  if echo "$result" | grep -qiE "out of (host )?capacity|toomanyrequests|429"; then
    echo "  -> capacity/rate-limit issue, will retry in ${RETRY_INTERVAL_SECONDS}s."
  else
    echo ""
    echo "Got an error that isn't a capacity/rate-limit issue -- stopping so you can look at it:"
    echo "$result"
    exit 1
  fi

  sleep "$RETRY_INTERVAL_SECONDS"
done
