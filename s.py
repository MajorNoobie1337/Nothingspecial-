#!/usr/bin/env python3
"""Search every IBM Cloud Schematics workspace for a Terraform resource type.
 
Equivalent to running `terraform state list` in each workspace and grepping,
but done through the Schematics API so nothing needs initializing locally.
 
Usage:
    export SCHEMATICS_APIKEY=...
    python check_schematics_resource.py --accounts accounts.txt -o found.csv
    python check_schematics_resource.py --prefix sailpoint_ --regions us-south,eu-de
    python check_schematics_resource.py --list-types      # what types exist at all
 
accounts.txt is one account_id per line. Omit it to use whatever account the
API key defaults to.
 
CSV columns:
    account, region, workspace_name, workspace_id, workspace_status,
    found, match_count, matched_types, resource_total, detail
 
Statuses in `found`:
    YES          at least one resource of the wanted type
    NO           state read fine, no such resource
    NO_STATE     workspace has no state yet (never applied, or draft)
    AUTH_ERROR   token rejected -> result is meaningless
    HTTP_ERROR   anything else -> result is meaningless, see detail
"""
 
import argparse
import csv
import os
import sys
import time
from collections import Counter
 
import requests
import urllib3
 
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
 
IAM_URL = "https://iam.cloud.ibm.com/identity/token"
APIKEY_GRANT = "urn:ibm:params:oauth:grant-type:apikey"
 
# Schematics is regional; a workspace is only visible from its own endpoint.
REGION_ENDPOINTS = {
    "us-south": "https://schematics.cloud.ibm.com",
    "us-east": "https://us-east.schematics.cloud.ibm.com",
    "eu-de": "https://eu-de.schematics.cloud.ibm.com",
    "eu-gb": "https://eu-gb.schematics.cloud.ibm.com",
    "ca-tor": "https://ca-tor.schematics.cloud.ibm.com",
    "jp-tok": "https://jp-tok.schematics.cloud.ibm.com",
    "au-syd": "https://au-syd.schematics.cloud.ibm.com",
}
 
FIELDS = [
    "account", "region", "workspace_name", "workspace_id", "workspace_status",
    "found", "match_count", "matched_types", "resource_total", "detail",
]
 
AUTH_MARKERS = ("401", "403", "Unauthorized", "invalid_grant", "Forbidden")
 
 
def is_auth_error(msg):
    return any(m in msg for m in AUTH_MARKERS)
 
 
# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
 
def get_tokens(api_key):
    """Return (access_token, refresh_token) for the key's default account."""
    resp = requests.post(
        IAM_URL,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={"grant_type": APIKEY_GRANT, "apikey": api_key, "response_type": "cloud_iam"},
        timeout=30,
    )
    resp.raise_for_status()
    body = resp.json()
    return body["access_token"], body.get("refresh_token", "")
 
 
def token_for_account(refresh_token, account_id):
    """Exchange a refresh token for a token scoped to *account_id*."""
    resp = requests.post(
        IAM_URL,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "account": account_id,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]
 
 
# ---------------------------------------------------------------------------
# Schematics
# ---------------------------------------------------------------------------
 
def list_workspaces(endpoint, token):
    """All workspaces at one regional endpoint, following pagination."""
    workspaces = []
    offset, limit = 0, 100
    while True:
        resp = requests.get(
            f"{endpoint}/v1/workspaces",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            params={"limit": limit, "offset": offset},
            timeout=60,
        )
        resp.raise_for_status()
        body = resp.json()
        page = body.get("workspaces", [])
        workspaces.extend(page)
        offset += limit
        if len(page) < limit or offset >= body.get("total_count", 0):
            break
    return workspaces
 
 
def get_state(endpoint, token, workspace_id, template_id):
    """Raw Terraform state JSON for one template in a workspace."""
    resp = requests.get(
        f"{endpoint}/v1/workspaces/{workspace_id}/runtime_data/{template_id}/state_store",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
        timeout=90,
    )
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    return resp.json()
 
 
def resource_types(state):
    """Every resource type in a state file, managed resources only."""
    types = []
    for res in (state or {}).get("resources", []):
        if res.get("mode") == "data":
            continue
        rtype = res.get("type")
        if rtype:
            count = max(1, len(res.get("instances", [])))
            types.extend([rtype] * count)
    return types
 
 
# ---------------------------------------------------------------------------
# Per-workspace check
# ---------------------------------------------------------------------------
 
def check_workspace(endpoint, token, workspace, prefix, account_label, region):
    row = {
        "account": account_label, "region": region,
        "workspace_name": workspace.get("name", ""),
        "workspace_id": workspace.get("id", ""),
        "workspace_status": workspace.get("status", ""),
        "found": "NO", "match_count": 0, "matched_types": "",
        "resource_total": 0, "detail": "",
    }
 
    templates = workspace.get("template_data") or []
    if not templates:
        row["found"] = "NO_STATE"
        row["detail"] = "workspace has no templates"
        return row, []
 
    all_types = []
    had_state = False
    for template in templates:
        template_id = template.get("id")
        if not template_id:
            continue
        try:
            state = get_state(endpoint, token, workspace["id"], template_id)
        except Exception as exc:
            msg = str(exc)
            row["found"] = "AUTH_ERROR" if is_auth_error(msg) else "HTTP_ERROR"
            row["detail"] = msg[:200]
            return row, []
        if state is None:
            continue
        had_state = True
        all_types.extend(resource_types(state))
 
    if not had_state:
        row["found"] = "NO_STATE"
        row["detail"] = "no state stored yet"
        return row, []
 
    matched = [t for t in all_types if t.startswith(prefix)]
    row["resource_total"] = len(all_types)
    row["match_count"] = len(matched)
    row["matched_types"] = ";".join(sorted(set(matched)))
    row["found"] = "YES" if matched else "NO"
    return row, all_types
 
 
# ---------------------------------------------------------------------------
 
def main():
    ap = argparse.ArgumentParser(
        description="Find a Terraform resource type across Schematics workspaces.")
    ap.add_argument("--api-key", default=os.environ.get("SCHEMATICS_APIKEY"),
                    help="IBM Cloud API key (or set SCHEMATICS_APIKEY)")
    ap.add_argument("--accounts", help="file of account_ids, one per line")
    ap.add_argument("--prefix", default="sailpoint_",
                    help="resource type prefix to look for (default sailpoint_)")
    ap.add_argument("--regions", default="us-south,eu-de",
                    help="comma-separated regions (default us-south,eu-de)")
    ap.add_argument("-o", "--out", help="CSV file to write (default stdout)")
    ap.add_argument("--list-types", action="store_true",
                    help="instead of matching, print every type seen and its count")
    args = ap.parse_args()
 
    if not args.api_key:
        sys.exit("No API key. Pass --api-key or set SCHEMATICS_APIKEY.")
 
    regions = [r.strip() for r in args.regions.split(",") if r.strip()]
    unknown = [r for r in regions if r not in REGION_ENDPOINTS]
    if unknown:
        sys.exit(f"Unknown region(s): {unknown}. Known: {sorted(REGION_ENDPOINTS)}")
 
    try:
        base_token, refresh_token = get_tokens(args.api_key)
    except Exception as exc:
        sys.exit(f"IAM auth failed: {exc}")
 
    if args.accounts:
        with open(args.accounts) as fh:
            accounts = [line.strip() for line in fh if line.strip()]
        if not refresh_token:
            sys.exit("No refresh token returned; cannot switch accounts. "
                     "Run without --accounts to use the key's default account.")
    else:
        accounts = [None]
 
    handle = open(args.out, "w", newline="") if args.out else sys.stdout
    writer = csv.DictWriter(handle, fieldnames=FIELDS)
    writer.writeheader()
 
    seen_types = Counter()
    counts = Counter()
 
    try:
        for account in accounts:
            label = account or "default"
            if account:
                try:
                    token = token_for_account(refresh_token, account)
                except Exception as exc:
                    writer.writerow({
                        **{f: "" for f in FIELDS},
                        "account": label, "found": "AUTH_ERROR",
                        "match_count": 0, "resource_total": 0,
                        "detail": f"cannot scope token to account: {str(exc)[:150]}",
                    })
                    counts["AUTH_ERROR"] += 1
                    continue
            else:
                token = base_token
 
            for region in regions:
                endpoint = REGION_ENDPOINTS[region]
                try:
                    workspaces = list_workspaces(endpoint, token)
                except Exception as exc:
                    msg = str(exc)
                    writer.writerow({
                        **{f: "" for f in FIELDS},
                        "account": label, "region": region,
                        "found": "AUTH_ERROR" if is_auth_error(msg) else "HTTP_ERROR",
                        "match_count": 0, "resource_total": 0,
                        "detail": msg[:200],
                    })
                    counts["HTTP_ERROR"] += 1
                    continue
 
                print(f"[{label}/{region}] {len(workspaces)} workspace(s)",
                      file=sys.stderr)
 
                for workspace in workspaces:
                    row, types = check_workspace(
                        endpoint, token, workspace, args.prefix, label, region)
                    seen_types.update(types)
                    counts[row["found"]] += 1
                    writer.writerow(row)
                    handle.flush()
                    time.sleep(0.1)  # be gentle with the API
    finally:
        if args.out:
            handle.close()
 
    if args.list_types:
        print("\nresource types seen:", file=sys.stderr)
        for rtype, n in seen_types.most_common():
            print(f"  {n:6d}  {rtype}", file=sys.stderr)
 
    summary = "  ".join(f"{k}={v}" for k, v in sorted(counts.items()))
    print(f"\n{sum(counts.values())} workspace(s):  {summary}", file=sys.stderr)
    sys.exit(0 if counts.get("YES") else 1)
 
 
if __name__ == "__main__":
    main()
 
