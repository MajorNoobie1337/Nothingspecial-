#!/usr/bin/env python3
"""Per-machine check: does this host's Schematics workspace contain a resource type?

Input is a CSV with (at least) a hostname, a subscription_id and a hub account
number. The workspace is located by finding the workspace whose *name contains
the subscription_id*, within that account. Its state is then read through the
Schematics API and searched for resources whose type starts with --prefix.

Usage:
    export SCHEMATICS_APIKEY=...
    python check_workspace_resource.py -f prod_check.csv -o sailpoint.csv
    python check_workspace_resource.py -f prod_check.csv --prefix sailpoint_
    python check_workspace_resource.py -f prod_check.csv --list-types

Input columns (rename with --host-col / --sub-col / --acct-col):
    hostname, subscription_id, account_number
Any other columns are ignored, so the output of check_prod.py works as-is once
you add the account column.

Output columns:
    hostname, subscription_id, account, region, workspace_name, workspace_id,
    workspace_status, found, match_count, matched_types, resource_total, detail

found:
    YES           at least one resource of the wanted type
    NO            state read fine, no such resource
    NO_STATE      workspace exists but has no state yet
    NO_WORKSPACE  no workspace name in that account contains the subscription_id
    NO_SUB_ID     input row had no subscription_id
    AUTH_ERROR    token rejected / account not reachable -> result meaningless
    HTTP_ERROR    anything else -> result meaningless, see detail
"""

import argparse
import csv
import os
import sys
from collections import Counter, OrderedDict

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
    "hostname", "subscription_id", "account", "region",
    "workspace_name", "workspace_id", "workspace_status",
    "found", "match_count", "matched_types", "resource_total", "detail",
]

AUTH_MARKERS = ("401", "403", "Unauthorized", "invalid_grant", "Forbidden")


def is_auth_error(msg):
    return any(m in msg for m in AUTH_MARKERS)


# ---------------------------------------------------------------------------
# Input
# ---------------------------------------------------------------------------

def read_rows(path, host_col, sub_col, acct_col):
    """Return [(hostname, subscription_id, account)] grouped by account."""
    with open(path, newline="") as fh:
        reader = csv.DictReader(fh)
        missing = [c for c in (host_col, sub_col, acct_col)
                   if c not in (reader.fieldnames or [])]
        if missing:
            sys.exit(f"Input CSV is missing column(s): {missing}\n"
                     f"Header found: {reader.fieldnames}")
        rows = [
            ((r.get(host_col) or "").strip(),
             (r.get(sub_col) or "").strip(),
             (r.get(acct_col) or "").strip())
            for r in reader
        ]

    rows = [r for r in rows if r[0] or r[1]]

    grouped = OrderedDict()
    for host, sub, acct in rows:
        grouped.setdefault(acct, []).append((host, sub, acct))
    return grouped


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def get_tokens(api_key):
    resp = requests.post(
        IAM_URL,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={"grant_type": APIKEY_GRANT, "apikey": api_key,
              "response_type": "cloud_iam"},
        timeout=30,
    )
    resp.raise_for_status()
    body = resp.json()
    return body["access_token"], body.get("refresh_token", "")


def token_for_account(refresh_token, account_id):
    """Exchange a refresh token for one scoped to *account_id*."""
    resp = requests.post(
        IAM_URL,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={"grant_type": "refresh_token", "refresh_token": refresh_token,
              "account": account_id},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


# ---------------------------------------------------------------------------
# Schematics
# ---------------------------------------------------------------------------

def list_workspaces(endpoint, token):
    workspaces, offset, limit = [], 0, 100
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


def build_index(endpoint_map, token, regions):
    """[(name_lower, workspace_dict, region)] for every workspace in the account."""
    index, errors = [], []
    for region in regions:
        endpoint = endpoint_map[region]
        try:
            for ws in list_workspaces(endpoint, token):
                index.append(((ws.get("name") or "").lower(), ws, region))
        except Exception as exc:
            errors.append(f"{region}: {str(exc)[:120]}")
    return index, errors


def get_state(endpoint, token, workspace_id, template_id):
    resp = requests.get(
        f"{endpoint}/v1/workspaces/{workspace_id}"
        f"/runtime_data/{template_id}/state_store",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
        timeout=90,
    )
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    return resp.json()


def resource_types(state):
    """Every managed resource type in a state file, one entry per instance."""
    types = []
    for res in (state or {}).get("resources", []):
        if res.get("mode") == "data":
            continue
        rtype = res.get("type")
        if rtype:
            types.extend([rtype] * max(1, len(res.get("instances", []))))
    return types


# ---------------------------------------------------------------------------
# Per-machine check
# ---------------------------------------------------------------------------

def blank(host, sub, acct, found, detail=""):
    return {
        "hostname": host, "subscription_id": sub, "account": acct, "region": "",
        "workspace_name": "", "workspace_id": "", "workspace_status": "",
        "found": found, "match_count": 0, "matched_types": "",
        "resource_total": 0, "detail": detail,
    }


def check_machine(host, sub, acct, index, token, endpoint_map, prefix):
    if not sub:
        return blank(host, sub, acct, "NO_SUB_ID", "no subscription_id in input row")

    needle = sub.lower()
    hits = [(ws, region) for name, ws, region in index if needle in name]

    if not hits:
        return blank(host, sub, acct, "NO_WORKSPACE",
                     "no workspace name contains this subscription_id")

    workspace, region = hits[0]
    note = f"{len(hits)} workspaces matched, used first; " if len(hits) > 1 else ""

    row = blank(host, sub, acct, "NO")
    row.update({
        "region": region,
        "workspace_name": workspace.get("name", ""),
        "workspace_id": workspace.get("id", ""),
        "workspace_status": workspace.get("status", ""),
    })

    templates = workspace.get("template_data") or []
    if not templates:
        row["found"] = "NO_STATE"
        row["detail"] = note + "workspace has no templates"
        return row

    endpoint = endpoint_map[region]
    all_types, had_state = [], False
    for template in templates:
        template_id = template.get("id")
        if not template_id:
            continue
        try:
            state = get_state(endpoint, token, workspace["id"], template_id)
        except Exception as exc:
            msg = str(exc)
            row["found"] = "AUTH_ERROR" if is_auth_error(msg) else "HTTP_ERROR"
            row["detail"] = note + msg[:180]
            return row
        if state is None:
            continue
        had_state = True
        all_types.extend(resource_types(state))

    if not had_state:
        row["found"] = "NO_STATE"
        row["detail"] = note + "no state stored yet"
        return row

    matched = [t for t in all_types if t.startswith(prefix)]
    row["resource_total"] = len(all_types)
    row["match_count"] = len(matched)
    row["matched_types"] = ";".join(sorted(set(matched)))
    row["found"] = "YES" if matched else "NO"
    row["detail"] = note.rstrip("; ")
    row["_types"] = all_types
    return row


# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description="Check each machine's Schematics workspace for a resource type.")
    ap.add_argument("-f", "--file", required=True, help="input CSV")
    ap.add_argument("-o", "--out", help="output CSV (default stdout)")
    ap.add_argument("--api-key", default=os.environ.get("SCHEMATICS_APIKEY"),
                    help="IBM Cloud API key (or set SCHEMATICS_APIKEY)")
    ap.add_argument("--prefix", default="sailpoint_",
                    help="resource type prefix to look for (default sailpoint_)")
    ap.add_argument("--regions", default="us-south,eu-de",
                    help="comma-separated regions (default us-south,eu-de)")
    ap.add_argument("--host-col", default="hostname")
    ap.add_argument("--sub-col", default="subscription_id")
    ap.add_argument("--acct-col", default="account_number")
    ap.add_argument("--list-types", action="store_true",
                    help="also print every resource type seen, with counts")
    args = ap.parse_args()

    if not args.api_key:
        sys.exit("No API key. Pass --api-key or set SCHEMATICS_APIKEY.")

    regions = [r.strip() for r in args.regions.split(",") if r.strip()]
    unknown = [r for r in regions if r not in REGION_ENDPOINTS]
    if unknown:
        sys.exit(f"Unknown region(s): {unknown}. Known: {sorted(REGION_ENDPOINTS)}")

    grouped = read_rows(args.file, args.host_col, args.sub_col, args.acct_col)
    total_rows = sum(len(v) for v in grouped.values())
    print(f"{total_rows} machine(s) across {len(grouped)} account(s)", file=sys.stderr)

    try:
        base_token, refresh_token = get_tokens(args.api_key)
    except Exception as exc:
        sys.exit(f"IAM auth failed: {exc}")

    handle = open(args.out, "w", newline="") if args.out else sys.stdout
    writer = csv.DictWriter(handle, fieldnames=FIELDS, extrasaction="ignore")
    writer.writeheader()

    counts, seen_types = Counter(), Counter()

    try:
        for account, machines in grouped.items():
            # One token and one workspace index per account.
            if account:
                if not refresh_token:
                    for host, sub, acct in machines:
                        writer.writerow(blank(host, sub, acct, "AUTH_ERROR",
                                              "no refresh token; cannot switch account"))
                        counts["AUTH_ERROR"] += 1
                    continue
                try:
                    token = token_for_account(refresh_token, account)
                except Exception as exc:
                    for host, sub, acct in machines:
                        writer.writerow(blank(host, sub, acct, "AUTH_ERROR",
                                              f"cannot scope token: {str(exc)[:150]}"))
                        counts["AUTH_ERROR"] += 1
                    handle.flush()
                    continue
            else:
                token = base_token

            index, errors = build_index(REGION_ENDPOINTS, token, regions)
            label = account or "default"
            print(f"[{label}] indexed {len(index)} workspace(s)"
                  + (f"; errors: {errors}" if errors else ""), file=sys.stderr)

            if not index and errors:
                kind = "AUTH_ERROR" if any(is_auth_error(e) for e in errors) \
                    else "HTTP_ERROR"
                for host, sub, acct in machines:
                    writer.writerow(blank(host, sub, acct, kind, "; ".join(errors)[:180]))
                    counts[kind] += 1
                handle.flush()
                continue

            for host, sub, acct in machines:
                row = check_machine(host, sub, acct, index, token,
                                    REGION_ENDPOINTS, args.prefix)
                seen_types.update(row.pop("_types", []))
                counts[row["found"]] += 1
                writer.writerow(row)
                handle.flush()
    finally:
        if args.out:
            handle.close()

    if args.list_types:
        print("\nresource types seen:", file=sys.stderr)
        for rtype, n in seen_types.most_common():
            print(f"  {n:6d}  {rtype}", file=sys.stderr)

    summary = "  ".join(f"{k}={v}" for k, v in sorted(counts.items()))
    print(f"\n{total_rows} machine(s):  {summary}", file=sys.stderr)
    sys.exit(0 if counts.get("YES") else 1)


if __name__ == "__main__":
    main()
