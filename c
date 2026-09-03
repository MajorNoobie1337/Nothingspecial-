#!/usr/bin/env python3
"""Per-machine check: does this host's Schematics workspace contain a resource type?

Input CSV: hostname, subscription_id, account_id (other columns ignored).
Auth: one API key from $SCHEMATICS_APIKEY, used for every account.

The workspace is found by locating the one whose *name contains the
subscription_id*. When several match, the account_id from the CSV picks the
right one -- compared against the account embedded in the workspace CRN. The
workspace's state is then read through the Schematics API and searched for
resources whose type starts with --prefix.

Usage:
    export SCHEMATICS_APIKEY=...
    python check_workspace_resource.py -f prod_check.csv -o sailpoint.csv
    python check_workspace_resource.py -f prod_check.csv --prefix sailpoint_
    python check_workspace_resource.py -f prod_check.csv --list-types
    python check_workspace_resource.py -f prod_check.csv --dump-index ws.csv

found:
    YES           at least one resource of the wanted type
    NO            state read fine, no such resource
    NO_STATE      workspace exists but has no state yet
    NO_WORKSPACE  no workspace name contains the subscription_id
    NO_SUB_ID     input row had no subscription_id
    AUTH_ERROR    token rejected -> result meaningless
    HTTP_ERROR    anything else -> result meaningless, see detail
"""

import argparse
import csv
import os
import re
import sys
import time
from collections import Counter

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

IAM_URL = "https://iam.cloud.ibm.com/identity/token"
APIKEY_GRANT = "urn:ibm:params:oauth:grant-type:apikey"
TOKEN_TTL = 45 * 60  # refresh well before the usual 60 minute expiry

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
    "hostname", "subscription_id", "account_id", "workspace_account", "region",
    "workspace_name", "workspace_id", "workspace_status",
    "found", "match_count", "matched_types", "resource_total", "detail",
]

AUTH_MARKERS = ("401", "403", "Unauthorized", "invalid_grant", "Forbidden")
CRN_ACCOUNT = re.compile(r":a/([0-9a-fA-F]+):")


def is_auth_error(msg):
    return any(m in msg for m in AUTH_MARKERS)


def account_from_crn(crn):
    """Pull the account id out of crn:v1:bluemix:public:...:a/<id>:..."""
    match = CRN_ACCOUNT.search(crn or "")
    return match.group(1) if match else ""


# ---------------------------------------------------------------------------
# Auth: one key, one token, refreshed on age
# ---------------------------------------------------------------------------

class Token:
    def __init__(self, api_key):
        self.api_key = api_key
        self._value = ""
        self._fetched = 0.0

    def get(self):
        if not self._value or (time.time() - self._fetched) > TOKEN_TTL:
            resp = requests.post(
                IAM_URL,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                data={"grant_type": APIKEY_GRANT, "apikey": self.api_key},
                timeout=30,
            )
            resp.raise_for_status()
            self._value = resp.json()["access_token"]
            self._fetched = time.time()
        return self._value

    def headers(self):
        return {"Authorization": f"Bearer {self.get()}",
                "Accept": "application/json"}


# ---------------------------------------------------------------------------
# Input
# ---------------------------------------------------------------------------

def read_rows(path, host_col, sub_col, acct_col):
    with open(path, newline="") as fh:
        reader = csv.DictReader(fh)
        header = reader.fieldnames or []
        missing = [c for c in (host_col, sub_col, acct_col)
                   if c and c not in header]
        if missing:
            sys.exit(f"Input CSV is missing column(s): {missing}\n"
                     f"Header found: {header}")
        rows = [
            ((r.get(host_col) or "").strip(),
             (r.get(sub_col) or "").strip(),
             (r.get(acct_col) or "").strip() if acct_col else "")
            for r in reader
        ]
    return [r for r in rows if r[0] or r[1]]


# ---------------------------------------------------------------------------
# Schematics
# ---------------------------------------------------------------------------

def list_workspaces(endpoint, token):
    workspaces, offset, limit = [], 0, 100
    while True:
        resp = requests.get(f"{endpoint}/v1/workspaces", headers=token.headers(),
                            params={"limit": limit, "offset": offset}, timeout=60)
        resp.raise_for_status()
        body = resp.json()
        page = body.get("workspaces", [])
        workspaces.extend(page)
        offset += limit
        if len(page) < limit or offset >= body.get("total_count", 0):
            break
    return workspaces


def build_index(token, regions):
    """[(name_lower, workspace, region)] for every workspace the key can see."""
    index, errors, seen = [], [], set()
    for region in regions:
        try:
            found = list_workspaces(REGION_ENDPOINTS[region], token)
        except Exception as exc:
            msg = f"{region}: {str(exc)[:120]}"
            errors.append(msg)
            print(f"  {msg}", file=sys.stderr)
            continue
        added = 0
        for ws in found:
            ws_id = ws.get("id")
            if ws_id in seen:
                continue
            seen.add(ws_id)
            index.append(((ws.get("name") or "").lower(), ws, region))
            added += 1
        print(f"  {region}: {added} workspace(s)", file=sys.stderr)
    return index, errors


def get_state(endpoint, token, workspace_id, template_id):
    resp = requests.get(
        f"{endpoint}/v1/workspaces/{workspace_id}"
        f"/runtime_data/{template_id}/state_store",
        headers=token.headers(), timeout=90,
    )
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    return resp.json()


def resource_types(state):
    """Managed resource types in a state file, one entry per instance."""
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
        "hostname": host, "subscription_id": sub, "account_id": acct,
        "workspace_account": "", "region": "", "workspace_name": "",
        "workspace_id": "", "workspace_status": "", "found": found,
        "match_count": 0, "matched_types": "", "resource_total": 0,
        "detail": detail,
    }


def pick_workspace(hits, acct):
    """Choose among workspaces matching the subscription_id, preferring the
    one whose CRN account matches the account_id from the input row."""
    notes = []
    if acct:
        same = [h for h in hits
                if account_from_crn(h[0].get("crn", "")).lower() == acct.lower()]
        if same:
            if len(hits) > 1:
                notes.append(f"{len(hits)} matched, {len(same)} in account {acct}")
            if len(same) > 1:
                notes.append("several in the same account, used first")
            return same[0], notes
        if hits:
            found_accts = sorted({account_from_crn(h[0].get("crn", "")) or "?"
                                  for h in hits})
            notes.append(f"no workspace in account {acct}; "
                         f"found in {', '.join(found_accts)}, used first")
            return hits[0], notes
    if len(hits) > 1:
        notes.append(f"{len(hits)} workspaces matched, used first")
    return hits[0], notes


def check_machine(host, sub, acct, index, token, prefix):
    if not sub:
        return blank(host, sub, acct, "NO_SUB_ID", "no subscription_id in input row")

    needle = sub.lower()
    hits = [(ws, region) for name, ws, region in index if needle in name]
    if not hits:
        return blank(host, sub, acct, "NO_WORKSPACE",
                     "no workspace name contains this subscription_id")

    (workspace, region), notes = pick_workspace(hits, acct)

    row = blank(host, sub, acct, "NO")
    row.update({
        "workspace_account": account_from_crn(workspace.get("crn", "")),
        "region": region,
        "workspace_name": workspace.get("name", ""),
        "workspace_id": workspace.get("id", ""),
        "workspace_status": workspace.get("status", ""),
    })

    templates = workspace.get("template_data") or []
    if not templates:
        row["found"] = "NO_STATE"
        row["detail"] = "; ".join(notes + ["workspace has no templates"])
        return row

    endpoint = REGION_ENDPOINTS[region]
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
            row["detail"] = "; ".join(notes + [msg[:180]])
            return row
        if state is None:
            continue
        had_state = True
        all_types.extend(resource_types(state))

    if not had_state:
        row["found"] = "NO_STATE"
        row["detail"] = "; ".join(notes + ["no state stored yet"])
        return row

    matched = [t for t in all_types if t.startswith(prefix)]
    row["resource_total"] = len(all_types)
    row["match_count"] = len(matched)
    row["matched_types"] = ";".join(sorted(set(matched)))
    row["found"] = "YES" if matched else "NO"
    row["detail"] = "; ".join(notes)
    row["_types"] = all_types
    return row


# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description="Check each machine's Schematics workspace for a resource type.")
    ap.add_argument("-f", "--file", required=True, help="input CSV")
    ap.add_argument("-o", "--out", help="output CSV (default stdout)")
    ap.add_argument("--api-key", default=os.environ.get("SCHEMATICS_APIKEY"),
                    help="IBM Cloud API key (default: $SCHEMATICS_APIKEY)")
    ap.add_argument("--prefix", default="sailpoint_",
                    help="resource type prefix to look for (default sailpoint_)")
    ap.add_argument("--regions", default=",".join(REGION_ENDPOINTS),
                    help="comma-separated regions (default: all known)")
    ap.add_argument("--host-col", default="hostname")
    ap.add_argument("--sub-col", default="subscription_id")
    ap.add_argument("--acct-col", default="account_id",
                    help="account column, used to disambiguate; '' to ignore")
    ap.add_argument("--list-types", action="store_true",
                    help="also print every resource type seen, with counts")
    ap.add_argument("--dump-index", metavar="CSV",
                    help="write the workspace index here and exit")
    args = ap.parse_args()

    if not args.api_key:
        sys.exit("No API key. Pass --api-key or set SCHEMATICS_APIKEY.")

    regions = [r.strip() for r in args.regions.split(",") if r.strip()]
    unknown = [r for r in regions if r not in REGION_ENDPOINTS]
    if unknown:
        sys.exit(f"Unknown region(s): {unknown}. Known: {sorted(REGION_ENDPOINTS)}")

    token = Token(args.api_key)
    try:
        token.get()
    except Exception as exc:
        sys.exit(f"IAM auth failed: {exc}")

    print("indexing workspaces...", file=sys.stderr)
    index, errors = build_index(token, regions)
    accounts_seen = {account_from_crn(ws.get("crn", "")) for _, ws, _ in index}
    accounts_seen.discard("")
    print(f"{len(index)} workspace(s) across {len(accounts_seen)} account(s)",
          file=sys.stderr)

    if args.dump_index:
        with open(args.dump_index, "w", newline="") as fh:
            dump = csv.writer(fh)
            dump.writerow(["name", "id", "status", "region", "account"])
            for _, ws, region in index:
                dump.writerow([ws.get("name", ""), ws.get("id", ""),
                               ws.get("status", ""), region,
                               account_from_crn(ws.get("crn", ""))])
        print(f"wrote {args.dump_index}", file=sys.stderr)
        return

    if not index:
        sys.exit("No workspaces visible. " + ("; ".join(errors) if errors else
                 "Check the key's permissions and --regions."))

    rows = read_rows(args.file, args.host_col, args.sub_col, args.acct_col)
    print(f"{len(rows)} machine(s) to check", file=sys.stderr)

    handle = open(args.out, "w", newline="") if args.out else sys.stdout
    writer = csv.DictWriter(handle, fieldnames=FIELDS, extrasaction="ignore")
    writer.writeheader()

    counts, seen_types = Counter(), Counter()
    try:
        for host, sub, acct in rows:
            row = check_machine(host, sub, acct, index, token, args.prefix)
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
    print(f"\n{len(rows)} machine(s):  {summary}", file=sys.stderr)
    if errors:
        print(f"region errors: {errors}", file=sys.stderr)
    sys.exit(0 if counts.get("YES") else 1)


if __name__ == "__main__":
    main()
