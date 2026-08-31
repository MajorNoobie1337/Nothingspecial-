#!/usr/bin/env python3
"""Check a list of hostnames against Orchestrator PROD only, output CSV.

Unlike main.py this does not fall back to preprod/int, does not consult
Astro/Reftec, and never skips a host based on cloud_type. Every hostname
gets exactly one prod State Manager query, plus an optional DB query.

Usage:
    python check_prod.py hosts.txt -o prod_check.csv
    python check_prod.py hosts.txt --db          # also try the reader/DB endpoint
    python check_prod.py hosts.txt --raw         # dump matched records to stderr

CSV columns:
    hostname        the host as it appeared in your input file
    status          FOUND / PARTIAL / NOT_FOUND / AUTH_ERROR / HTTP_ERROR
    subscription_id uuid from state manager, or subscription_id from the DB
    matches         how many returned rows matched the hostname exactly
    matched_field   which field in the record held the hostname
    returned_rows   how many rows the filter returned in total
    sub_status      the subscription's own status field
    source          state_manager / db
    detail          error text or note; empty when clean

Statuses:
    FOUND        exact hostname match in prod
    PARTIAL      filter returned rows, but none match the hostname exactly
    NOT_FOUND    filter returned zero rows -> genuinely absent from prod
    AUTH_ERROR   401/403/keycloak -> result is meaningless, fix creds
    HTTP_ERROR   anything else -> result is meaningless, see detail
"""

import argparse
import csv
import json
import sys

import requests
import urllib3

# Adjust these two imports to match your layout if needed.
from orchestrator import OrchestratorClient
from main import load_config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

ENV = "prod"
DEFAULT_GATEWAY = "https://orchestrator-gw.group.echonet"
AUTH_MARKERS = ("401", "403", "Unauthorized", "invalid_client", "Keycloak auth failed")

FIELDS = [
    "hostname", "status", "subscription_id", "matches", "matched_field",
    "returned_rows", "sub_status", "source", "detail",
]


def is_auth_error(msg):
    return any(m in msg for m in AUTH_MARKERS)


def matching_field(record, hostname):
    """Return the key whose value equals the hostname, or None."""
    target = hostname.strip().lower()
    for key, value in record.items():
        if isinstance(value, str) and value.strip().lower() == target:
            return key
    return None


def subscription_id_of(record):
    for key in ("uuid", "subscription_id", "id"):
        value = record.get(key)
        if value:
            return value
    return ""


def query_state_manager(client, gateway, hostname):
    url = (
        f"{gateway}/state_manager/api/v1/subscriptions"
        f"?page=1&size=100&filters={hostname}"
    )
    resp = requests.get(url, verify=False, headers=client._auth_headers(ENV))
    resp.raise_for_status()
    return resp.json().get("subscriptions", [])


def query_db(client, gateway, hostname):
    url = f"{gateway}/reader/api/v1/products/server/{hostname}/views/server-details"
    resp = requests.get(url, verify=False, headers=client._auth_headers(ENV))
    resp.raise_for_status()
    return resp.json()


def blank_row(hostname, status, detail=""):
    return {
        "hostname": hostname, "status": status, "subscription_id": "",
        "matches": 0, "matched_field": "", "returned_rows": 0,
        "sub_status": "", "source": "", "detail": detail,
    }


def check_host(client, gateway, hostname, want_db):
    """Return (csv_row_dict, matched_record_or_None)."""
    try:
        subs = query_state_manager(client, gateway, hostname)
    except Exception as exc:
        msg = str(exc)
        kind = "AUTH_ERROR" if is_auth_error(msg) else "HTTP_ERROR"
        return blank_row(hostname, kind, msg[:200]), None

    hits = [(rec, matching_field(rec, hostname)) for rec in subs]
    exact = [(rec, field) for rec, field in hits if field]

    if exact:
        record, field = exact[0]
        row = blank_row(hostname, "FOUND")
        row.update({
            "subscription_id": subscription_id_of(record),
            "matches": len(exact),
            "matched_field": field,
            "returned_rows": len(subs),
            "sub_status": record.get("status", ""),
            "source": "state_manager",
            "detail": "more than one exact match" if len(exact) > 1 else "",
        })
        return row, record

    if subs and not want_db:
        row = blank_row(hostname, "PARTIAL")
        row.update({
            "returned_rows": len(subs),
            "source": "state_manager",
            "detail": "filter matched rows, none exact",
        })
        return row, subs[0]

    if not want_db:
        row = blank_row(hostname, "NOT_FOUND")
        row["source"] = "state_manager"
        return row, None

    # State manager gave nothing usable; try the DB endpoint before concluding.
    try:
        data = query_db(client, gateway, hostname)
    except Exception as exc:
        msg = str(exc)
        if is_auth_error(msg):
            return blank_row(hostname, "AUTH_ERROR", msg[:200]), None
        row = blank_row(hostname, "PARTIAL" if subs else "NOT_FOUND")
        row.update({
            "returned_rows": len(subs),
            "source": "state_manager",
            "detail": f"DB fallback failed: {msg[:150]}",
        })
        return row, (subs[0] if subs else None)

    row = blank_row(hostname, "FOUND")
    row.update({
        "subscription_id": subscription_id_of(data),
        "matches": 1,
        "matched_field": matching_field(data, hostname) or "",
        "returned_rows": len(subs),
        "sub_status": data.get("status", ""),
        "source": "db",
    })
    return row, data


def main():
    ap = argparse.ArgumentParser(description="Verify hosts against Orchestrator prod.")
    ap.add_argument("file", help="file with one hostname per line")
    ap.add_argument("-o", "--out", help="CSV file to write (default: stdout)")
    ap.add_argument("--db", action="store_true",
                    help="also try the reader/DB endpoint when state manager misses")
    ap.add_argument("--raw", action="store_true",
                    help="dump matched records to stderr for inspection")
    args = ap.parse_args()

    loader = load_config()
    keycloak_config = loader.get("keycloak", {})
    if ENV not in keycloak_config:
        sys.exit(f"No '{ENV}' entry in keycloak config; found: "
                 f"{list(keycloak_config)}")

    client = OrchestratorClient(keycloak_config)
    gateway = keycloak_config.get(ENV, {}).get("gateway_url", DEFAULT_GATEWAY)

    with open(args.file) as fh:
        hostnames = [line.strip() for line in fh if line.strip()]

    handle = open(args.out, "w", newline="") if args.out else sys.stdout
    counts = {}
    try:
        writer = csv.DictWriter(handle, fieldnames=FIELDS)
        writer.writeheader()
        for hostname in hostnames:
            row, record = check_host(client, gateway, hostname, args.db)
            counts[row["status"]] = counts.get(row["status"], 0) + 1
            writer.writerow(row)
            handle.flush()
            if args.raw and record:
                print(f"--- {hostname} ---\n{json.dumps(record, indent=2)}",
                      file=sys.stderr)
    finally:
        if args.out:
            handle.close()

    summary = "  ".join(f"{k}={v}" for k, v in sorted(counts.items()))
    print(f"\n{len(hostnames)} host(s):  {summary}", file=sys.stderr)

    # Non-zero exit if anything is unverified, so this can gate a pipeline.
    unverified = sum(v for k, v in counts.items() if k != "FOUND")
    sys.exit(1 if unverified else 0)


if __name__ == "__main__":
    main()
