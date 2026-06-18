#!/usr/bin/env python3
"""
Website audit using DNS + TCP port checks.

Network policy in this environment:
  - DNS resolution: real results (authoritative)
  - TCP port 443/80: real connectivity (confirms server is listening)
  - SSL certs: shows proxy's cert (unreliable — excluded from output)
  - HTTP responses: intercepted by proxy (excluded from output)

Classification:
  Active        — DNS resolves + TCP port 443 or 80 accepts connections
  DNS Failure   — Domain does not resolve (likely defunct/expired)
  Timeout       — DNS ok but both ports time out (server unreachable)
  Manual Review — Missing URL or inconclusive
"""
import socket, ssl, time
from datetime import datetime
import pandas as pd
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

TIMEOUT = 10

STATUS_COLORS = {
    "Active": "C6EFCE",
    "Active - Redirected": "FFEB9C",
    "DNS Failure": "FFC7CE",
    "Timeout": "FFCCAA",
    "SSL Error": "FFD9B3",
    "404 Error": "FFC7CE",
    "Server Error": "FF9999",
    "Parked Domain": "E2EFDA",
    "Abandoned/Outdated": "D9E1F2",
    "Manual Review Required": "F2CCFF",
}


def normalize_url(u):
    u = u.strip()
    if not u:
        return ""
    if not u.startswith(("http://", "https://")):
        u = "https://" + u
    return u


def check_dns(host):
    try:
        socket.getaddrinfo(host, None)
        return True, ""
    except socket.gaierror as e:
        return False, str(e)


def check_tcp(host, port):
    t0 = time.monotonic()
    try:
        s = socket.create_connection((host, port), timeout=TIMEOUT)
        lat = round((time.monotonic() - t0) * 1000)
        s.close()
        return True, lat, ""
    except socket.timeout:
        return False, round((time.monotonic() - t0) * 1000), "TIMEOUT"
    except ConnectionRefusedError:
        return False, round((time.monotonic() - t0) * 1000), "REFUSED"
    except OSError as e:
        return False, round((time.monotonic() - t0) * 1000), str(e)


def parse_host(url):
    import urllib.parse
    prs = urllib.parse.urlparse(url)
    return prs.hostname or "", prs.scheme == "https"


def audit_url(raw):
    url = normalize_url(raw)
    base = dict(
        normalized_url="", http_status_code="N/A",
        response_time_ms=None, website_status="", issue_type="", notes="",
        dns_resolved="", port_443="", port_80="",
    )

    if not url:
        d = base.copy(); d.update(website_status='Manual Review Required',
                issue_type="Missing URL", notes="No URL in source data"); return d

    host, use_ssl = parse_host(url)
    base["normalized_url"] = url

    # DNS
    t0 = time.monotonic()
    dns_ok, dns_err = check_dns(host)
    if not dns_ok:
        r = base.copy()
        r.update(website_status="DNS Failure", issue_type="DNS Resolution Failed",
                 notes=f"Domain does not resolve: {dns_err}", dns_resolved="No",
                 port_443="N/A", port_80="N/A",
                 response_time_ms=round((time.monotonic() - t0) * 1000))
        return r

    base["dns_resolved"] = "Yes"

    # TCP 443
    tcp443_ok, tcp443_lat, tcp443_err = check_tcp(host, 443)
    base["port_443"] = "Open" if tcp443_ok else f"Closed/Timeout ({tcp443_err})"

    # TCP 80
    tcp80_ok, tcp80_lat, tcp80_err = check_tcp(host, 80)
    base["port_80"] = "Open" if tcp80_ok else f"Closed/Timeout ({tcp80_err})"

    # Response time = fastest successful connection
    lats = [l for ok, l in [(tcp443_ok, tcp443_lat), (tcp80_ok, tcp80_lat)] if ok]
    response_time = min(lats) if lats else max(tcp443_lat, tcp80_lat)
    base["response_time_ms"] = response_time

    # Classify
    if not tcp443_ok and not tcp80_ok:
        err_detail = f"Port 443: {tcp443_err} | Port 80: {tcp80_err}"
        if "TIMEOUT" in tcp443_err or "TIMEOUT" in tcp80_err:
            ws, it = "Timeout", "Connection Timeout"
            notes = f"Server unreachable — {err_detail}"
        else:
            ws, it = "Manual Review Required", "Connection Refused"
            notes = f"Both ports closed — {err_detail}"
        r = base.copy(); r.update(website_status=ws, issue_type=it, notes=notes); return r

    # Active
    ports_open = []
    if tcp443_ok:
        ports_open.append(f"443 ({tcp443_lat}ms)")
    if tcp80_ok:
        ports_open.append(f"80 ({tcp80_lat}ms)")
    notes = (f"Domain resolves and server accepts connections on port(s): {', '.join(ports_open)}. "
             "Full HTTP checks (redirects, 404, parked domain, content) require unrestricted network access.")
    r = base.copy(); r.update(website_status="Active", issue_type="", notes=notes); return r


def main():
    src = "/home/user/MedSpa/MedSpa Leads onyamarketing .xlsx"
    out = "/home/user/MedSpa/onyamarketing_website_audit.xlsx"

    print("Reading source file...")
    df = pd.read_excel(src, dtype=str).fillna("")
    url_col = "Company URL"
    total = len(df)

    records = []
    for i, row in df.iterrows():
        raw = str(row[url_col]).strip()
        print(f"[{i+1}/{total}  {(i+1)/total*100:.0f}%] {raw[:70]}", flush=True)
        records.append(audit_url(raw))

    res = pd.DataFrame(records)
    out_df = df.copy()
    out_df["Normalized URL"]     = res["normalized_url"].values
    out_df["HTTP Status Code"]   = res["http_status_code"].values
    out_df["Response Time (ms)"] = res["response_time_ms"].values
    out_df["Website Status"]     = res["website_status"].values
    out_df["Issue Type"]         = res["issue_type"].values
    out_df["Notes"]              = res["notes"].values
    out_df["DNS Resolved"]       = res["dns_resolved"].values
    out_df["Port 443"]           = res["port_443"].values
    out_df["Port 80"]            = res["port_80"].values

    sc = out_df["Website Status"].value_counts()
    cats = [
        ("Total websites reviewed", total),
        ("Active",                  int(sc.get("Active", 0))),
        ("Active - Redirected",     int(sc.get("Active - Redirected", 0))),
        ("DNS Failure",             int(sc.get("DNS Failure", 0))),
        ("Timeout",                 int(sc.get("Timeout", 0))),
        ("SSL Error",               int(sc.get("SSL Error", 0))),
        ("404 Error",               int(sc.get("404 Error", 0))),
        ("Server Error",            int(sc.get("Server Error", 0))),
        ("Parked Domain",           int(sc.get("Parked Domain", 0))),
        ("Abandoned/Outdated",      int(sc.get("Abandoned/Outdated", 0))),
        ("Manual Review Required",  int(sc.get("Manual Review Required", 0))),
    ]
    sum_df = pd.DataFrame(cats, columns=["Category", "Count"])

    # Methodology note sheet
    method_rows = [
        ("Audit Date", datetime.utcnow().strftime("%Y-%m-%d")),
        ("Total Records", total),
        ("", ""),
        ("Checks Performed", ""),
        ("DNS Resolution", "Authoritative — confirms domain exists and resolves"),
        ("TCP Port 443", "Direct socket connect — confirms HTTPS server is listening"),
        ("TCP Port 80", "Direct socket connect — confirms HTTP server is listening"),
        ("", ""),
        ("Checks NOT Available in This Environment", ""),
        ("HTTP Response Content", "Blocked by network egress policy (proxy returns 403)"),
        ("SSL Certificate Details", "Proxy performs TLS inspection; cert shown is proxy cert, not origin"),
        ("Redirect Following", "Requires HTTP response body — blocked"),
        ("Parked Domain Detection", "Requires HTTP page content — blocked"),
        ("Abandoned Site Detection", "Requires HTTP page content — blocked"),
        ("", ""),
        ("Recommendation", "Run full_audit_http.py from an unrestricted network to get HTTP-level details"),
    ]
    method_df = pd.DataFrame(method_rows, columns=["Check", "Detail"])

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Audit Results", index=False)
        sum_df.to_excel(writer, sheet_name="Summary", index=False)
        method_df.to_excel(writer, sheet_name="Methodology", index=False)

        ws  = writer.sheets["Audit Results"]
        ws2 = writer.sheets["Summary"]
        ws3 = writer.sheets["Methodology"]

        si = list(out_df.columns).index("Website Status") + 1
        for ri in range(2, len(out_df) + 2):
            c = ws.cell(row=ri, column=si)
            col = STATUS_COLORS.get(str(c.value or ""), "FFFFFF")
            c.fill = PatternFill(start_color=col, end_color=col, fill_type="solid")

        for sheet in (ws, ws2, ws3):
            for cell in sheet[1]:
                cell.font = Font(bold=True)
            for col in sheet.columns:
                w = max((len(str(c.value or "")) for c in col), default=10)
                sheet.column_dimensions[get_column_letter(col[0].column)].width = min(w + 2, 70)

    print(f"\nSaved to {out}\n=== SUMMARY ===")
    for lbl, cnt in cats:
        print(f"  {lbl}: {cnt}")


if __name__ == "__main__":
    main()
