import csv
import datetime
import json
import re
import requests

# Base endpoint configuration for MSRC CVRF API v3.0
BASE_URL = "https://api.msrc.microsoft.com/cvrf/v3.0"
HEADERS = {"Accept": "application/json"}


def get_validated_api_months():
    """Queries the live Microsoft API directory to pull exact, valid month IDs

    matching the past two months. This avoids broken/empty manual date strings.
    """
    url = f"{BASE_URL}/updates"
    print("[*] Contacting Microsoft MSRC directory to find valid update IDs...")

    try:
        response = requests.get(url, headers=HEADERS)
        if response.status_code != 200:
            print(
                f"[!] Failed to pull directory. Status Code: {response.status_code}"
            )
            return []

        data = response.json()
        updates = data.get("value", []) if isinstance(data, dict) else data

        available_months = []
        for item in updates:
            if isinstance(item, dict) and "ID" in item:
                available_months.append(item["ID"])

        def parse_ms_date(id_str):
            try:
                return datetime.datetime.strptime(id_str, "%Y-%b")
            except ValueError:
                return datetime.datetime.min

        available_months.sort(key=parse_ms_date, reverse=True)

        validated_targets = available_months[:2]
        print(f"[+] Found live Microsoft update IDs to scrape: {validated_targets}")
        return validated_targets

    except Exception as e:
        print(f"[!] Error reading live directory: {e}")
        current_date = datetime.datetime.now()
        fallback = []
        for i in range(2):
            target = current_date - datetime.timedelta(days=i * 30)
            fallback.append(target.strftime("%Y-%b"))
        return fallback


def fetch_cvrf_document(month_id):
    """Fetches the complete security bulletin JSON document for a month ID."""
    url = f"{BASE_URL}/cvrf/{month_id}"
    print(f"[*] Fetching vulnerability details for {month_id}...")
    try:
        response = requests.get(url, headers=HEADERS)
        if response.status_code == 200:
            return response.json()
        print(
            f"[!] Microsoft API responded with status {response.status_code} for {month_id}"
        )
        return None
    except Exception as e:
        print(f"[!] Error fetching {month_id}: {e}")
        return None


def parse_vulnerabilities(cvrf_data):
    """Parses vulnerabilities out into a flat list format for Grafana tables."""
    extracted_cves = []
    if not cvrf_data:
        return extracted_cves

    product_map = {}
    product_tree = cvrf_data.get("ProductTree", {})
    if isinstance(product_tree, dict) and "FullProductName" in product_tree:
        prod_list = product_tree["FullProductName"]
        if isinstance(prod_list, list):
            for prod in prod_list:
                if isinstance(prod, dict) and "ProductID" in prod:
                    product_map[prod["ProductID"]] = prod.get("Value", "Unknown")

    vulnerabilities = cvrf_data.get("Vulnerability", [])
    if not isinstance(vulnerabilities, list):
        return extracted_cves

    for vuln in vulnerabilities:
        if not isinstance(vuln, dict):
            continue

        cve_num = vuln.get("CVE", "N/A")

        # 1. Base CVSS Score & Vector Extraction
        cvss_score = "N/A"
        cvss_vector = "N/A"

        cvss_score_sets_raw = vuln.get("CVSSScoreSets", {})
        score_sets = []

        if isinstance(cvss_score_sets_raw, dict):
            raw_set = cvss_score_sets_raw.get("ScoreSet", [])
            score_sets = raw_set if isinstance(raw_set, list) else [raw_set]
        elif isinstance(cvss_score_sets_raw, list):
            score_sets = cvss_score_sets_raw

        for score_set in score_sets:
            if not isinstance(score_set, dict):
                continue
            vector = score_set.get("Vector", "") or ""
            if "CVSS:3.1" in vector or "CVSS" in vector:
                cvss_score = score_set.get("BaseScore", "N/A")
                cvss_vector = vector
                break

        # 2. Extract Original Release Date
        release_date = "N/A"
        revision_history_raw = vuln.get("RevisionHistory", {})
        revision_history = []
        if isinstance(revision_history_raw, dict):
            revision_history = revision_history_raw.get("Revision", [])
            if isinstance(revision_history, dict):
                revision_history = [revision_history]
        elif isinstance(revision_history_raw, list):
            revision_history = revision_history_raw

        if revision_history and len(revision_history) > 0:
            first_rev = revision_history[0]
            if isinstance(first_rev, dict):
                release_date = first_rev.get("Date", "N/A")

        # 3. Compile Affected Products (Using newline separator for Grafana)
        affected_products = []
        product_statuses_raw = vuln.get("ProductStatuses", {})
        status_set = []
        if isinstance(product_statuses_raw, dict):
            status_set = product_statuses_raw.get("Status", [])
            if isinstance(status_set, dict):
                status_set = [status_set]
        elif isinstance(product_statuses_raw, list):
            status_set = product_statuses_raw

        for status in status_set:
            if not isinstance(status, dict):
                continue
            prod_ids = status.get("ProductID", [])
            if isinstance(prod_ids, str):
                prod_ids = [prod_ids]

            if isinstance(prod_ids, list):
                for prod_id in prod_ids:
                    if prod_id in product_map:
                        affected_products.append(product_map[prod_id])

        unique_products = list(dict.fromkeys(affected_products)) if affected_products else ["N/A"]
        products_flat_str = "\n".join(unique_products)

        # 4. Exploitability Ratings (Removing duplicates while preserving order)
        exploitability_info = []
        threats_raw = vuln.get("Threats", {})
        threats = []
        if isinstance(threats_raw, dict):
            threats = threats_raw.get("Threat", [])
            if isinstance(threats, dict):
                threats = [threats]
        elif isinstance(threats_raw, list):
            threats = threats_raw

        for threat in threats:
            if not isinstance(threat, dict):
                continue
            if str(threat.get("Type")) == "0" or threat.get("Type") == 0:
                description_obj = threat.get("Description", {})
                description = (
                    description_obj.get("Value", "")
                    if isinstance(description_obj, dict)
                    else ""
                )
                if description and description not in exploitability_info:
                    exploitability_info.append(description)

        exploitability_str = (
            " | ".join(exploitability_info) if exploitability_info else "N/A"
        )

        # 5. Robust FAQ and Notes Extraction
        faq_content = []
        notes_raw = vuln.get("Notes", {})
        notes = []
        if isinstance(notes_raw, dict):
            notes = notes_raw.get("Note", [])
            if isinstance(notes, dict):
                notes = [notes]
        elif isinstance(notes_raw, list):
            notes = notes_raw

        for note in notes:
            if not isinstance(note, dict):
                continue
            note_type = str(note.get("Type", ""))
            title = note.get("Title", "") or ""

            if note_type in ["5", "6"] or "faq" in title.lower():
                faq_text = note.get("Value", "")
                if faq_text:
                    clean_faq = re.sub(r"<[^<]+?>", "", faq_text).strip()
                    clean_faq = re.sub(r"\s+", " ", clean_faq)
                    if clean_faq:
                        faq_content.append(
                            f"[{title}] {clean_faq}" if title else clean_faq
                        )

        faq_str = " | ".join(faq_content) if faq_content else "No FAQ available."

        # Corrected URL structures
        update_guide_url = f"https://msrc.microsoft.com/update-guide/vulnerability/{cve_num}"
        update_guide_url_href = f'<a href="{update_guide_url}" target="_blank">{cve_num}</a>'

        try:
            final_score = float(cvss_score)
        except ValueError:
            final_score = 0.0

        extracted_cves.append(
            {
                "cve_number": cve_num,
                "cvss_score": final_score,
                "cvss_vector": cvss_vector if cvss_vector else "N/A",
                "release_date": release_date,
                "affected_products": products_flat_str,
                "update_guide_url": update_guide_url,
                "update_guide_url_href": update_guide_url_href,
                "exploitability": exploitability_str,
                "faq": faq_str,
            }
        )

    return extracted_cves


def main():
    target_months = get_validated_api_months()
    all_collected_data = []

    for month in target_months:
        cvrf_doc = fetch_cvrf_document(month)
        if cvrf_doc:
            month_cves = parse_vulnerabilities(cvrf_doc)
            all_collected_data.extend(month_cves)
            print(f"[+] Successfully parsed {len(month_cves)} CVEs from {month}.")

    json_output_file = "msrc_cves_grafana.json"
    with open(json_output_file, "w", encoding="utf-8") as f:
        json.dump(all_collected_data, f, indent=2, ensure_ascii=False)

    csv_output_file = "msrc_cves_grafana.csv"
    if all_collected_data:
        keys = all_collected_data[0].keys()

        with open(csv_output_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=keys)
            writer.writeheader()
            writer.writerows(all_collected_data)
        print(f"\n[++] Success! Total records generated: {len(all_collected_data)}")
        print(f"    -> JSON Output Path: {json_output_file}")
        print(f"    -> CSV Output Path:  {csv_output_file}")


if __name__ == "__main__":
    main()
