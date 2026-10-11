from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
import json
import re
from bs4 import BeautifulSoup
import feedparser
import requests

# RSS Feed URL for Oracle Security Alerts
RSS_URL = "https://www.oracle.com/ocom/groups/public/@otn/documents/webcontent/rss-otn-sec.xml"

# Set to True to generate the Grafana-friendly flat JSON file alongside nested output
GENERATE_FLAT_OUTPUT = True


def clean_text(text):
  """Cleans up non-breaking spaces, multi-spaces, and leading/trailing whitespace."""
  if not text:
    return ""
  text = text.replace("\xa0", " ").replace("&nbsp;", " ")
  return re.sub(r"\s+", " ", text).strip()


def normalize_header(header_text):
  """Unifies variant names, enforces specific names, and replaces spaces with underscores."""
  if not header_text:
    return ""

  # Remove hyphenation breaks caused by table wrapping
  header_text = re.sub(r"(\w+)-\s+(\w+)", r"\1\2", header_text)
  header_text = header_text.replace("- ", "")

  # Strip CVSS version prefix
  header_text = re.sub(
      r"(?i)CVSS\s+VERSION\s+\d\.\d\s+RISK\s*\([^)]*\)\s*", "", header_text
  )
  header_text = clean_text(header_text)
  h_low = header_text.lower()

  # Normalize and deduplicate to exact requested names
  if any(t in h_low for t in ["attack complex", "attackcomplex"]):
    return "Attack_Complex"
  if any(t in h_low for t in ["attack vector", "attackvector", "av"]):
    return "Attack_Vector"
  if any(t in h_low for t in ["availability", "avail-ability"]):
    return "Availability"
  if any(t in h_low for t in ["base score", "basescore"]):
    return "Base_Score"
  if any(t in h_low for t in ["confidentiality", "confid-entiality"]):
    return "Confidentiality"
  if any(t in h_low for t in ["integrity", "inte-grity"]):
    return "Integrity"
  if any(t in h_low for t in ["privs", "privreq", "privsreq", "privs req'd"]):
    return "PrivsReqd"
  if any(t in h_low for t in ["remote exploit", "remoteexploit", "remote exploit without auth"]):
    return "RemotExploitNoAuth"
  if any(t in h_low for t in ["user interact", "userinteract"]):
    return "User_Interact"
  if any(
      t in h_low
      for t in ["component", "package and/or", "package/componen", "package/privs"]
  ) and "supported versions" not in h_low:
    return "Package_Privileges"
  if "scope" in h_low:
    return "Scope"
  if "cve" in h_low:
    return "CVE_ID"
  if "supported versions" in h_low:
    return "Supported_Versions_Affected"
  if "protocol" in h_low:
    return "Protocol"
  if "notes" in h_low:
    return "Notes"

  # For any remaining uncaught headers, replace spaces with underscores
  return header_text.replace(" ", "_")


def extract_base_score(cell_text):
  """Extracts float score using regex to handle footnotes or extra characters."""
  match = re.search(r"(\d+\.\d+)", cell_text)
  if match:
    try:
      val = float(match.group(1))
      if 0.0 <= val <= 10.0:
        return val
    except ValueError:
      pass
  match_int = re.search(r"\b(10|\d)\b", cell_text)
  if match_int:
    try:
      val = float(match_int.group(1))
      if 0.0 <= val <= 10.0:
        return val
    except ValueError:
      pass
  return 0.0


def parse_cvss_metric(metric_type, val_str):
  """Maps cell string values to standard CVSS v3.1 single-letter codes."""
  clean_v = re.sub(r"[\*\d\[\]]", "", val_str).strip()
  clean_v_low = clean_v.lower()

  if metric_type == "av":
    if "net" in clean_v_low or clean_v_low == "n":
      return "N"
    if "adj" in clean_v_low or clean_v_low == "a":
      return "A"
    if "loc" in clean_v_low or clean_v_low == "l":
      return "L"
    if "phy" in clean_v_low or clean_v_low == "p":
      return "P"
  elif metric_type == "ac":
    if "low" in clean_v_low or clean_v_low == "l":
      return "L"
    if "high" in clean_v_low or clean_v_low == "h":
      return "H"
  elif metric_type == "pr":
    if "none" in clean_v_low or clean_v_low == "n":
      return "N"
    if "low" in clean_v_low or clean_v_low == "l":
      return "L"
    if "high" in clean_v_low or clean_v_low == "h":
      return "H"
  elif metric_type == "ui":
    if "none" in clean_v_low or clean_v_low == "n":
      return "N"
    if "req" in clean_v_low or clean_v_low == "r":
      return "R"
  elif metric_type == "s":
    val_unspaced = clean_v.replace("-", "").replace(" ", "")
    if "unchanged" in val_unspaced.lower() or val_unspaced.lower() == "u":
      return "U"
    if "changed" in val_unspaced.lower() or val_unspaced.lower() == "c":
      return "C"
  elif metric_type in ["c", "i", "a"]:
    if "none" in clean_v_low or clean_v_low == "n":
      return "N"
    if "low" in clean_v_low or clean_v_low == "l":
      return "L"
    if "high" in clean_v_low or clean_v_low == "h":
      return "H"
  return None


def parse_date_string(date_str, rss_published_str=""):
  """Robustly parses Oracle's various date strings, with a fallback to the RSS date."""
  if date_str:
      # Common Oracle date formats
      formats = [
          "%Y-%b-%d", "%d-%b-%Y", "%Y-%B-%d", "%d-%B-%Y",
          "%B %d, %Y", "%b %d, %Y", "%Y-%m-%d", "%m/%d/%Y"
      ]
      clean_date = date_str.strip()
      for fmt in formats:
          try:
              return datetime.strptime(clean_date, fmt)
          except ValueError:
              continue

  # Fallback to RSS date parsing if custom string parsing fails
  if rss_published_str:
      try:
          return parsedate_to_datetime(rss_published_str).replace(tzinfo=None)
      except Exception:
          pass

  # Final failsafe so the script doesn't crash, but defaults to 'now' (which will be kept)
  return datetime.now()


def scrape_initial_release_date(soup):
  """Extracts publication date from the Modification History table ('Rev 1. Initial Release')."""
  for h in soup.find_all(["h2", "h3", "h4"]):
    if "modification history" in h.get_text().lower():
      table = h.find_next("table")
      if table:
        for tr in table.find_all("tr"):
          cells = [c.get_text(strip=True) for c in tr.find_all(["td", "th"])]
          text_row = " ".join(cells).lower()
          if "initial release" in text_row or "rev 1" in text_row:
             # The date is usually the first column, check cells for date-like string or fallback to first cell
             for c in cells:
                 if re.match(r"\d{4}-\w{3,9}-\d{1,2}|\d{2}-\w{3,9}-\d{4}|\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\w+\s\d{1,2},\s\d{4}", c):
                     return clean_text(c)
             if cells:
                 return clean_text(cells[0])
  return ""


def extract_product_from_title(matrix_title):
  """Derives product name from the Risk Matrix title."""
  title = re.sub(
      r"(?i)risk\s+matrix|content|definitions", "", matrix_title
  ).strip()
  title = title.strip(" -.:")
  return title if title else "Oracle Product"


def parse_table_structure(table):
  """Parses table headers using a 2D grid matrix and isolates valid data rows."""
  rows = table.find_all("tr")
  if not rows:
    return [], [], 0

  header_rows = []
  data_rows = []

  for tr in rows:
    text = tr.get_text()
    if re.search(r"CVE-\d{4}-\d+", text, re.I):
      data_rows.append(tr)
    elif not data_rows and (
        tr.find_all("th")
        or "base score" in text.lower()
        or "vector" in text.lower()
        or "cve" in text.lower()
    ):
      header_rows.append(tr)
    elif data_rows:
      data_rows.append(tr)
    else:
      if len(header_rows) >= 1:
        data_rows.append(tr)
      else:
        header_rows.append(tr)

  if not header_rows or not data_rows:
    return [], [], 0

  num_header_rows = len(header_rows)
  max_cols = 0
  for r in header_rows:
    cols_cnt = sum(
        int(cell.get("colspan", 1)) for cell in r.find_all(["th", "td"])
    )
    max_cols = max(max_cols, cols_cnt)

  grid = [["" for _ in range(max_cols)] for _ in range(num_header_rows)]

  for r_idx, r in enumerate(header_rows):
    c_idx = 0
    for cell in r.find_all(["th", "td"]):
      while c_idx < max_cols and grid[r_idx][c_idx] != "":
        c_idx += 1
      if c_idx >= max_cols:
        break

      text = clean_text(cell.get_text())
      rowspan = int(cell.get("rowspan", 1))
      colspan = int(cell.get("colspan", 1))

      for dr in range(rowspan):
        for dc in range(colspan):
          if r_idx + dr < num_header_rows and c_idx + dc < max_cols:
            grid[r_idx + dr][c_idx + dc] = text
      c_idx += colspan

  final_headers = []
  seen_headers = set()
  for c in range(max_cols):
    col_texts = [grid[r][c] for r in range(num_header_rows) if grid[r][c]]
    combined = normalize_header(" ".join(col_texts))
    if not combined:
      combined = f"Col_{c}"

    # Deduplicate column names
    base_name = combined
    counter = 1
    while combined in seen_headers:
      combined = f"{base_name}_{counter}"
      counter += 1

    seen_headers.add(combined)
    final_headers.append(combined)

  return final_headers, data_rows, max_cols


def scrape_advisory_details(url):
  """Scrapes enrichment details from an individual Oracle security alert advisory page."""
  headers = {
      "User-Agent": (
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      )
  }

  advisory_data = {
      "published_date": "",
      "products_patches": [],
      "risk_matrices": [],
  }

  try:
    print(f"  -> Scraping details from: {url}")
    response = requests.get(url, headers=headers, timeout=15)
    if response.status_code != 200:
      print(f"     Failed to fetch page. Status code: {response.status_code}")
      return advisory_data

    soup = BeautifulSoup(response.text, "html.parser")

    # Extract accurate Published Date from Modification History table
    advisory_data["published_date"] = scrape_initial_release_date(soup)

    # 1. Parse Affected Products and Patch Availability Table
    for table in soup.find_all("table"):
      rows = table.find_all("tr")
      if not rows:
        continue

      sample_cols = rows[0].find_all(["td", "th"])
      if len(sample_cols) == 2:
        for tr in rows:
          cols = tr.find_all(["td", "th"])
          if len(cols) == 2:
            col1_text = clean_text(cols[0].get_text())

            a_tag = cols[1].find("a", href=True)
            if a_tag:
              href = a_tag["href"]
              patch_url = (
                  href
                  if href.startswith("http")
                  else f"https://www.oracle.com{href}"
              )
            else:
              patch_url = clean_text(cols[1].get_text())

            if col1_text.lower() in [
                "affected products and versions",
                "product",
            ] or patch_url.lower() in [
                "patch availability document",
                "document",
            ]:
              continue

            if col1_text and patch_url:
              item = {
                  "affected_products_and_versions": col1_text,
                  "patch_availability_url": patch_url,
              }
              if item not in advisory_data["products_patches"]:
                advisory_data["products_patches"].append(item)

    # 2. Parse Risk Matrices
    for h in soup.find_all(["h2", "h3", "h4"]):
      htext = clean_text(h.get_text())
      htext_lower = htext.lower()

      if "risk matrix" in htext_lower and "history" not in htext_lower:
        matrix_title = htext
        default_product = extract_product_from_title(matrix_title)

        table = h.find_next("table")
        if not table:
          continue

        headers_list, data_rows, num_cols = parse_table_structure(table)
        if not headers_list or not data_rows:
          continue

        # Map metric indices
        col_map = {}
        has_product_col = False

        for idx, h_name in enumerate(headers_list):
          h_low = h_name.lower()
          if h_name == "Base_Score":
            col_map["score"] = idx
          elif h_name == "Attack_Vector":
            col_map["av"] = idx
          elif h_name == "Attack_Complex":
            col_map["ac"] = idx
          elif h_name == "PrivsReqd":
            col_map["pr"] = idx
          elif h_name == "User_Interact":
            col_map["ui"] = idx
          elif h_name == "Scope":
            col_map["s"] = idx
          elif h_name == "Confidentiality":
            col_map["c"] = idx
          elif h_name == "Integrity":
            col_map["i"] = idx
          elif h_name == "Availability":
            col_map["a"] = idx
          elif "product" in h_low:
            has_product_col = True

        # Positional fallback if Base Score is identified
        if "score" in col_map:
          score_idx = col_map["score"]
          metrics = ["av", "ac", "pr", "ui", "s", "c", "i", "a"]
          for offset, m in enumerate(metrics, start=1):
            if m not in col_map and (score_idx + offset) < num_cols:
              col_map[m] = score_idx + offset

        matrix_rows = []
        last_component = default_product

        for tr in data_rows:
          cells = [clean_text(c.get_text()) for c in tr.find_all(["td", "th"])]
          if not cells or len(cells) < 3:
            continue

          row_dict = {}
          for idx, cell_val in enumerate(cells):
            if idx < len(headers_list):
              col_name = headers_list[idx]

              # Specifically clean up Scope values (remove "- ")
              if col_name == "Scope":
                  cell_val = cell_val.replace("- ", "")

              row_dict[col_name] = cell_val

          values_str = " ".join(cells).lower()
          if "basescore" in values_str or "attackvector" in values_str:
            continue

          # Ensure Product field is present
          if not has_product_col:
            row_dict["Product"] = default_product

          # Ensure Component field is populated (carry forward or fallback)
          comp_val = row_dict.get("Package_Privileges", "") or row_dict.get("Component", "")
          if not comp_val or comp_val == "-":
            row_dict["Component"] = last_component
          else:
            last_component = comp_val
            row_dict["Component"] = comp_val

          # Base Score Extraction (assigns to "Base_Score", removing Extracted_Base_Score logic)
          base_score = 0.0
          if "score" in col_map and col_map["score"] < len(cells):
            base_score = extract_base_score(cells[col_map["score"]])

          if base_score == 0.0:
            for c_val in cells:
              sc = extract_base_score(c_val)
              if sc > 0.0:
                base_score = sc
                break

          row_dict["Base_Score"] = base_score

          # Complete CVSS Vector Construction
          metric_keys = ["av", "ac", "pr", "ui", "s", "c", "i", "a"]
          extracted_metrics = {}

          for m in metric_keys:
            val_str = ""
            if m in col_map and col_map[m] < len(cells):
              val_str = cells[col_map[m]]
            extracted_metrics[m] = parse_cvss_metric(m, val_str) if val_str else None

          vector_parts = [
              f"{m.upper()}:{extracted_metrics[m]}"
              for m in metric_keys
              if extracted_metrics[m] is not None
          ]
          row_dict["Extracted_CVSS_Vector"] = (
              f"CVSS:3.1/{'/'.join(vector_parts)}" if vector_parts else ""
          )

          if row_dict:
            matrix_rows.append(row_dict)

        if matrix_rows:
          matrix_obj = {
              "matrix_title": matrix_title,
              "vulnerabilities": matrix_rows,
          }
          advisory_data["risk_matrices"].append(matrix_obj)

  except Exception as e:
    print(f"     Error scraping details: {e}")

  return advisory_data


def main():
  print("Fetching Oracle Security Alert RSS feed...")
  feed = feedparser.parse(RSS_URL)

  if not feed.entries:
    print("No entries found in the RSS feed or feed is inaccessible.")
    return

  six_months_ago = datetime.now() - timedelta(days=180)
  print(
      "Filtering advisories published after:"
      f" {six_months_ago.strftime('%Y-%m-%d')}"
  )

  alerts_list = []
  skipped_count = 0

  for entry in feed.entries:
    title = clean_text(entry.get("title", "Unknown Title"))
    link = entry.get("link", "")
    summary = entry.get("summary", "")

    print(f"\nProcessing advisory link: {link}")

    # Scrape enrichment details & obtain official publication date from modification history
    enrichment = {}
    if link:
      enrichment = scrape_advisory_details(link)

    pub_date_str = enrichment.get("published_date", "")
    rss_date_str = entry.get("published", entry.get("updated", ""))

    # Evaluate the date aggressively
    parsed_date = parse_date_string(pub_date_str, rss_date_str)

    if parsed_date < six_months_ago:
      skipped_count += 1
      print(f"  -> Skipped (Older than 6 months: {parsed_date.strftime('%Y-%m-%d')})")
      continue

    print(f"  -> Keeping advisory: {title} ({parsed_date.strftime('%Y-%m-%d')})")

    # Override the JSON's published_date with the cleanly formatted string we parsed
    clean_pub_date_str = parsed_date.strftime('%Y-%m-%d')

    alert_item = {
        "title": title,
        "link": link,
        "published_date": clean_pub_date_str,
        "summary": summary,
        "products_patches": enrichment.get("products_patches", []),
        "risk_matrices": enrichment.get("risk_matrices", []),
    }

    alerts_list.append(alert_item)

  print(f"\nSkipped {skipped_count} advisories older than 6 months.")

  # 1. Save standard nested JSON
  with open("oracle_security_alerts.json", "w", encoding="utf-8") as f:
    json.dump({"alerts": alerts_list}, f, indent=2, ensure_ascii=False)

  # 2. Save Flat JSON Output
  if GENERATE_FLAT_OUTPUT:
    flat_vulnerabilities = []
    for alert in alerts_list:
      for matrix in alert.get("risk_matrices", []):
        matrix_title = matrix.get("matrix_title", "")
        for vuln in matrix.get("vulnerabilities", []):
          flat_entry = {
              "advisory_title": alert["title"],
              "published_date": alert["published_date"],
              "link": alert["link"],
              "matrix_title": matrix_title,
          }
          flat_entry.update(vuln)
          flat_vulnerabilities.append(flat_entry)

    with open("oracle_vulnerabilities_flat.json", "w", encoding="utf-8") as f:
      json.dump(
          {"vulnerabilities": flat_vulnerabilities},
          f,
          indent=2,
          ensure_ascii=False,
      )
    print(
        f"Saved flat output with {len(flat_vulnerabilities)} vulnerabilities"
        " to oracle_vulnerabilities_flat.json"
    )


if __name__ == "__main__":
  main()
