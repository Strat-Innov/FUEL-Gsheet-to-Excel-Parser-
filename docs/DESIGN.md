# FUEL Response Parser — Design Specification (v0.6, for approval)

| Item | Value |
|---|---|
| Phase | DESIGN (no production code) |
| Source analysed | `FUEL RESP (CSP-AD MARF) v3.3.xlsx` (Google Forms response export) |
| Destination analysed | `FUEL (MGM-AD ADMIN) (2).xlsx` → sheet `MASTERLIST` (Microsoft Forms export) |
| Status | **v0.6: exports-only access model (Section M). Blocked on B1 (seller-slot evidence) and B2/B3 policy approvals. See M.5** |

> This document contains structure, counts and rules only. It holds no client records, phone numbers or e-mail addresses. Both workbooks were opened read-only, and neither was modified.

---

## 0. Executive Summary

The two workbooks come from **different form platforms**: the source is a Google Forms export and the destination is a Microsoft Forms export. That difference explains most of the structural gaps.

The five findings that most affect the design:

1. **The 13 "client blocks" in `MARF RAW` are team branches, not extra clients.** Every one of the 3,064 responses fills **exactly one** block. Which block it fills depends on the sales leader and the form version. The parser must therefore *select* the single filled block, not *explode* rows. It must still handle more than one filled block safely, because a future form could have a real "add another client" section.
2. **Block headers cannot be identified by their numeric suffix.** Google Forms numbers duplicate titles per distinct text. Block 6's middle-name column is called `Client's Middle Name 2`, while its other fields end in `6`. Blocks must be detected by **header grouping**: a block starts at each `Seller Name` and holds 10 known base fields.
3. **The source contains two form versions.** v1 (≈ 398 responses, up to ~15 May 2026) used different lead-source values and detail columns (`Ownsource`, `Booth`, `Digital`, `Referral`) and had no *Source of Awareness*. v2 matches the destination's vocabulary. v1 values need an approved translation table.
4. **The source has one seller per response and no seller-email question.** The only e-mail is the submitting Google account. That account differs from the seller's usual account in **213 responses (≈ 7%)**, and 3 accounts submitted for more than one seller. Using it as "Seller's Email Address" without a safeguard would mis-pair some sellers and e-mails.
5. **The destination has semantic gaps.** Two columns are unnamed MS Forms placeholders (`Question`, `Question1`). There is no column for any "If Others, please specify" text. The destination's own sample rows already store mobile numbers as numbers, which loses the leading `0`.

Recommended solution: a **local Python tool**. It has a pure processing core (`openpyxl`), a **config-driven header-alias mapping** (no column letters), a **SQLite import ledger** for re-import protection, and a thin **Streamlit** UI that runs on localhost. It produces a **new** import-ready workbook plus a separate audit workbook.

---

## A. Source Workbook Structure Analysis

### A.1 Sheets

| Sheet | State | Role | Use in parser |
|---|---|---|---|
| `MARF RAW` | visible | Raw Google Forms responses: 3,064 data rows, 168 columns (A:FL), header in row 1 | **Authoritative input** |
| `TEAM BERNADETTE`, `TEAM LESLIE`, `TEAM FERDINAND`, `TEAM TINE`, `TEAM RALPH`, `AM JESSA` | visible | Derived "Leads Log" views. Title rows 1–2, header row 4. Built by Google-Sheets-only formulas (`__xludf.DUMMYFUNCTION`, cached values). All text is UPPER-CASED. Lead-source detail is concatenated into one "SPECIFIC LEAD SOURCE" column | **Ignored** (derived and lossy). Used only as evidence of business intent |
| `AM JANESA` | hidden | Empty team log | Ignored |

Why the team sheets are not a source: they are lossy (upper-cased, details concatenated with `", "`). Some "Others" columns are omitted (`If others, please specify 3`, `If Others, please specify below 2`). Their formulas use fixed row ranges (e.g. `A2:C971`) that do not match the exported row positions. They confirm one mapping intent, though: the team sheets label `MARF RAW` column *Email Address* as **"SELLER EMAIL ADDRESS"**.

### A.2 `MARF RAW` column groups

Column letters below are for orientation only. The parser never uses them.

| Group | Headers (as found; trailing spaces exist) | Level | Notes |
|---|---|---|---|
| Response metadata | `Timestamp`, `Email Address`, `SALES LEADER` | Response | Timestamp is a real datetime with milliseconds, and all 3,064 values are unique. Email Address is the submitting Google account (50 distinct values). SALES LEADER has 6 values (see A.4) |
| Client block × 13 | `Seller Name `, `Client's Surname`, `Client's First Name`, `Client's Middle Name`, `Mobile No. (Client)`, `Email Address (Client)`, `Residence Address (Client)`, `Project of Interest (First Choice)`, `Project of Interest (Second Choice)`, `REMARKS ` (+ Google suffixes ` 2`…` 13`) | Block | Blocks sit at D, N, X, AH, AR, BP, BZ, CJ, CT, DD, EF, ES, FC. Blocks 5, 9 and 13 have never been used |
| Lead source | `Lead Source ` | Response | 16 distinct values covering both versions |
| v1 lead-source details | `Own Source `, `Referral`, `Digital`, `Seller Facebook Page `, `If Saturation, please specify location`, `If Booth, please specify Booth Location`, `If Sponsorship Event, please specify name of event & date (…)`, `Projects owned/bought by Client `, 5 × `If others, please specify …` | Response | Located between the client blocks (BC–BO) |
| v2 lead-source details | `If Booth, please choose which Booth Location below:` … `If Sponsorship Event, please specify below the name and date of event (…)`, `If Corporate…`, `If Showroom Walk-In…`, `If Roadshow…`, 8 × `If Others, please specify below N` | Response | DN–ED and EP–ER |
| Awareness | `Source of Awareness` | Response | Multi-select joined with `", "`. **Option text contains commas** (`Social Media (Facebook, Instagram, etc)`). Empty for all v1 responses |

### A.3 Client-block usage (observed)

| Block | Sales leader | Responses | Date span | Interpretation |
|---|---|---|---|---|
| 1 (D) | TEAM BERNADETTE | 518 | May–Sep 2026 | One block across both versions |
| 2 (N) / 6 (BP) | TEAM LESLIE | 169 / 854 | v1 / v2 | Team section re-created in v2 |
| 3 (X) / 7 (BZ) | TEAM FERDINAND | 114 / 428 | v1 / v2 | same |
| 4 (AH) / 8 (CJ) | TEAM RALPH | 29 / 130 | v1 / v2 | same |
| 10 (DD) / 12 (ES) | AM: JESSA | 441 / 1 | v2 / test-like | |
| 11 (EF) | TEAM CHRISTINE | 380 | Aug–Oct 2026 | New team added later |
| **Total** | | **3,064** | | **= response count; 0 responses with 0 or ≥ 2 blocks** |

### A.4 Value profiles relevant to mapping

| Field | Observation | Design consequence |
|---|---|---|
| SALES LEADER | `TEAM LESLIE (AAD: …)`, `TEAM FERDINAND (AAD: …)`, `TEAM BERNADETTE (AAD: …)`, `TEAM CHRISTINE (AAD: …)`, `TEAM RALPH (AAD: …)`, `AM: JESSA`. Same "TEAM X (AAD: Name)" pattern as the destination | Preserve verbatim. Map team via config |
| Sheet name vs leader | Sheet `TEAM TINE` holds `TEAM CHRISTINE` responses | Never derive team from sheet names |
| Seller Name | 30 distinct. Prefixes `AS: `, `AM: `, `AAD ` (no colon) | Preserve verbatim (the destination keeps `AS: ` too) |
| Submitter e-mail vs seller | 21 sellers appear under more than one submitting account. 213 rows differ from the seller's most-used account. 3 accounts submitted for several sellers | Submitter ≠ seller identity (see D.2) |
| Mobile No. | Stored as **text** in the source. 2,967 are `09XXXXXXXXX` (11 digits), 9 are `+639XXXXXXXXX`, 88 are blank | Keep as text and never cast to a number |
| Client e-mail | 216 values are not valid e-mail syntax. Most are null tokens (`N/A`, `NA`, `Na`, `N/a`, `None`: 309 in total across block fields), but some are other text (e.g. a place name) | Validation and null-token rule needed |
| Projects | Multi-select joined with `", "` in **both** first and second choice (e.g. `Studio N, The Levels`). One typo variant `Gold Ridge` | Split against a known project list. Typos → review, never auto-fix |
| Residence | v2 uses `REGION - City` (same as the destination). v1 has free text | Preserve verbatim |
| Lead Source | v2: `Own Source`, `Saturation`, `Sponsorship Event`, `Digital - Seller Social Media`, `Booth Manning`, `Digital - Official Social Media`, `Buyer Referral`, `Others`, `Marketing Partner`, `Repeat Buyer`, `Digital - Official Website`. v1 only: `Ownsource`, `Booth`, `Digital`, `Referral`, `Site Walk-in` | v1 → v2 translation table (decision H) |
| Others-specify pairing | Observed co-occurrence: `…below 3`↔Own Source, `…below 2`↔Seller Social Media, `…below 4`↔Repeat Buyer, `…below 5`↔**Lead Source = Others** (not Sponsorship, despite sitting next to it) | "Others" columns cannot be resolved by "nearest preceding column". They need explicit config **plus** a co-occurrence check |
| Duplicates | 0 duplicate timestamps. **2 pairs** have identical content but different timestamps (2–3 weeks apart). **38 groups / 77 rows** share client surname, first name and mobile number | Tiered duplicate rule (D.4) |

---

## B. Destination `MASTERLIST` Schema Analysis

Microsoft Forms export: an Excel Table `OfficeForms.Table` (A1:AI11) with 35 columns, 10 sample/test rows, no data validation, no conditional formatting and no freeze panes. **Header text must be reproduced byte-for-byte**: several headers contain non-breaking spaces (U+00A0) and trailing newlines.

| # | Col | Header (cleaned for reading) | Type / format in sample | Meaning |
|---|---|---|---|---|
| 1 | A | Id | integer | MS Forms response sequence |
| 2 | B | Start time | datetime `m/d/yy h:mm` | Form opened |
| 3 | C | Completion time | datetime `m/d/yy h:mm` | Form submitted |
| 4 | D | Email | text (`anonymous`) | Responder identity |
| 5 | E | Name | blank | Responder name |
| 6 | F | SALES LEADER | text `@` | Team label |
| 7–8 | G–H | Seller Name / Seller's Email Address | text | **Seller slot 1** |
| 9–10 | I–J | Seller Name1 / Seller's Email Address1 | text | **Seller slot 2** (sample: used by `TEAM HD`) |
| 11–12 | K–L | Seller Name2 / Seller's Email Address2 | text | **Seller slot 3** |
| 13 | M | Client's Surname | text | |
| 14 | N | Client's First Name | text | |
| 15 | O | Client's Middle Name␠(nbsp) | text | |
| 16 | P | Client's Mobile No.⏎ | **number in sample (leading 0 lost: defect)** | |
| 17 | Q | Client's Email Address | text | |
| 18 | R | Client's Residence Address | text `REGION - City` | |
| 19 | S | Project of Interest (First Choice) | text, single value in samples | |
| 20 | T | Project of Interest (Second Choice) | text, multi joined with **`;`** | |
| 21 | U | REMARKS | text | |
| 22 | V | Lead Source | text (v2 vocabulary + `Showroom Walk-In`) | |
| 23 | W | If Booth, please choose which Booth Location below: | text | |
| 24 | X | Name of existing buyer who referred the client to you | text | |
| 25 | Y | **Question** | never populated | Unnamed MS Forms question (see C) |
| 26 | Z | If Digital - Official Social Media, please choose below: | text | |
| 27 | AA | If Digital - Seller Social Media, please choose below: | text | |
| 28 | AB | If Own Source, please choose below: | text | |
| 29 | AC | If Repeat Buyer, please choose from the listed projects below: | text | |
| 30 | AD | If Roadshow, please specify below the name and date of roadshow event (…) | text | |
| 31 | AE | If Saturation, please specify exact location below | text | |
| 32 | AF | **Question1** | text (sample values are showroom names when Lead Source = `Showroom Walk-In`) | Unnamed MS Forms question (see C) |
| 33 | AG | If Sponsorship Event, please specify below the name and date of event (…) | text | |
| 34 | AH | Source of Awareness — CLIENT'S PERSPECTIVE … (you may choose more than one) | text, multi joined with **`;`** | |
| 35 | AI | If Digital - Seller Facebook Page, please choose below: | text | Positioned last (added to the form later) |

Destination conventions the parser must follow: multi-select separator `;`, text cells formatted `@`, and datetimes formatted `m/d/yy h:mm`. The destination has **no** columns for v1-only details, "Others, please specify" text, the Corporate detail (unless `Question` is it), or any source identifier.

The sample rows are test data (`Option 1`, `anonymous`, `TEAM CALOY`, `TEAM HD`). None of the 6 source team labels appears in the destination yet.

---

## C. Source → Destination Field Mapping

**Header matching key** = NFKC-normalise → replace U+00A0 with a space → lower-case → collapse whitespace → strip → remove a trailing Google duplicate suffix (` \d+$`) → remove a trailing `:`. Each destination field has an **alias list** in config. Block fields are resolved *inside the detected block*. Duplicated response-level headers (the "Others" columns) are resolved by *occurrence order among identical keys* and must pass the stated co-occurrence check.

Confidence: **H** = same meaning, verified by data. **M** = same meaning, but the value form or vocabulary differs, or the evidence is limited. **L** = inferred and needs approval. **—** = no source.

| # | Destination field | Source header (match key / alias) | Source location | Transformation rule | Conf. | Validation |
|---|---|---|---|---|---|---|
| 1 | Id | — | — | Generated sequence (decision H-5) | — | Unique within output |
| 2 | Start time | — | — | Copy of Timestamp, or blank (decision H-4) | L | — |
| 3 | Completion time | `timestamp` | Response | Keep the datetime value unchanged (ms retained in the value). Cell format `m/d/yy h:mm`. Text timestamps are parsed only with configured formats; ambiguous day/month → review | H | Required. Must be a datetime or parse unambiguously. Not in the future |
| 4 | Email | `email address` (response level, **not** `(client)`) | Response | Verbatim, trimmed, lower-cased for comparison only (decision H-3) | M | Valid e-mail syntax, else warning |
| 5 | Name | — | — | Blank | H | — |
| 6 | SALES LEADER | `sales leader` | Response | Verbatim, trimmed | H | Required. Must exist in `teams` config, else review |
| 7–12 | Seller Name / Seller's Email Address (slot 1/2/3) | `seller name` | Filled block | Name verbatim → slot chosen by `teams[leader].seller_slot`. Email per decision H-1. The name and its e-mail are always written as a **pair into the same slot** | M | Seller name required. Seller e-mail valid syntax. Mismatch against the seller directory → warning |
| 13 | Client's Surname | `client's surname` | Filled block | Trim. Keep case | H | Required |
| 14 | Client's First Name | `client's first name` | Filled block | Trim. Keep case | H | Required |
| 15 | Client's Middle Name | `client's middle name` | Filled block | Trim. Null tokens per H-11 | H | — |
| 16 | Client's Mobile No. | `mobile no. (client)` | Filled block | **Text.** Trim only. If the reader receives a number, write its digits as text and flag `MOBILE_WAS_NUMERIC`. **Never add a leading 0, never infer or invent digits.** Cell format `@` | H | Pattern `^09\d{9}$` or `^\+639\d{9}$` → OK. Anything else → `MOBILE_FORMAT_UNEXPECTED` (WARNING, classified by reason: see D.8). Blank → `MOBILE_MISSING` (WARNING) |
| 17 | Client's Email Address | `email address (client)` | Filled block | Trim. Null tokens per H-11 | H | Valid syntax, else warning (value kept) |
| 18 | Client's Residence Address | `residence address (client)` | Filled block | Verbatim | H | — |
| 19 | Project (First Choice) | `project of interest (first choice)` | Filled block | Split on `", "` **against the known project list**, re-join with `;` | M | Each token must be in the project list, else warning `UNKNOWN_PROJECT`. More than one token → warning `MULTI_FIRST_CHOICE` (H-9) |
| 20 | Project (Second Choice) | `project of interest (second choice)` | Filled block | Same split and re-join | H | Same list check |
| 21 | REMARKS | `remarks` | Filled block | Verbatim, plus appended "Others" texts if H-8 option (a) is approved | H | — |
| 22 | Lead Source | `lead source` | Response | v2 values verbatim. v1 values through the approved translation table (H-7). Untranslatable → review | H (v2) / M–L (v1) | Required. Must be in the destination lead-source list |
| 23 | Booth location | v2 `if booth, please choose which booth location below`; v1 `if booth, please specify booth location` | Response | Verbatim (vocabularies differ: H-10) | H / M | Expected only when Lead Source ∈ {Booth Manning} |
| 24 | Referrer name | `name of existing buyer who referred the client to you` | Response | Verbatim | H | Expected only when Lead Source = Buyer Referral. **v1 `Referral` (Friend/Relative) is a category, not a name: it is never mapped here** |
| 25 | Question | `if corporate, please choose below` | Response | Verbatim | **L** | Inferred from question order only. No data in either file (H-6) |
| 26 | Digital – Official SM | `if digital - official social media, please choose below`; v1 `digital` where value ∈ {`Project's Main Page`} | Response | v2 verbatim. v1 via value map | H / M | Lead Source = Digital - Official Social Media |
| 27 | Digital – Seller SM | `if digital - seller social media, please choose below`; v1 `digital` where value ∈ {`Seller Facebook Page`, `Facebook Marketplace`, `Tiktok Inquiry`} | Response | v2 verbatim. v1 via value map | H / M | Lead Source = Digital - Seller Social Media |
| 28 | Own Source | `if own source, please choose below`; v1 `own source` | Response | Verbatim (same option set) | H | Lead Source = Own Source |
| 29 | Repeat Buyer project | `if repeat buyer, please choose from the listed projects below`; v1 `projects owned/bought by client` | Response | Split and re-join with `;` | H / M | Lead Source = Repeat Buyer |
| 30 | Roadshow | `if roadshow, please specify below the name and date of roadshow event…` | Response | Verbatim | H | Lead Source = Roadshow |
| 31 | Saturation location | v2 `if saturation, please specify exact location below`; v1 `if saturation, please specify location` | Response | Verbatim | H | Lead Source = Saturation |
| 32 | Question1 | `if showroom walk-in, please specify which showroom below` | Response | Verbatim | **M** | Supported by destination sample data (showroom names under `Showroom Walk-In`). Source column is empty (H-6) |
| 33 | Sponsorship event | v2 `if sponsorship event, please specify below the name and date of event…`; v1 `if sponsorship event, please specify name of event & date…` | Response | Verbatim | H | Lead Source = Sponsorship Event |
| 34 | Source of Awareness | `source of awareness` | Response | **Tokenise by the known option list** (longest match first, so `Social Media (Facebook, Instagram, etc)` stays whole), then re-join with `;`. Unknown leftover text → kept as a token, with a warning | M | Every token known, else warning. Blank for v1 rows is **not** an error |
| 35 | Digital – Seller FB Page | `if digital - seller facebook page, please choose below`; v1 `seller facebook page` | Response | v2 verbatim. v1 `…Own/Seller Boosting Budget` → `…Own/Seller's Boosting Budget` (value map) | H / M | Seller SM = Seller Facebook Page |

**Source columns with no destination column** (each is listed in the audit, never silently dropped):

| Source | Pairs with (verified co-occurrence) | Proposed handling |
|---|---|---|
| v1 `If others, please specify` (after Own Source) | Ownsource = Others | H-8 |
| v1 `If others, please specify 2` (after Referral) | Referral = Others | H-8 |
| v1 `If others, please specify 3` (after Digital) | Digital = Others | H-8 |
| v1 `If Others, please specify project` | Repeat Buyer = Others | H-8 |
| v1 `If others, please specify` (last, BO) | Lead Source = Others | H-8 |
| v1 `Referral` (Friend / Relative / Others) | Lead Source = Referral | H-7 |
| v2 `If Others, please specify below` (after Corporate) | Corporate = Others | H-8 |
| v2 `…below 2` | Seller Social Media = Others | H-8 |
| v2 `…below 3` | Own Source = Others | H-8 |
| v2 `…below 4` | Repeat Buyer = Others | H-8 |
| v2 `…below 5` | **Lead Source = Others** | H-8 |
| v2 `…below 6/7/8` | never filled | Report if ever non-empty |

Any non-empty source column that matches no alias is reported as `UNMAPPED_COLUMN` in the mapping preview. The run cannot continue until the user marks it *ignore* or maps it.

---

## D. Strategies for Critical Record Handling

### D.1 Client blocks

1. **Detection (per workbook, from headers only).** Walk the headers left to right. Each header whose match key is `seller name` opens a new block. The following headers join the block while their key is one of the 10 block base keys. The block closes at the first non-block key or at the next `seller name`. Validate that each block has all 10 base keys exactly once. Otherwise raise `BLOCK_STRUCTURE_ERROR` and stop with a clear message; never guess.
2. **Selection (per response).** A block is *filled* if any of its 10 cells is non-blank after the null-token rule.
   - **Exactly 1 filled** → one destination record (the normal case: 3,064 of 3,064).
   - **0 filled** → no record. Status `NO_CLIENT_DATA`, sent to review.
   - **≥ 2 filled** → one record **per filled block**, each with the shared response-level fields. All of them are held in review with `MULTI_BLOCK_RESPONSE` (H-15). Nothing is merged or discarded.
3. **Team consistency check.** Config may list the expected blocks per team (learned from the A.3 table). A response whose filled block does not match its SALES LEADER → warning `BLOCK_TEAM_MISMATCH`.

### D.2 Seller-to-e-mail pairing

- The source has **one** seller per response, and the seller name comes from the filled block. The pair (name, e-mail) is assembled **once** as a unit and written into **one** slot (`G/H`, `I/J` or `K/L`). E-mails are never shifted across slots.
- **Slot choice** comes from `teams.yaml` (`seller_slot: 1|2|3`). Evidence suggests the destination form also branches per team (sample `TEAM HD` uses slot 2). An unknown team → review, never a default slot (H-2).
- **E-mail source** (H-1, recommended option b):
  1. If an optional `sellers.yaml` directory contains the seller name, use the directory e-mail. If the submitting e-mail differs → warning `SUBMITTER_NOT_SELLER` (expected ≈ 213 rows).
  2. Otherwise, use the submitting e-mail (the team-sheet convention) with info flag `SELLER_EMAIL_FROM_SUBMITTER`.
- Missing name with an e-mail present → error `SELLER_NAME_MISSING` (review). A name with no e-mail available → warning, and the e-mail cell is left blank. A blank never pulls in another slot's value.

### D.3 Team identification and output scope

- SALES LEADER is preserved verbatim. `teams.yaml` maps each label (exact, after normalisation) to `team_key`, `seller_slot`, `expected_blocks` and optional aliases.
- Output scope: one output containing all teams by default, with a UI filter to export selected teams only (H-16). No team is hard-coded in the code.

### D.4 Duplicate protection (documented rule; nothing is deleted)

| Tier | Rule | Action | Expected on current file |
|---|---|---|---|
| D0 Re-import | `src_key` already in the import ledger | Skipped, counted as `ALREADY_IMPORTED` | 0 on first run, all on a repeat run |
| D1 Same file | SHA-256 of the file bytes already in the ledger | Banner warning before processing (not blocking) | — |
| D2 Exact content | Same `content_hash` = all mapped fields except timestamps/Id | First occurrence imported. Later ones held, `PROBABLE_DUPLICATE` (H-14) | 2 pairs |
| D3 Same client | Same normalised mobile (last 10 digits) + case-folded surname + first name | Imported with warning `POSSIBLE_DUPLICATE_CLIENT` and a group id | 38 groups / 77 rows |

Names alone or timestamps alone never trigger an action.

### D.5 Source record identifier and import history

- `src_key = SHA-256( "FUEL-MARF" | timestamp ISO-8601 to microseconds | submitter e-mail (lower) | seller name | client surname (case-folded) | mobile digits )`.
  - Timestamp plus submitter is unique for every current response. The client fields keep multiple blocks in one response distinct. **The file name is not part of the key**, so a later export (e.g. v3.4) containing the same responses is still recognised.
  - Known limit: editing a response in Google Forms changes its timestamp. The edited response then gets a new key, and D2/D3 catch it as a duplicate instead.
- **Ledger:** a local SQLite file, `import_history.sqlite` (stdlib `sqlite3`, no server).

```
import_batch(batch_id TEXT PK, created_at, source_file_name, source_file_sha256,
             source_sheet, mapping_profile_id, responses_read, records_built,
             records_exported, records_review, records_skipped, status  -- exported|voided
             , output_path)
imported_record(src_key TEXT PK, batch_id FK, source_row INT, dest_id INT,
                content_hash TEXT, exported_at)
```

The ledger stores hashes and row numbers only. **No client PII is stored.** Keys are written only when the export file is saved successfully, in one transaction. A **"Void batch"** action removes a batch's keys, for when its output was never appended to the live MASTERLIST.

### D.6 Data preservation

- Workbooks are opened with `read_only=True`. Output always goes to a **new** file whose name includes a timestamp. The tool refuses an output path equal to either input.
- All text cells are written as strings with format `@`. Mobile numbers are never numeric.
- **Formula-injection guard:** openpyxl turns any string starting with `=` into a formula. Such values are written as explicit string cells, and the event is logged.
- Header row, column widths and the cell format pattern are copied from the destination template. Headers are copied **byte-exact** (nbsp and newlines included).
- Original timestamps are kept as datetime values; the format only hides seconds.

### D.7 Missing or ambiguous data

Every issue is an `Issue(code, severity, source_row, field, message)`.

- **ERROR**: the record goes to review only.
- **WARNING**: the record is exported **and** listed in the Review report (H-13). It is not held back.
- **INFO**: written to the audit log only.

The parser never fills a value it cannot trace to a source cell or an approved rule.

### D.8 Validation policy (approved v0.3)

Severity decides what happens to a record. Every code below is non-destructive: the flagged value is written exactly as received (after the approved whitespace trim) and never corrected. Severities live in `config/validation_policy.yaml`, so a later change in policy is a config change, not a code change. Each run records the policy file's hash in the audit Summary.

| Code | Severity | Raised when | Value written |
|---|---|---|---|
| `MOBILE_FORMAT_UNEXPECTED` | WARNING | Mobile is not `09XXXXXXXXX` or `+639XXXXXXXXX`. The issue carries one or more **reason classes**: `missing_leading_zero` (10 digits starting with 9), `wrong_length`, `non_digit_characters`, `repeated_digits` (one digit repeated, e.g. `9999999999`) | Original text, unchanged. No digit added or removed |
| `MOBILE_WAS_NUMERIC` | WARNING | The cell was stored as a number, so a leading 0 may already have been lost upstream | Digits as text, nothing prefixed |
| `MOBILE_MISSING` | WARNING | Mobile is blank | Blank |
| `INVALID_EMAIL` | WARNING | Seller or client e-mail fails syntax: exactly one `@`, no spaces or commas, a domain containing a dot, valid characters | Original text, unchanged |
| `WHITESPACE_TRIMMED` | INFO | Leading or trailing whitespace removed from any text field. The audit stores the field and the before/after lengths, not the value | Trimmed text |
| `SELLER_SLOT_MISMATCH` | ERROR | The slot configured for the SALES LEADER is empty but another slot holds a seller name or e-mail | Nothing is moved. Record → Review only |
| `UNKNOWN_TEAM` | ERROR | SALES LEADER is not in the verified team configuration | Record → Review only |
| `SELLER_NAME_MISSING`, `MISSING_TIMESTAMP`, `INVALID_DATE`, `AMBIGUOUS_DATE`, `MULTI_BLOCK_RESPONSE`, `NO_CLIENT_DATA`, `UNMAPPED_VALUE` | ERROR | As defined in C and D | Record → Review only |
| `PROBABLE_DUPLICATE` | HELD | D.4 tier D2 | Later occurrence → Review only |
| `POSSIBLE_DUPLICATE_CLIENT`, `ORPHAN_DETAIL`, `UNKNOWN_OPTION`, `UNKNOWN_PROJECT`, `MULTI_FIRST_CHOICE`, `SUBMITTER_NOT_SELLER` | WARNING | As defined in C and D | Original value |
| `SECOND_INCLUDES_FIRST`, `SELLER_EMAIL_FROM_SUBMITTER` | INFO | As defined | Original value |

E-mail validation is **syntax only** (H-24). The parser does not reject unfamiliar domains and does not suggest or apply corrections. A plausible misspelling with valid syntax, such as `name@gmaol.com`, is therefore **not flagged** (accepted residual risk R16).

---

## E. Data-Quality Risks and Safeguards

| # | Risk | Evidence | Severity | Safeguard |
|---|---|---|---|---|
| R1 | Block mis-identification from suffix numbers | `Client's Middle Name 2` belongs to block 6 | High | Header-group detection plus a 10-key completeness check (D.1) |
| R2 | Seller e-mail mis-attributed | 213 rows where submitter ≠ seller's usual account | High | Seller directory, mismatch warning, pair written atomically (D.2) |
| R3 | Leading-zero loss in mobile numbers | Destination samples are already numeric | High | Text cells `@`, `MOBILE_WAS_NUMERIC` flag |
| R4 | Comma inside multi-select options | `Social Media (Facebook, Instagram, etc)` | Medium | Tokenise by the known option list, longest match first |
| R5 | v1/v2 vocabulary mix | 398 v1 responses | Medium | Explicit translation table. Untranslatable → review |
| R6 | "Others" text lost | 12 others-specify columns with no destination column | Medium | H-8, plus a co-occurrence check |
| R7 | Unnamed destination columns (`Question`, `Question1`) | No source data for either | Medium | Low/Medium confidence until approved (H-6) |
| R8 | Re-import duplicates | Same responses re-exported in later versions | High | `src_key` ledger (D.5) |
| R9 | Genuine duplicates / re-registrations | 2 exact pairs, 38 same-client groups | Medium | Tiered rule, no deletion (D.4) |
| R10 | Null tokens (`N/A`, `Na`, `None`) treated as data | 309 occurrences | Low | Configurable null-token list (H-11) |
| R11 | New team or new block added to the form | Team Christine appeared in August | Medium | Unknown leader/block → review. Mapping preview flags new headers |
| R12 | Header renamed or reordered | Trailing spaces and wording drift between v1 and v2 | Medium | Normalised alias matching. Fuzzy matches are *suggested only*, never auto-accepted |
| R13 | Overwriting the live destination | — | High | Read-only inputs. New output file. Path-equality check |
| R14 | PII leakage | Client names and phones | High | Local only. No network calls. Streamlit telemetry off and bound to `localhost`. Logs hold row numbers and codes, never values |
| R16 | Misspelled but syntactically valid e-mail domains pass validation | a seller e-mail with a misspelled `gmail` domain in the destination sample | Low | Accepted under H-24: syntax-only checks, no domain guessing |
| R15 | Time zone ambiguity | Timestamps have no time zone | Low | Treated as local form time (Asia/Manila assumed). Documented, no conversion |

---

## F. Parser Architecture and User Workflow

### F.1 Technology choice

| Option | Pros | Cons | Verdict |
|---|---|---|---|
| CLI only | Smallest, scriptable | No mapping preview for a non-developer | Kept as a secondary entry point |
| Tkinter (stdlib) | No extra installs | Weak table previews, more UI code to maintain | Not recommended |
| **Streamlit (local)** | Fast to build table previews, filters and download buttons. Runs on localhost. Free | Extra dependency; runs in the browser | **Recommended** |
| Web app (React + API) | — | Enterprise overhead | Rejected (out of scope) |

Dependencies: `openpyxl`, `PyYAML`, `streamlit`, plus `pytest` for development. `pandas` is used only for display, not as the core data model. Python ≥ 3.11.

### F.2 Component architecture

```
            ┌──────────── app/streamlit_app.py ─────────────┐   cli.py
            │  thin UI: calls pipeline steps, shows results  │     │
            └───────────────────────┬────────────────────────┘     │
                                    ▼                              ▼
  fuel_parser/pipeline.py   (orchestrates; pure functions, no UI imports)
     │
     ├─ reader.py      open workbook read-only, list sheets, find header row, yield rows
     ├─ headers.py     normalise header → match key; alias + suggestion matching
     ├─ blocks.py      detect client blocks; select filled block(s) per response
     ├─ mapping.py     load/save mapping profiles (YAML); header-signature hashing
     ├─ transform.py   field rules: text, datetime, multi-select tokeniser, value maps
     ├─ validate.py    per-field and cross-field rules → Issue list
     ├─ dedupe.py      src_key, content_hash, D2/D3 grouping
     ├─ history.py     SQLite ledger (batches, keys, void)
     ├─ writer.py      build output workbook from destination template header
     └─ audit.py       reconciliation + audit workbook
  config/
     mapping_profile.yaml   destination fields → aliases, scope (response|block), rules
     teams.yaml             leader label → team_key, seller_slot, expected_blocks
     vocab.yaml             projects, lead sources, awareness options, value maps, null tokens
     sellers.yaml           (optional) seller name → e-mail directory
```

Core data model (dataclasses): `SourceResponse(row, ts, submitter, leader, fields)`, `ClientBlock(index, cells)`, `DestRecord(values[35], src_key, content_hash, source_row, block_index)`, `Issue(code, severity, source_row, field, message)`, `RunResult(records, issues, counts)`.

### F.3 Configuration sketch (specification, not code)

```yaml
# mapping_profile.yaml (excerpt)
destination_template_sheet: MASTERLIST
source_sheet_candidates: ["MARF RAW"]
fields:
  client_mobile:
    dest_header_key: "client's mobile no."
    scope: block
    source_aliases: ["mobile no. (client)"]
    type: text
    rules: [trim, mobile_text]
  lead_source:
    dest_header_key: "lead source"
    scope: response
    source_aliases: ["lead source"]
    value_map: lead_source_v1_to_v2
  own_source_others_text:
    scope: response
    source_aliases: ["if others, please specify below", "if others, please specify"]
    occurrence_after: "if own source, please choose below"
    requires: {field: own_source, equals: "Others"}
    dest: remarks_append        # pending decision H-8
```

### F.4 User workflow and wireframe

```
┌ FUEL Response Parser ────────────────────────────────────────────────┐
│ Steps: ① Source ② Sheet ③ Mapping ④ Validate ⑤ Preview ⑥ Export     │
├──────────────────────────────────────────────────────────────────────┤
│ ① Source workbook  [Browse…]  FUEL RESP v3.3.xlsx   sha256 ✓        │
│    ⚠ This file was processed on 2026-10-28 (batch #12)               │
│    Destination template [Browse…]  FUEL (MGM-AD ADMIN).xlsx           │
│ ② Sheet [MARF RAW ▾]  header row: 1   rows: 3,064   blocks found: 13  │
│ ③ Mapping  profile: MARF-v3 (matched signature)                      │
│   ┌──────────────────────┬───────────────────────┬──────┬────────┐   │
│   │ Destination field    │ Source header         │ Conf │ Action │   │
│   │ Completion time      │ Timestamp             │  H   │   ✓    │   │
│   │ Question             │ If Corporate…         │  L   │ [Confirm]│ │
│   │ (unmapped)           │ If Others…below 6     │  —   │ [Ignore]│  │
│   └──────────────────────┴───────────────────────┴──────┴────────┘   │
│ ④ Validate   errors 3 · warnings 412 · info 640    [View issues]     │
│ ⑤ Preview   team filter [All ▾]  first 200 rows of MASTERLIST output │
│ ⑥ Export   [Export import file]  [Export audit]   Last Id in live    │
│            MASTERLIST: [ 11 ]   → new Ids start at 12                │
│   Reconciliation: read 3,064 = exported 3,059 + review 3 + dup 2 ✓   │
└──────────────────────────────────────────────────────────────────────┘
```

UX decisions:
- Steps are gated: Export stays disabled until every L-confidence and unmapped item is confirmed or ignored. This enforces "never guess silently".
- A confirmed mapping is saved as a profile keyed by header signature. The next file with the same headers needs no re-confirmation, and changed headers are highlighted as a diff.
- Reconciliation is always visible before export.

### F.5 Outputs

1. `FUEL_MASTERLIST_IMPORT_<sourcestem>_<YYYYMMDD-HHMMSS>.xlsx`: sheet `MASTERLIST` only, holding the exact destination headers and exportable records. Ready to append or paste into the live file.
2. `FUEL_IMPORT_AUDIT_<same stamp>.xlsx` with these sheets:
   - `Summary`: reconciliation, file hashes, profile, batch id.
   - `Review`: flagged records with full values and reasons.
   - `Issues`: one row per issue.
   - `RowLog`: every source row → status, `src_key`, output Id.
   - `Mapping`: the mapping actually used, with confidence.
   - `Unmapped`: non-empty columns not mapped.
3. *(Optional, H-27)* `FUEL_LEADS_LOG_<same stamp>.xlsx`: the 17-column Leads Log projection (Section I), with the exact team-sheet headers, one sheet per team or one combined sheet. It is for validation and comparison only and never replaces MASTERLIST.

### F.6 Folder structure and naming conventions

```
FUEL-Gsheet-to-Excel-Parser-/
  README.md
  pyproject.toml
  config/ (mapping_profile.yaml, teams.yaml, vocab.yaml, sellers.example.yaml)
  fuel_parser/ (modules in F.2)
  app/streamlit_app.py
  cli.py
  tests/ (unit/, integration/, fixtures/builders.py)
  docs/DESIGN.md
  .gitignore            # *.xlsx, *.sqlite, output/ — real data never committed
```

- Python: `snake_case` modules and functions, `PascalCase` dataclasses.
- Issue codes: `UPPER_SNAKE` (e.g. `MULTI_BLOCK_RESPONSE`).
- Config keys: `snake_case` destination field ids, independent of header text.

### F.7 Performance

The source is about 3k × 168 cells. Streaming the read with `read_only=True` keeps memory low. Target: under 10 s end-to-end for 10k responses on a laptop. The run is single-pass and needs no database other than the ledger.

---

## G. Test Plan

All test fixtures are **synthetic workbooks generated by test code**. No real client data goes into the repository. An optional local "golden run" against the real file checks the expected counts.

| # | Case | Fixture | Expected result |
|---|---|---|---|
| T1 | Normal v2 response, one client | 1 row, block 6 filled | 1 record, every mapped field equal, Lead Source verbatim, no issues |
| T2 | Several filled blocks in one response | 1 row, blocks 1 and 6 filled | 2 records sharing response fields, both in Review with `MULTI_BLOCK_RESPONSE`, 0 in the import file, reconciliation balances |
| T3 | No block filled | 1 row, metadata only | 0 records, `NO_CLIENT_DATA` in Review |
| T4 | Seller slot by team | Leader configured to slot 2 | Name and e-mail in I/J. G/H and K/L blank |
| T5 | Seller e-mail from directory differs from submitter | Directory entry ≠ column B | Directory e-mail in the slot, `SUBMITTER_NOT_SELLER` warning |
| T6 | Missing seller name, e-mail present | Blank seller name | Error `SELLER_NAME_MISSING`. No e-mail written to any slot |
| T7 | Missing seller e-mail | No directory entry, blank submitter | Name written, e-mail blank, warning |
| T8 | Unknown team label | `TEAM NEW (AAD: X)` | Record → Review, `UNKNOWN_TEAM` |
| T9 | Team-specific layout | Fixture with blocks reordered and an extra team block | Same records as the canonical fixture |
| T10 | Multi-select awareness with commas in options | `Word of Mouth, Social Media (Facebook, Instagram, etc), Events` | `Word of Mouth;Social Media (Facebook, Instagram, etc);Events` |
| T11 | Unknown awareness token | `Word of Mouth, Radio` | Value kept as `Word of Mouth;Radio`, warning `UNKNOWN_OPTION` |
| T12 | Blank optional fields | Middle name, client e-mail and second choice blank | Blanks written, no errors |
| T13 | Null tokens | Client e-mail `N/A` | Per H-11 (default: blank + INFO) |
| T14 | Invalid or ambiguous dates | Text `03/04/2026 10:00`, `"yesterday"`, blank | Ambiguous → Review `AMBIGUOUS_DATE`. Unparseable → `INVALID_DATE`. Blank → `MISSING_TIMESTAMP` |
| T15 | Leading zeros | Mobile `09XXXXXXXXX` as text | Output cell is the string `09XXXXXXXXX` with format `@` |
| T16 | Mobile arrives numeric | Number `9XXXXXXXXX` | Written as text `9XXXXXXXXX` (no `0` added). WARNINGs `MOBILE_WAS_NUMERIC` + `MOBILE_FORMAT_UNEXPECTED[missing_leading_zero]`. Record **is exported** |
| T31 | Placeholder mobile | Text `9999999999` | Written exactly. `MOBILE_FORMAT_UNEXPECTED[missing_leading_zero, repeated_digits]`. Record exported and listed in Review |
| T32 | Malformed e-mail syntax | Seller e-mail `name@gmail,com` | Written exactly. WARNING `INVALID_EMAIL`. No correction |
| T33 | Unfamiliar or misspelled domain | `x@gmaol.com`, `x@company.ph` | No issue raised. Value written exactly |
| T34 | Seller in a non-configured slot | Leader configured to slot 1, data only in slot 2 | ERROR `SELLER_SLOT_MISMATCH`. Slot 2 values not moved. Record → Review only |
| T35 | Leads Log projection | Fixture I.4 | The 17-column output equals table I.4 cell for cell. TIMESTAMP is a datetime cell with format `mmm/dd/yyyy hh:mm:ss` |
| T17 | Same file processed twice | Run, export, run again | Banner `FILE_ALREADY_PROCESSED`. All rows `ALREADY_IMPORTED`. Export contains 0 records |
| T18 | Newer export with old and new rows | v3.3 + 5 new rows | Only the 5 new rows are exported |
| T19 | Exact-content duplicate | Two rows identical except timestamp | First exported, second Review `PROBABLE_DUPLICATE` |
| T20 | Same client, different seller | Same mobile and name | Both exported, `POSSIBLE_DUPLICATE_CLIENT`, same group id |
| T21 | Renamed header | `Mobile Number (Client)` | Suggested match shown as uncertain. Export blocked until confirmed |
| T22 | Unrecognised non-empty column | Extra column `Budget Range` | `UNMAPPED_COLUMN`. User must ignore or map it |
| T23 | Moved columns | Same headers, shuffled order | Output identical to the canonical fixture |
| T24 | v1 lead source | `Ownsource` + `Own Source`=`Friend` | `Own Source` / AB=`Friend` (after H-7 approval) |
| T25 | Untranslatable v1 value | `Site Walk-in` | Review `UNMAPPED_VALUE` |
| T26 | Others co-occurrence violation | `…below 3` text, but Own Source ≠ Others | Warning `ORPHAN_OTHERS_TEXT`. Text still preserved per H-8 |
| T27 | Formula-like text | Remarks `=SUM(A1)` | Written as a literal string |
| T28 | Header with nbsp and newline in the template | — | Output header bytes equal the template's |
| T29 | Inputs unchanged | Any run | SHA-256 of both inputs is identical before and after the run |
| T30 | Void batch | Export, void, re-run | Rows become exportable again |

**Reconciliation checks** (a run fails loudly if any check is false):

```
data_rows_read         = blank_rows + responses
records_built          = Σ filled blocks           (= responses − no_client + extra_blocks)
records_built          = exported + review_only + already_imported + duplicate_held
exported               = rows in MASTERLIST output
every source row       appears exactly once in RowLog
```

**Golden-run expectations for v3.3** (computed during this analysis, before any decisions): 3,064 responses read · 3,064 records built · 0 multi-block · 0 no-client · 2 `PROBABLE_DUPLICATE` · 38 `POSSIBLE_DUPLICATE_CLIENT` groups (77 rows) · 398 v1 responses · 0 unmapped non-empty columns once H-6/H-8 are settled.

---

## H. Decisions Requiring Approval

### Decision register (v0.3)

| # | Topic | Status |
|---|---|---|
| H-1 | Seller e-mail source | **Approved: option (b).** A directory with submitter fallback and a mismatch warning. The **draft directory is under review** (J.1). It is not used until reviewed |
| H-2 | Team → seller slot mapping | **OPEN: blocking → B1** (M.5). The rule is approved. Only export-supported or approved evidence counts. Unsupported teams → Review |
| H-3 | Destination Email / Name | **Approved**: Email = submitter e-mail. Name blank |
| H-4 | Start time | **Approved**: copy of the source Timestamp |
| H-5 | Id sequence | **Approved**: continue from the last Id you enter from the live MASTERLIST |
| H-6 | Meaning of `Question` / `Question1` | `Question1` = Showroom Walk-In detail: **approved**. `Question`: **OPEN → B3** (M.5). Cannot be verified from the exports |
| H-7 | v1 → v2 lead-source translation | **Approved** for the four translations in J.4. `Referral` and `Site Walk-in` stay **OPEN**, and those records go to Review (`UNMAPPED_VALUE`, ERROR) |
| H-8 | "Others, please specify" text | **Approved: option (a)**, appended to REMARKS with the originating-question label (J.5) |
| H-9 – H-12, H-14 – H-21 | Defaults | Accepted as recommended (no objection raised) |
| H-13 | WARNING vs ERROR routing | **Approved**, refined by D.8: warnings export and are listed in Review, errors go to Review only |
| X1 | SPECIFIC LEAD SOURCE from the paired detail column | **Approved** |
| X2 | SOURCE OF AWARENESS preserved, `;`-joined, comma-containing options kept whole | **Approved** |
| H-22 | TIMESTAMP is a real datetime, format `mmm/dd/yyyy hh:mm:ss` | **Approved** |
| H-23 | LEAD SOURCE (DISTINCT) blank until a rule is approved | **Approved** |
| H-24 | E-mail: syntax check only, no domain rejection or correction | **Approved** |
| H-25 | Trim leading/trailing whitespace, logged | **Approved** |
| H-26 | (v0.2 "every flag → review") | **Withdrawn**, replaced by the D.8 validation policy |
| H-27 | 17-column Leads Log as an optional separate output | **Approved** |
| H-28 | Source SALES LEADER labels vs the MS Form's SALES LEADER options | **OPEN → B2** (M.5). Cannot be verified from the exports |
| H-29 | Access model: user-provided `.xlsx`/`.csv` exports only | **Approved** (Section M) |

### Blocking (the build needs an answer)

| # | Decision | Options | Recommendation and reason |
|---|---|---|---|
| H-1 | Where does **Seller's Email Address** come from? | (a) Submitting e-mail (col "Email Address") · (b) Seller directory, with the submitting e-mail as a fallback and a mismatch warning · (c) Leave blank | **(b)**. The submitting account differs from the seller's usual one in about 7% of rows, so (a) mis-pairs some sellers. (b) stays correct as the directory fills in. Please supply or approve `sellers.yaml` (I can generate a draft from the most-used e-mail per seller for your review). |
| H-2 | **Seller slot** (G/H vs I/J vs K/L) per team | (a) Everyone in slot 1 · (b) Team → slot table | **(b)** if the MGM form branches by team (the sample suggests it does). Please give the slot for each of the 6 source teams. |
| H-3 | Destination **Email** (D) / **Name** (E) | D = submitter e-mail, or `anonymous` like MS Forms. E blank | D = submitter e-mail (keeps the provenance), E blank. |
| H-4 | **Start time** | Copy Timestamp, or leave blank | Copy Timestamp. MS Forms always fills it, and blank values may break the live sheet's formulas or sorting. |
| H-5 | **Id** | (a) Continue from the last Id in the live MASTERLIST, entered by you · (b) Start at 1 · (c) Leave blank | **(a)**. The tool never reads or writes the live file, and you type the last Id. |
| H-6 *(see register)* | Meaning of **`Question`** (Y) and **`Question1`** (AF) | Y = Corporate detail? AF = Showroom detail? | Please confirm from the MS Form. AF is supported by sample data; Y is inferred only from question order. |
| H-7 | **v1 → v2 lead-source translation** | `Ownsource`→`Own Source` · `Booth`→`Booth Manning` · `Digital`+`Seller Facebook Page`/`Facebook Marketplace`/`Tiktok Inquiry`→`Digital - Seller Social Media` · `Digital`+`Project's Main Page`→`Digital - Official Social Media` · `Referral` (Friend/Relative)→? · `Site Walk-in`→? | Approve the first four. `Referral` and `Site Walk-in` → review until you decide (Referral holds a relationship, not a buyer name, and may really be Own Source = Friend/Relative). |
| H-8 | **"Others, please specify" text** (no destination column) | (a) Append to REMARKS as `[<Question> – Others] <text>` · (b) Write `Others: <text>` into the detail column · (c) Drop it (not recommended) | **(a)**. Lossless, and detail columns keep valid option values. |

### Defaults (proceed as recommended unless you object)

| # | Decision | Recommendation |
|---|---|---|
| H-9 | First-choice project with multiple values | Keep all values joined with `;`, plus warning `MULTI_FIRST_CHOICE`. Never move extras into Second Choice. |
| H-10 | Vocabulary differences (awareness `Print Ad` vs `Print Ads`, booth `NONO's` vs venue names, the v1 apostrophe variant) | Preserve source wording. Apply only value maps you approve (an initial `vocab.yaml` will be supplied for review). |
| H-11 | Null tokens (`N/A`, `NA`, `Na`, `N/a`, `None`, `-`) in client fields | Write blank, log as INFO. The original stays visible in the audit. |
| H-12 | Mobile number cell type | Text, format `@` (do not copy the destination's numeric samples). |
| H-13 | Records with **warnings** | Export them and also list them in Review. Records with **errors** go to Review only. |
| H-14 | Exact-content duplicates | Export the first occurrence and hold later ones in Review. |
| H-15 | Multi-block responses | Hold every resulting record in Review (none today). |
| H-16 | Output scope | One file for all teams, with an optional team filter in the UI. |
| H-17 | UI technology | Streamlit, local only (`localhost`, telemetry off). A CLI is provided as well. |
| H-18 | Config format | YAML (`PyYAML`). |
| H-19 | When rows count as "imported" | On successful export, with a "Void batch" undo. |
| H-20 | Team sheets | Ignored as input. `MARF RAW` is the only authoritative sheet. |
| H-21 | Seller-name prefixes (`AS: `, `AM: `, `AAD `) | Preserve verbatim. |

---

## Build Instructions (for the BUILD phase, after approval)

1. Create the project skeleton (F.6). Add `.gitignore` for `*.xlsx`, `*.sqlite` and `output/` before any data touches the repository.
2. Implement in this order: `headers` → `blocks` → `mapping`/config → `transform` → `validate` → `dedupe` → `history` → `writer` → `audit` → `pipeline` → CLI → Streamlit UI.
3. Write the synthetic fixture builders and T1–T30 alongside each module.
4. Run the golden run locally against v3.3 and compare it with the counts in G.

## Acceptance Criteria

- AC1: No mapping depends on column letters. T9 and T23 pass.
- AC2: The reconciliation equations hold on every run, and the run fails otherwise.
- AC3: The input files' SHA-256 hashes are unchanged after any run (T29).
- AC4: Every mobile number in the output is a text cell with its leading zero intact.
- AC5: The seller name and e-mail always occupy the same slot. No e-mail appears without its name.
- AC6: Re-processing an exported file exports 0 new records (T17).
- AC7: Every L-confidence or unmapped item needs explicit confirmation before export.
- AC8: Output headers are byte-identical to the template's `MASTERLIST` headers.
- AC9: The golden-run counts match Section G.
- AC10: No PII is written to logs or the ledger, and the tool makes no network calls.
- AC11: Mobile values are written exactly as received (as text). Unexpected formats raise a classified `MOBILE_FORMAT_UNEXPECTED` WARNING and do not block export (T16, T31).
- AC12: The seller slot is chosen only from the verified SALES LEADER config. A mismatch raises `SELLER_SLOT_MISMATCH`, and no value is ever moved between slots (T34).
- AC13: The Leads Log fixture (I.4) reproduces cell for cell (T35).

---

## I. Leads Log (17-column) Validation Fixture — v0.2 addendum

### I.1 What the 17-column format is

The 17 columns `TIMESTAMP … SOURCE OF AWARENESS` match **byte for byte** the header row (row 4, columns B–R) of the source workbook's team "LEADS LOG" sheets. That includes the trailing spaces in `TIMESTAMP `, `MOBILE NO. ` and `LEAD SOURCE `.

**Schema position (unchanged):** `MASTERLIST` (35 columns) stays the destination. The Leads Log is a **projection** of a MASTERLIST record. It is used as a validation fixture and is available as an optional, separate output (H-27 approved).

```
MARF RAW ──transform──► MASTERLIST record (35) ──project──► Leads Log row (17)
```

### I.2 Projection rules (MASTERLIST record → Leads Log row)

| # | Leads Log column | Taken from (MASTERLIST) | Rule |
|---|---|---|---|
| 1 | TIMESTAMP | Completion time | Datetime value kept unchanged. Display format `mmm/dd/yyyy hh:mm:ss` (renders `Oct/26/2026 17:38:01`). H-22 |
| 2 | SELLER EMAIL ADDRESS | Seller's Email Address **of the slot mapped to SALES LEADER** | Verbatim. Validated, never corrected |
| 3 | AAD | SALES LEADER | Verbatim |
| 4 | AS | Seller Name **of the same slot** | Verbatim. The prefix (`AS:`, `AM:`, `AAD:`) is **never** used to reassign the AAD/AS fields |
| 5–7 | SURNAME / FIRST / MIDDLE NAME (CLIENT) | Client's Surname / First / Middle | Verbatim apart from whitespace trimming (H-25). Blank stays blank |
| 8 | MOBILE NO. | Client's Mobile No. | Text. Numeric input → digits as text, **no leading 0 added**. Validated |
| 9 | EMAIL ADDRESS | **Client's** Email Address | Never the submitter or seller e-mail. Blank stays blank |
| 10 | RESIDENCE ADDRESS | Client's Residence Address | Verbatim |
| 11–12 | PROJECT (FIRST / SECOND CHOICE) | Project of Interest (First / Second Choice) | Verbatim, `;`-separated |
| 13 | REMARKS | REMARKS | Verbatim |
| 14 | LEAD SOURCE | Lead Source | Verbatim |
| 15 | SPECIFIC LEAD SOURCE | **Only** the detail column paired with the selected Lead Source (table I.3) | Blank if that column is blank or the lead source has no detail column. A value in a non-paired detail column → `ORPHAN_DETAIL` warning, not used |
| 16 | LEAD SOURCE (DISTINCT) | — | **Blank until a rule is approved** (H-23) |
| 17 | SOURCE OF AWARENESS | Source of Awareness | Verbatim, `;`-separated. Blank stays blank |

**Seller slot rule:**
- `teams.yaml` maps SALES LEADER → slot.
- Sample evidence: `TEAM CALOY` → slot 1 (G/H) and `TEAM HD` → slot 2 (I/J).
- If the configured slot is empty but another slot is filled → `SELLER_SLOT_MISMATCH` (review). The value is not moved.
- An unknown leader → `UNKNOWN_TEAM` (review).

### I.3 Lead Source → SPECIFIC LEAD SOURCE pairing

| Lead Source | Detail column (MASTERLIST) |
|---|---|
| Booth Manning | W (Booth Location) |
| Buyer Referral | X (referrer name) |
| Corporate | Y `Question` (pending H-6) |
| Digital - Official Social Media | Z |
| Digital - Seller Social Media | AA |
| Own Source | AB |
| Repeat Buyer | AC |
| Roadshow | AD |
| Saturation | AE |
| Showroom Walk-In | AF `Question1` (pending H-6) |
| Sponsorship Event | AG |
| Digital - Official Website, Marketing Partner, Others | none → blank |

### I.4 Fixture: five TEAM CALOY rows plus one slot-2 control

**Privacy:** names and e-mails are shown as ‹placeholders›. The test suite rebuilds this fixture as a **synthetic** workbook that reproduces each value's *pattern* (comma in domain, trailing space, 10-digit mobile, and so on). The real sample rows are never committed.

**Source:** the destination workbook's sample rows (Ids 6, 8, 9, 10, 11 = TEAM CALOY; Id 7 = TEAM HD as the slot-2 control). Ids 2–5 (`Option 1/2`) are form test entries and are excluded. *Assumption to confirm: these are the "five sample rows".*

Expected output under the rules above (`∅` = blank):

| Col | Id 6 | Id 8 | Id 9 | Id 10 | Id 11 | Id 7 (control) |
|---|---|---|---|---|---|---|
| TIMESTAMP | Oct/26/2026 17:38:01 | Oct/27/2026 10:26:58 | Oct/27/2026 10:32:17 | Oct/27/2026 14:54:20 | Oct/27/2026 15:55:30 | Oct/27/2026 10:23:31 |
| SELLER EMAIL | ‹e-mail with comma in domain› | ‹e-mail, misspelled gmail domain› | ‹valid e-mail› | ‹valid e-mail› | ‹valid e-mail› | ‹valid e-mail› (slot 2) |
| AAD | TEAM CALOY (AAD: …) | same | same | same | same | TEAM HD (AAD: …) |
| AS | ‹AAD:-prefixed name› | ‹AS: seller A› | ‹AS: seller B› | ‹AS: seller C› | ‹AS: seller B› | ‹AS: seller D› (slot 2) |
| SURNAME | ‹surname› | ‹surname› | ‹surname› | ‹surname› | ‹surname› | ‹surname› |
| FIRST NAME | ‹first name› (source has a trailing space, H-25) | ‹first name› | ‹first name› | ‹first name› | ‹first name› | ‹first name› |
| MIDDLE NAME | ∅ | ∅ | ∅ | ∅ | ∅ | ‹middle name› |
| MOBILE NO. | 9999999999 | 9999999999 | 9999999999 | 9999999999 | ‹10 digits, starts with 9› | 9999999999 |
| EMAIL ADDRESS | ∅ | ∅ | ∅ | ∅ | ∅ | ∅ |
| RESIDENCE | METRO MANILA - Manila City | METRO MANILA - Valenzuela City | LUZON - Baliwag City, Bulacan | METRO MANILA - Caloocan City | METRO MANILA - Las Piñas City | METRO MANILA - Muntinlupa City |
| PROJECT 1st | 1001 Parkway | The Levels | Botanika | Botanika | 1001 Parkway | Golf Ridge |
| PROJECT 2nd | Botanika | Timberland Heights;Commercial Lot - Filinvest City | Studio N | Botanika;Brentville;Golf Ridge;Studio N;The Levels | The Levels | Brentville |
| REMARKS | ∅ | ∅ | ∅ | ∅ | ∅ | ∅ |
| LEAD SOURCE | Digital - Seller Social Media | Showroom Walk-In | Digital - Official Website | Digital - Seller Social Media | Digital - Seller Social Media | Showroom Walk-In |
| SPECIFIC | **Seller Facebook Page** | **The Levels Sales Hub** | ∅ (no detail column) | Seller Facebook Page | Seller Facebook Page | Filigree Showsuites |
| DISTINCT | ∅ (H-23) | ∅ | ∅ | ∅ | ∅ | ∅ |
| AWARENESS | **Billboard / Lamp Post Banners** | Billboard / Lamp Post Banners | Email / Online Newsletter;Events;Google / Youtube Ads;Mall Kiosk / Booths;Print Ads;Roadshow;Sales Materials - Brochure / Flyers;Social Media (Facebook, Instagram, etc);Word of Mouth | Events | Mall Kiosk / Booths | Events |

**Expected flags** under the D.8 policy. Values are never changed. None of these is an ERROR, so all six rows are **exported** and also listed in the Review report.

| Id | Code | Field | Reason |
|---|---|---|---|
| 6 | `INVALID_EMAIL` (WARNING) | Seller e-mail | Comma in the domain (`…@gmail,com`) |
| 6, 7, 8, 9, 10 | `MOBILE_FORMAT_UNEXPECTED` [`missing_leading_zero`, `repeated_digits`] (WARNING) | Mobile | 10 digits, no leading `0`, one repeated digit |
| 11 | `MOBILE_FORMAT_UNEXPECTED` [`missing_leading_zero`] (WARNING) | Mobile | 10 digits, no leading `0`. A `0` is **not** added |
| all | `MOBILE_WAS_NUMERIC` (WARNING) | Mobile | Stored as a number in MASTERLIST |
| 8 | *(none)* | Seller e-mail | The misspelled domain (`…@gmaol.com`) has valid syntax. Under H-24 it is **not** flagged (R16) |
| 6 | `WHITESPACE_TRIMMED` (INFO) | First name | Trailing space (H-25) |
| 10 | `SECOND_INCLUDES_FIRST` (INFO) | Projects | Botanika appears in both choices |
| 9 | *(none expected)* | Awareness | `Print Ads` and `Roadshow` come from the MS Forms vocabulary, not MARF (`Print Ad`). The known-option list is the **union** of both forms, so no `UNKNOWN_OPTION` is raised |

### I.5 Differences found against the supplied sample

| # | Supplied expectation | What the data and stated rules give | Resolution needed |
|---|---|---|---|
| X1 | Id 6 SPECIFIC LEAD SOURCE = blank | The paired detail column AA holds `Seller Facebook Page`. Requirement 7 means it should be populated | **Resolved (X1 approved)**: populated from the paired detail |
| X2 | Id 6 SOURCE OF AWARENESS = blank | AH holds `Billboard / Lamp Post Banners`. Requirement 8 preserves non-blank values too | **Resolved (X2 approved)**: preserved, `;`-joined |
| X3 | First name without a trailing space | The source value has a trailing space | **Resolved (H-25)**: trimmed, `WHITESPACE_TRIMMED` logged |
| X4 | TIMESTAMP shown as `Oct/26/2026 17:38:01` | Equals **Completion time**, not Start time (17:35:38) | **Resolved (H-22)**: Completion time → datetime cell, `mmm/dd/yyyy hh:mm:ss` |
| X5 | Showroom Walk-In example | Its detail sits only in the unnamed `Question1` column | Still open under H-6 until confirmed from the MS Form |
| X6 | Malformed e-mails flagged | The `…@gmaol.com` address passes any syntax check | **Resolved (H-24)**: syntax only. Not flagged (R16) |
| X7 | — | The source's existing team-sheet formula leaves SPECIFIC blank for **all v2 Own Source rows** (the `If Own Source` column is missing from its ranges). It puts *Others* texts and the Seller-FB-Page boosting choice in DISTINCT | The team sheets cannot be a fully trusted comparison. They are compared case-insensitively, except SPECIFIC/DISTINCT |

### I.6 Decisions raised in v0.2 (status in the H register)

| # | Decision | Outcome |
|---|---|---|
| H-22 | TIMESTAMP cell type | **Approved**: real Excel datetime, format `mmm/dd/yyyy hh:mm:ss` |
| H-23 | LEAD SOURCE (DISTINCT) | **Approved**: blank until an explicit rule is approved. (Candidate on record: the second-level detail, as the team sheets use) |
| H-24 | E-mail checks | **Approved**: syntax only (`INVALID_EMAIL`). No domain rejection, no typo detection, no correction. `SUSPECT_EMAIL_DOMAIN` is dropped |
| H-25 | Leading and trailing whitespace | **Approved**: trim, log `WHITESPACE_TRIMMED` (INFO). Inner spacing untouched |
| H-26 | Routing of flagged records | **Withdrawn**. Replaced by the D.8 validation policy (mobile and e-mail format issues are WARNINGs, not blocking) |
| H-27 | Leads Log output | **Approved**: an optional, separate 17-column output for validation and comparison, with the exact headers. MASTERLIST (35 columns) unchanged |


---

## J. Verification Evidence and Approved Rules — v0.4 addendum

### J.1 Seller directory (H-1): draft for review

**How the draft was built.** For each distinct `Seller Name` in `MARF RAW`, the parser-to-be counts the submitting e-mails and proposes the most-used one. Each entry carries this evidence:

| Column | Meaning |
|---|---|
| `proposed_email`, `share_pct` | Most-used submitting e-mail and its share of that seller's responses |
| `proposed_email_matches_name` | Whether the e-mail's local part contains a token of the seller's name |
| `other_emails`, `alt_email_matching_name` | Every other submitting e-mail with counts, and any that match the seller's name |
| `email_also_used_for` | Other sellers submitted from the same account (e.g. a team leader submitting for members) |
| `teams`, `last_used` | Context for the reviewer |
| `review_flags` | `TIE`, `LOW_SHARE` (< 80%), `EMAIL_USED_FOR_OTHER_SELLERS`, `NAME_MISMATCH`, `BETTER_CANDIDATE_IN_ALTERNATES`, `FEW_RESPONSES`, `MULTIPLE_TEAMS` |
| `status`, `reviewer_email`, `reviewer_note` | Filled in by the reviewer |

**Result on v3.3:**
- 30 sellers. 20 are `OK` (proposed e-mail has ≥ 80% share, matches the seller's name, and isn't shared).
- 10 need a decision: 2 ties, 6 low-share, 5 accounts also used for other sellers, 2 name mismatches. Several entries carry more than one flag.

**Privacy:** the draft holds real names and e-mails. It is delivered **outside the repository** for review. The real `config/sellers.yaml` is git-ignored, and only `config/sellers.example.yaml` (fake values) is committed.

**Approval rule:**
- The parser uses a directory row only when `status = APPROVED`.
- A missing, unapproved or `REJECTED` row falls back to the submitter e-mail with `SELLER_EMAIL_FROM_SUBMITTER` (INFO).
- When an approved directory e-mail differs from the submitter, the record gets `SUBMITTER_NOT_SELLER` (WARNING).
- The directory's file hash is recorded in the audit Summary.

### J.2 Team → seller slot (H-2) and leader labels (H-28)

**Rule (approved):**
- The slot comes only from a verified `teams.yaml` entry for the exact SALES LEADER label.
- An unverified or unknown leader → `UNKNOWN_TEAM` (ERROR).
- Expected slot empty but another slot filled → `SELLER_SLOT_MISMATCH` (ERROR). Nothing is ever moved between slots.

**Evidence available today:**

| SALES LEADER (source / destination) | In source | In destination sample | Slot evidence | Proposed slot | Status |
|---|---|---|---|---|---|
| `TEAM CALOY (AAD: …)` | no | yes (5 rows) | All 5 rows in G/H | 1 (G/H) | **Sample-confirmed**, still to be checked against the form |
| `TEAM HD (AAD: …)` | no | yes (1 row) | Row in I/J | 2 (I/J) | **Sample-confirmed** (single row), still to be checked against the form |
| `Option 1`, `Option 2` | no | yes (4 rows) | All in G/H | — | Placeholder options from form testing. Ignored |
| `TEAM LESLIE (AAD: …)` | 1,023 | **no** | none | **none proposed** | UNVERIFIED |
| `TEAM FERDINAND (AAD: …)` | 542 | **no** | none | **none proposed** | UNVERIFIED |
| `TEAM BERNADETTE (AAD: …)` | 518 | **no** | none | **none proposed** | UNVERIFIED |
| `AM: JESSA` | 442 | **no** | none | **none proposed** | UNVERIFIED |
| `TEAM CHRISTINE (AAD: …)` | 380 | **no** | none | **none proposed** | UNVERIFIED |
| `TEAM RALPH (AAD: …)` | 159 | **no** | none | **none proposed** | UNVERIFIED |

**Why no slot is proposed for the six source teams:**
- None of the six labels appears in the destination workbook.
- The source block number tells us nothing about the destination: the source form has 13 blocks and the destination form has 3 slots, so sections must be shared across leaders.
- This also raises **H-28**: the MS Form (an MGM-AD form) may not offer these CSP-AD leaders as options at all, or may word them differently. The parser writes SALES LEADER verbatim (H-21), so a label the live MASTERLIST does not recognise would be structurally valid but wrong.

**How to verify each mapping** (in order of authority):

> **v0.6:** the user has no access to the forms or live sheets. Methods 1 and 2 are available only if a form owner supplies the result, and test submissions are not used at all. See M.4 for the export-only alternatives.

1. **Form editor (authoritative, read-only; requires someone with form access).** In Microsoft Forms, open the MGM-AD form → *…* → **Branching**, then:
   - On the SALES LEADER question, note the "Go to" section for **every** option, and the exact option text (this answers H-28 too).
   - For each target section, note which seller question it contains. MS Forms exports repeated titles in **form order**, with a numeric suffix:

     | Seller question in form order | Export columns | Slot |
     |---|---|---|
     | 1st `Seller Name` | `Seller Name` / `Seller's Email Address` | 1 (G/H) |
     | 2nd | `Seller Name1` / `…Address1` | 2 (I/J) |
     | 3rd | `Seller Name2` / `…Address2` | 3 (K/L) |

   - Record the result in `teams.yaml` with `verified_by: form_branching` and the date.
2. **Test submission: withdrawn (v0.6).** Not used. Tests use synthetic workbooks only.
3. **Live data cross-check (supporting).** Optionally provide a read-only copy of the live MASTERLIST. A pre-flight report then counts SALES LEADER × filled slot. Every leader must appear in exactly one slot. Any leader in two slots is a contradiction to resolve before import.

`teams.yaml` entries carry `slot`, `destination_label`, `verified_by` and `verified_on`. The allowed `verified_by` values are redefined in M.4. A **runtime gate** holds in Review (`UNVERIFIED_TEAM_SLOT`, ERROR) any record whose leader entry is not exportable. Sample-only entries (CALOY, HD) are used in tests but are not exportable until verified.

### J.3 `Question` (H-6)

- `Question1` → Showroom Walk-In detail: **approved** (supported by both sample rows with Lead Source `Showroom Walk-In`).
- `Question` → **unmapped** until verified. No source value is written to column Y. If the source has a non-empty value that would have gone there (the `If Corporate, please choose below:` column, currently empty for all 3,064 responses), the record gets `UNMAPPED_COLUMN` and goes to Review.
- **How to verify:** in the form editor, find the question between "Name of existing buyer who referred the client to you" and "If Digital - Official Social Media". MS Forms exports untitled or duplicate-titled questions as `Question`, `Question1`, … in form order. Note its type, options and which Lead Source option branches to it. If it is the Corporate detail, the mapping in C row 25 is activated as written.

### J.4 Legacy lead-source translation (H-7, approved)

| v1 `Lead Source` | v1 detail condition | → v2 `Lead Source` | Detail written to |
|---|---|---|---|
| `Ownsource` | any | `Own Source` | AB (from v1 `Own Source`) |
| `Booth` | any | `Booth Manning` | W (from v1 booth location, verbatim) |
| `Digital` | `Digital` ∈ {`Seller Facebook Page`, `Facebook Marketplace`, `Tiktok Inquiry`} | `Digital - Seller Social Media` | AA (the v1 `Digital` value, verbatim) |
| `Digital` | `Digital` = `Project's Main Page` | `Digital - Official Social Media` | Z (verbatim) |
| `Digital` | any other value (e.g. `Others`) | — | Review, `UNMAPPED_VALUE` |
| `Referral` | — | **open** | Review, `UNMAPPED_VALUE` (ERROR) |
| `Site Walk-in` | — | **open** | Review, `UNMAPPED_VALUE` (ERROR) |

- Translated records keep their original v1 value in the audit RowLog (`lead_source_original`).
- Detail values are not translated, only placed. Any value-level mapping (e.g. `Facebook Marketplace` vs `Seller Facebook Marketplace Post`) still needs separate approval under H-10.
- Expected effect on v3.3: about 20 `Referral` and 1 `Site Walk-in` responses held in Review.

### J.5 "Others, please specify" text (H-8, approved)

Format, appended to REMARKS: `[<label> – Others] <text>`. Several entries, and any existing remark, are joined with ` | ` in a fixed order: original REMARKS first, then the table order below. The original text is never edited, only trimmed (H-25).

| Source column (resolved by occurrence + co-occurrence, C) | Label |
|---|---|
| v2 `If Others, please specify below` (after Corporate) | `Corporate` |
| v2 `…below 2` | `Digital - Seller Social Media` |
| v2 `…below 3` | `Own Source` |
| v2 `…below 4` | `Repeat Buyer` |
| v2 `…below 5` | `Lead Source` |
| v1 `If others, please specify` (after Own Source) | `Own Source` |
| v1 `If others, please specify 2` | `Referral` |
| v1 `If others, please specify 3` | `Digital` |
| v1 `If Others, please specify project` | `Repeat Buyer` |
| v1 `If others, please specify` (last) | `Lead Source` |

- Example: `[Own Source – Others] <text>`.
- If the parent answer is not `Others` (`ORPHAN_OTHERS_TEXT`, WARNING), the text is still appended, so nothing is lost.

### J.6 Build gate

BUILD may start once **all** of the following hold:
1. **B1 / H-2:** an approved evidence route for team → slot (M.4). Teams without one are held in Review at runtime rather than blocking the build.
2. **B2 / H-28:** approval of the SALES LEADER label policy (M.5).
3. **B3 / H-6:** approval of the `Question` policy (M.5).
4. H-1: the seller directory review is returned. Rows may stay unapproved; they fall back to the submitter e-mail.

`Referral` / `Site Walk-in` (H-7) do not block the build. Their records stay in Review (`UNMAPPED_VALUE`, ERROR) and are never exported automatically until a mapping is approved.

**Scope of the gate.** Gate items 1–3 block only the components that depend on them: team/slot resolution in `teams.yaml`, the `Question` column, and the final export. They do not block work on components that are independent of those decisions (reader, header matching, block detection, transforms, validation policy, de-duplication, ledger, audit). Implementation of any component still starts only after your explicit go-ahead.

**Test submissions** are **not used** (v0.6). Verification relies on exported files and approved evidence only (M.4).

---

## K. Microsoft Form Verification Checklist (read-only) — v0.5

> **v0.6: optional.** This checklist applies only if someone who owns or edits the MGM-AD form agrees to complete it. You can forward this section to them as is. Nothing in the parser depends on it. Without it, M.4 applies.

Do this in the MGM-AD Microsoft Form's editor, without submitting any responses. Record the results in the table at the end. Before building, those results go into `teams.yaml` (labels and slots) and the H-6 mapping.

**K.1 Open the form**
- [ ] Open the form that produces `FUEL (MGM-AD ADMIN)` in **edit** mode (not the response link).
- [ ] Open *… → Settings* and note whether responses are anonymous. This explains `Email = anonymous` in the export.

**K.2 SALES LEADER options (H-28)**
- [ ] Find the SALES LEADER question and copy every option's text **exactly**: spacing, capitals, parentheses, the `AAD:` / `AM:` prefix.
- [ ] For each of the six source teams, mark it *present (identical)*, *present (worded differently)* or *absent*.
- [ ] If any source team is absent, stop and decide whether CSP-AD responses belong in this MASTERLIST at all. This is a scope question, not a mapping question.

**K.3 Branching → seller slot (H-2)**
- [ ] Open *… → Branching*. On the SALES LEADER question, note the "Go to" target for **every** option.
- [ ] List the form's sections in order, and mark which ones contain a `Seller Name` question.
- [ ] Number the `Seller Name` questions in form order: the 1st exports to G/H (`Seller Name` / `Seller's Email Address`), the 2nd to I/J (`…1`), the 3rd to K/L (`…2`).
- [ ] Confirm that each `Seller Name` question is paired with its own `Seller's Email Address` in the **same** section.
- [ ] Confirm TEAM CALOY → 1st and TEAM HD → 2nd, so the sample-confirmed entries become form-verified.
- [ ] Check for a 4th `Seller Name` question. If one exists, the 35-column schema is missing a slot; report it rather than mapping it.

**K.4 `Question` and `Question1` (H-6)**
- [ ] Find the question between "Name of existing buyer who referred the client to you" and "If Digital - Official Social Media". Note its title (it may be blank), type and options.
- [ ] Note which **Lead Source** option branches to it (e.g. Corporate).
- [ ] Confirm that the question between "If Saturation…" and "If Sponsorship Event…" is the Showroom Walk-In detail (`Question1`).
- [ ] If `Question` cannot be identified, say so. Column Y then stays blank (allowed by the gate).

**K.5 Lead Source options (supports H-7 and the vocabulary)**
- [ ] Copy every Lead Source option exactly. Is there an option equivalent to the legacy `Referral` (Friend/Relative) or `Site Walk-in`?

**Results to return** (no personal data needed):

| Source SALES LEADER (redacted) | Exact form option text | Present? | Section reached | Seller question # → slot | Verified by / date |
|---|---|---|---|---|---|
| TEAM LESLIE (AAD: …) | | | | | |
| TEAM FERDINAND (AAD: …) | | | | | |
| TEAM BERNADETTE (AAD: …) | | | | | |
| AM: JESSA | | | | | |
| TEAM CHRISTINE (AAD: …) | | | | | |
| TEAM RALPH (AAD: …) | | | | | |
| TEAM CALOY (AAD: …) | | | | | |
| TEAM HD (AAD: …) | | | | | |

| Item | Finding |
|---|---|
| `Question` (Y): title, type, options, branched from | |
| `Question1` (AF) confirmed as Showroom Walk-In | |
| Lead Source options that could hold `Referral` / `Site Walk-in` | |

---

## L. Repository Privacy Controls — v0.5

- Tracked files are limited to `docs/DESIGN.md` and `.gitignore`. `*.xlsx`, `*.csv`, `*.sqlite`, `output/`, `config/sellers.yaml` and `config/teams.local.yaml` are git-ignored.
- The design uses only ‹placeholders› for names and e-mails. Team labels appear as *first-name team names* with the leader's full name redacted (`TEAM LESLIE (AAD: …)`). They are organisational labels needed for configuration.
- Test fixtures are synthetic. Real workbook data is never committed.
- Before each commit, the tracked tree is scanned against a term list built locally from the workbooks. The term list itself is never committed.
- History was rewritten on 2026-10-28 to remove an unredacted fixture from earlier commits (see the delivery notes for any residual exposure).


---

## M. Access Model: Exported Files Only — v0.6

### M.1 Requirement (H-29, approved)

The parser works **entirely on files the user provides**:
- **Source:** Google Sheets response exports, as `.xlsx` or `.csv`.
- **Template:** the destination `FUEL (MGM-AD ADMIN) (2).xlsx`.
- **Optional evidence:** further MS Forms exports of MASTERLIST.

It needs no Google or Microsoft account, API, credentials, network connection, or access to any live form or sheet. Every input is a **static snapshot**: headers and values are inspected at runtime, and the approved mapping configuration is then applied. Nothing in the design calls out to, or waits on, an external system.

### M.2 What changes in the design

| Area | Before | v0.6 |
|---|---|---|
| Source input | `.xlsx` with a `MARF RAW` sheet | `.xlsx` **or** `.csv`. In an `.xlsx`, `MARF RAW` is used when present. Otherwise the sheet is chosen by **header signature** (`Timestamp` + `SALES LEADER` + at least one `Seller Name` block) and confirmed in the preview. A `.csv` is treated as a single response sheet |
| `.csv` specifics | — | UTF-8 with or without BOM. Quoted multi-line cells supported. **All values arrive as text**, which suits mobile numbers. Timestamps are parsed with a configured format (Google exports use the sheet's locale, e.g. `M/D/YYYY H:MM:SS`). A day/month order that cannot be resolved → `AMBIGUOUS_DATE` (ERROR). The format is confirmed in the mapping preview from the first values |
| Team sheets | Ignored | **Optional reference only.** If present, an audit comparison of their cached values against the parser's Leads Log view (case-insensitive, excluding SPECIFIC/DISTINCT, I.5 X7). Never required, never an input, formulas never evaluated |
| Destination template | Header copied from the template | Unchanged. The template header row is also hashed and compared with the approved schema snapshot in config. A mismatch stops the run with a clear message. The live destination is never written |
| New or changed exports | Mapping preview | Unchanged and mandatory: new, renamed or moved headers, new blocks and new form versions are shown as a diff against the saved profile. Ambiguous items must be confirmed before export |
| Seller slots | Verified from the form | Export-supported or explicitly approved evidence only (M.4). Otherwise → Review |
| Test data | Synthetic + golden run | Unchanged: synthetic workbooks, plus a local golden run on the provided v3.3 export. No live-form submissions |

### M.3 What can and cannot be verified from the available files

| Question | Answerable from exports? | Basis |
|---|---|---|
| Source structure, blocks, form versions, lead-source vocabulary | **Yes** | `MARF RAW` (sections A, C) |
| Which client block a team fills in the **source** | **Yes** | A.3 |
| Seller name ↔ submitter e-mail evidence | **Yes**, as evidence | J.1 draft directory (needs human approval) |
| Destination schema, header text, formats | **Yes** | Template (section B) |
| `Question1` = Showroom Walk-In | **Yes**, from sample rows | Approved |
| Slot for TEAM CALOY / TEAM HD | **Partly**: 5 rows and 1 row | Sample only |
| Slot for the six source teams (H-2) | **No** | No destination rows exist for them |
| Whether the MS Form offers the six source leaders, and their exact wording (H-28) | **No** | No destination rows exist for them |
| Meaning of `Question` (H-6) | **No** | Never populated in either file |
| Destination option lists (full Lead Source / awareness vocab) | **Partly** | Only values seen in the 10 sample rows |

### M.4 Alternatives that do not require form or live-sheet access

`teams.yaml` → `verified_by` now accepts these **evidence types**. Each entry records which one applies and who approved it, and the audit Summary lists them.

| Evidence type | What you would provide | Strength | Exportable? |
|---|---|---|---|
| `masterlist_export` | A fuller **exported copy** of the live MASTERLIST (any `.xlsx` downloaded from the MS Form or its Excel file) that contains rows for the team. The parser's pre-flight counts SALES LEADER × filled slot. The leader must appear in exactly **one** slot across ≥ 3 rows, with no contradictions | Strong (observed behaviour) | Yes |
| `owner_confirmation` | A written statement (e-mail or screenshot) from whoever administers the MGM-AD form, giving each option's text and slot. Section K can be forwarded to them as the request | Strong (authoritative, second-hand) | Yes |
| `business_rule` | An explicit decision **by you**, e.g. "all CSP-AD teams go to slot 1 (G/H)". Recorded as a business decision, not as a verified fact. The audit labels every affected record `SLOT_BY_BUSINESS_RULE` (INFO) | Policy, not verification | Yes, once approved |
| `sample_only` | The 10-row destination sample (CALOY, HD) | Weak | **No** (fixtures only) unless upgraded by one of the above |
| *(none)* | — | — | No → every record for that team is held in Review (`UNVERIFIED_TEAM_SLOT`) |

**Consequence to be aware of:** with no evidence at all, all 3,064 current responses would be held in Review, because all of them belong to the six unverified teams. The parser would still run and audit correctly, but the import file would be empty.

**H-28 labels without form access.** Either:
- (a) write the source SALES LEADER text verbatim (H-21), accepting that it may not match the form's wording; or
- (b) supply a label map (source label → destination label) taken from a MASTERLIST export or an owner confirmation.

**H-6 `Question` without form access.** The source column it would map from (`If Corporate, please choose below:`) is empty in all 3,064 responses. A policy therefore resolves it safely: column Y is always written blank, and any future non-empty Corporate value → `UNMAPPED_COLUMN` (ERROR, Review). Nothing is guessed, and nothing is lost.

### M.5 Remaining blocking decisions (exports-only)

| # | Decision | Options | Recommendation |
|---|---|---|---|
| **B1** (H-2) | Evidence route for team → seller slot, per team | `masterlist_export` · `owner_confirmation` · `business_rule` · hold in Review | Ask for a **MASTERLIST export** first. If none exists for these teams, choose between a `business_rule` and holding records in Review. A rule makes the output usable now. Holding keeps it strictly verified but empty |
| **B2** (H-28) | SALES LEADER label written to the destination | (a) source text verbatim · (b) label map from evidence | **(a)** now. Switch to (b) for any team where a MASTERLIST export or owner confirmation shows different wording |
| **B3** (H-6) | `Question` (column Y) | (a) always blank, future Corporate values → Review · (b) wait for owner confirmation | **(a)**. Resolves the blocker without guessing. It can be revisited if evidence appears |

**Not blocking:**
- **H-1:** the directory review. Unapproved rows fall back to the submitter e-mail.
- **H-7:** `Referral` and `Site Walk-in` stay in Review.
- **Repository visibility / GitHub purge (L):** a repository-owner action, independent of the build.

### M.6 Added tests

| # | Case | Expected |
|---|---|---|
| T36 | `.csv` export of the synthetic fixture | Same records as the `.xlsx` fixture. Mobile values identical as text |
| T37 | `.csv` with ambiguous `03/04/2026` timestamps and no confirmed format | `AMBIGUOUS_DATE` (ERROR). Nothing exported until the format is confirmed |
| T38 | `.xlsx` without a `MARF RAW` sheet but with a matching header signature | Sheet proposed in the preview. Processing only after confirmation |
| T39 | Workbook with only team sheets (no raw responses) | Stops: "no raw response sheet found". Derived sheets are never used as input |
| T40 | Team with `verified_by: business_rule` | Exported, each record tagged `SLOT_BY_BUSINESS_RULE` (INFO), listed in the audit Summary |
| T41 | Team with no evidence | All its records → Review `UNVERIFIED_TEAM_SLOT`. Reconciliation still balances |
| T42 | MASTERLIST-export pre-flight shows a leader in two slots | Contradiction reported. Team not exportable until resolved |
| T43 | Template header altered | Run stops with a schema-mismatch message. No output written |
| T44 | Run with networking disabled | Completes normally (no external calls) |
