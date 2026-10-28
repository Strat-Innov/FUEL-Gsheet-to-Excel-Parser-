# FUEL Response Parser — Design Specification (v0.1, for approval)

| Item | Value |
|---|---|
| Phase | DESIGN (no production code) |
| Source analysed | `FUEL RESP (CSP-AD MARF) v3.3.xlsx` (Google Forms response export) |
| Destination analysed | `FUEL (MGM-AD ADMIN) (2).xlsx` → sheet `MASTERLIST` (Microsoft Forms export) |
| Status | **Awaiting approval of the decisions in Section H** |

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
| 16 | Client's Mobile No. | `mobile no. (client)` | Filled block | **Text.** Trim only. If the reader receives a number (from a future file), convert to an integer string and **flag** `MOBILE_WAS_NUMERIC` (a leading 0 may already be lost). Cell format `@` | H | Pattern `^09\d{9}$` or `^\+639\d{9}$` → OK. Other patterns → warning. Blank → warning |
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
- **WARNING**: the record is exported and also listed in review (H-13).
- **INFO**: written to the audit log only.

The parser never fills a value it cannot trace to a source cell or an approved rule.

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
| T16 | Mobile arrives numeric | Number `9XXXXXXXXX` | Written as text `9XXXXXXXXX`, `MOBILE_WAS_NUMERIC` warning |
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

### Blocking (the build needs an answer)

| # | Decision | Options | Recommendation and reason |
|---|---|---|---|
| H-1 | Where does **Seller's Email Address** come from? | (a) Submitting e-mail (col "Email Address") · (b) Seller directory, with the submitting e-mail as a fallback and a mismatch warning · (c) Leave blank | **(b)**. The submitting account differs from the seller's usual one in about 7% of rows, so (a) mis-pairs some sellers. (b) stays correct as the directory fills in. Please supply or approve `sellers.yaml` (I can generate a draft from the most-used e-mail per seller for your review). |
| H-2 | **Seller slot** (G/H vs I/J vs K/L) per team | (a) Everyone in slot 1 · (b) Team → slot table | **(b)** if the MGM form branches by team (the sample suggests it does). Please give the slot for each of the 6 source teams. |
| H-3 | Destination **Email** (D) / **Name** (E) | D = submitter e-mail, or `anonymous` like MS Forms. E blank | D = submitter e-mail (keeps the provenance), E blank. |
| H-4 | **Start time** | Copy Timestamp, or leave blank | Copy Timestamp. MS Forms always fills it, and blank values may break the live sheet's formulas or sorting. |
| H-5 | **Id** | (a) Continue from the last Id in the live MASTERLIST, entered by you · (b) Start at 1 · (c) Leave blank | **(a)**. The tool never reads or writes the live file, and you type the last Id. |
| H-6 | Meaning of **`Question`** (Y) and **`Question1`** (AF) | Y = Corporate detail? AF = Showroom detail? | Please confirm from the MS Form. AF is supported by sample data; Y is inferred only from question order. |
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

---

## I. Leads Log (17-column) Validation Fixture — v0.2 addendum

### I.1 What the 17-column format is

The 17 columns `TIMESTAMP … SOURCE OF AWARENESS` match **byte for byte** the header row (row 4, columns B–R) of the source workbook's team "LEADS LOG" sheets. That includes the trailing spaces in `TIMESTAMP `, `MOBILE NO. ` and `LEAD SOURCE `.

**Schema position (unchanged):** `MASTERLIST` (35 columns) stays the destination. The Leads Log is a **projection** of a MASTERLIST record. It is used as a validation fixture and can optionally be exported per team (H-27).

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

**Expected flags.** Values are never changed; every flag sends the record to review.

| Id | Code | Field | Reason |
|---|---|---|---|
| 6 | `INVALID_EMAIL` | Seller e-mail | Comma in the domain (`…@gmail,com`) |
| 6, 7, 8, 9, 10 | `MOBILE_FORMAT` + `MOBILE_PLACEHOLDER` | Mobile | 10 digits, no leading `0`, and one repeated digit |
| 11 | `MOBILE_FORMAT` | Mobile | 10 digits, no leading `0` (prefixing `0` is **not** done) |
| all | `MOBILE_WAS_NUMERIC` | Mobile | Stored as a number in MASTERLIST |
| 8 | `SUSPECT_EMAIL_DOMAIN` | Seller e-mail | `gmaol.com` is valid syntax, so it is caught only by the domain-typo rule (H-24) |
| 6 | `WHITESPACE_TRIMMED` (INFO) | First name | Trailing space (H-25) |
| 10 | `SECOND_INCLUDES_FIRST` (INFO) | Projects | Botanika appears in both choices |
| 9 | `UNKNOWN_OPTION`? | Awareness | `Print Ads` and `Roadshow` exist in the MS Forms vocabulary but not in MARF (`Print Ad`). The vocabulary must be the **union** of both forms |

### I.5 Differences found against the supplied sample

| # | Supplied expectation | What the data and stated rules give | Resolution needed |
|---|---|---|---|
| X1 | Id 6 SPECIFIC LEAD SOURCE = blank | The paired detail column AA holds `Seller Facebook Page`. Requirement 7 means it should be populated | Confirm: populate from the paired detail (recommended), or leave blank |
| X2 | Id 6 SOURCE OF AWARENESS = blank | AH holds `Billboard / Lamp Post Banners`. Requirement 8 preserves non-blank values too | Confirm: preserve (recommended) |
| X3 | First name without a trailing space | The source value has a trailing space | H-25: trim with an INFO log (recommended), or keep exact |
| X4 | TIMESTAMP shown as `Oct/26/2026 17:38:01` | Equals **Completion time**, not Start time (17:35:38) | Confirms Completion time → TIMESTAMP. H-22 decides datetime cell vs text |
| X5 | Showroom Walk-In example | Its detail sits only in the unnamed `Question1` column | Treating AF as the Showroom detail closes H-6 for AF. `Question` (Y) is still open |
| X6 | Malformed e-mails flagged | The `…@gmaol.com` address passes any syntax check | H-24: add a configurable common-domain typo check, or accept that it will not be flagged |
| X7 | — | The source's existing team-sheet formula leaves SPECIFIC blank for **all v2 Own Source rows** (the `If Own Source` column is missing from its ranges). It puts *Others* texts and the Seller-FB-Page boosting choice in DISTINCT | The team sheets cannot be a fully trusted comparison. They are compared case-insensitively, except SPECIFIC/DISTINCT |

### I.6 New decisions

| # | Decision | Recommendation |
|---|---|---|
| H-22 | TIMESTAMP cell type | Datetime value with display format `mmm/dd/yyyy hh:mm:ss`. Sorting and filtering keep working, and the value is never re-typed. Alternative: a literal text string |
| H-23 | LEAD SOURCE (DISTINCT) | Leave blank. Candidate rule if you want one: the second-level detail, i.e. the Seller-FB-Page boosting choice (AI) or the *Others* specify text, as the existing team sheets do |
| H-24 | Plausible but misspelled e-mail domains | Warning `SUSPECT_EMAIL_DOMAIN` when the domain is within edit distance 1–2 of a configured list (gmail.com, yahoo.com, …) and is not on it |
| H-25 | Leading and trailing whitespace | Trim, and log `WHITESPACE_TRIMMED` (INFO). Inner spacing is untouched |
| H-26 | Records with these flags | Every listed flag sends the record to review **with its values unchanged**. This supersedes H-13 for `INVALID_EMAIL`, `SUSPECT_EMAIL_DOMAIN` and `MOBILE_FORMAT` |
| H-27 | Export the Leads Log view? | Optional per-team export with the exact 17 headers. Off by default |
