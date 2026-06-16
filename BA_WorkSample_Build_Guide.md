# Business Analyst Work Sample — From-Scratch Build Guide

**Tool split:** Power Query (Excel) cleans everything → feeds the Excel D1 table and the Power BI D3 visual. PowerPoint holds the D2 recommendation and D4 questions. D5 lives in Excel with a one-page summary slide.

**Why this order:** clean once in Power Query, reuse the same tables everywhere. Raw data is never touched, every transform is listed and auditable, and Excel + Power BI read from the same source so they can't disagree.

---

## PHASE 0 — Project setup (5 min)

1. Make a folder, e.g. `C:\BA_Exercise\`, and drop all six CSVs in it. Keep the original filenames.
2. Open a new Excel workbook. Save it as `BA_WorkSample.xlsx` in that folder.
3. Create these tabs (you'll fill them as you go): `D1_Summary`, `D5_DataQuality`, `AI_Log`.

**AI_Log tab now:** the exercise explicitly asks you to log how you used AI. Add columns: *Step | Prompt summary | What the AI produced | What I changed/verified.* Fill one row each time you use AI. This earns transparency points.

---

## PHASE 1 — Clean the data in Power Query (Excel)

`Data ▸ Get Data ▸ From File ▸ From Text/CSV` for each file. After it previews, click **Transform Data** (not Load) to open Power Query. Do the steps below per query, then `Home ▸ Close & Load To… ▸ Only Create Connection` so raw data doesn't clutter the workbook.

### Query A — `PersonnelActions` (the fact table)

Do these transforms in order. Or open `Home ▸ Advanced Editor` and paste the M block (adjust the file path):

```m
let
    Source = Csv.Document(File.Contents("C:\BA_Exercise\Personnel_Action_Workload.csv"),
             [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    // 1. Fix Department ID losing its leading zero (41200 -> 041200)
    DeptText = Table.TransformColumns(Promoted,
               {{"Department ID", each Text.PadStart(Text.From(_), 6, "0"), type text}}),
    // 2. Remove the duplicate A-1004 extract row (keeps first per Action ID -> 23 rows)
    Dedup = Table.Distinct(DeptText, {"Action ID"}),
    // 3. Flag open/incomplete work
    AddOpen = Table.AddColumn(Dedup, "IsOpen",
              each if List.Contains({"Open","Pending"}, [Status]) then 1 else 0, Int64.Type),
    // 4. Flag rework (trust the rework flag, not error category)
    AddRework = Table.AddColumn(AddOpen, "IsRework",
                each if Text.Trim([Rework flag]) = "Y" then 1 else 0, Int64.Type)
in
    AddRework
```

**Check:** the query should show **23 rows**. If it shows 24, the dedup didn't apply.

### Query B — `Funding`

```m
let
    Source = Csv.Document(File.Contents("C:\BA_Exercise\Department_Funding.csv"),
             [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    DeptText = Table.TransformColumns(Promoted,
               {{"Department ID", each Text.PadStart(Text.From(_), 6, "0"), type text}}),
    // Standardize the inconsistent department name
    FixName = Table.ReplaceValue(DeptText,"East Campus Research Ops",
              "East Campus Research Operations", Replacer.ReplaceText, {"Department name"}),
    Types = Table.TransformColumnTypes(FixName,{
              {"Beginning budget", Int64.Type},{"Actuals to date", Int64.Type},
              {"Encumbrances", Int64.Type},{"Available balance", Int64.Type}}),
    // Recompute available balance to expose the row that doesn't tie (Temporary Support: 12,000 vs 15,000)
    Recalc = Table.AddColumn(Types,"Available_recalc",
             each [Beginning budget]-[Actuals to date]-[Encumbrances], Int64.Type),
    Flag = Table.AddColumn(Recalc,"Balance_ties",
           each [Available balance] = [Available_recalc], type logical)
in
    Flag
```

### Query C — `Capacity` (build as a small documented lookup)

The "Assigned departments" field is a delimited list, and "over capacity" is a judgment call, so a hand-built lookup of 6 staff is cleaner and more defensible than a fragile auto-split. In a blank tab make a table named `tbl_Capacity`:

| Staff | Util % | Status | Primary dept(s) |
|-------|--------|--------|-----------------|
| Jordan Lee | 133% | Over | 041200; 041220 |
| Riley Chen | 108% | Over | 041200; 041215 |
| Morgan Tate | 94% | Normal | 041200; 041210 |
| Casey Brooks | 61% | Normal | 041205; 041225 |
| Taylor Monroe | 80% | Partial back-up | All |
| Vacant Position | 0% | Vacant since Mar 2026 | 041200; 041215 |

Util % = Current hours ÷ Available hours. Document that formula on the tab.

### Build the D1 summary in Power Query

1. **Action aggregates:** Reference `PersonnelActions` (right-click ▸ Reference). `Transform ▸ Group By`, group on **Department ID**, add: `Actions` = Count Rows; `Open` = Sum of IsOpen; `Rework#` = Sum of IsRework. Add a custom column `Completed = [Actions] - [Open]` and `Rework% = [Rework#] / [Actions]`. Name this `Agg_Actions`.
2. **Funding aggregates:** Reference `Funding`. Group on **Department ID** with `Avail_Total` = Sum of Available balance. Separately filter to `Budget category = "Salaries and Benefits"` and group for `Avail_SB`. Merge those two. Name it `Agg_Funding`.
3. **Department spine:** Reference `Funding`, keep Department ID + Department name, `Remove Duplicates`. Name it `Dim_Dept`.
4. **Merge it all:** Start from `Dim_Dept`, `Merge Queries` with `Agg_Actions` on Department ID (Left Outer), expand. Merge again with `Agg_Funding`. Add a manual `Capacity` text column and `Risk` column per the answer key (these are documented judgments).
5. `Close & Load To…` ▸ Table ▸ put it on the **D1_Summary** tab.

**Verify against the answer key** (this is your correctness check):

| Dept | Actions | Completed | Open | Rework# | Rework% | Avail S&B | Avail Total |
|------|---------|-----------|------|---------|---------|-----------|-------------|
| 041200 | 9 | 7 | 2 | 4 | 44% | 129,500 | 149,700 |
| 041205 | 3 | 3 | 0 | 1 | 33% | 93,000 | 94,000 |
| 041210 | 4 | 3 | 1 | 2 | 50% | 74,200 | 74,200 |
| 041215 | 2 | 2 | 0 | 0 | 0% | 59,800 | 59,800 |
| 041220 | 3 | 2 | 1 | 2 | 67% | 0 | 80,000 |
| 041225 | 2 | 2 | 0 | 1 | 50% | 40,000 | 40,000 |

Total actions = **23**. Add a Risk column: 041200 **HIGH**, 041215/041220 **MED**, rest **LOW**. Format the table, bold the East Campus row. **D1 done.**

---

## PHASE 2 — Power BI visual (D3)

1. Open Power BI Desktop ▸ `Get Data ▸ Text/CSV`, import the three source files. In the Power Query window inside Power BI, **paste the same M from Phase 1** (copy each query's Advanced Editor text over). Same clean tables, zero rework.
2. **Model view:** relate `PersonnelActions[Department ID]` → `Dim_Dept[Department ID]` (many-to-one). Relate `Capacity` to `Dim_Dept` only if you split assignments to rows; otherwise leave Capacity as a standalone visual table.
3. **DAX measures** (New Measure):

```DAX
Total Actions = COUNTROWS('PersonnelActions')
Open Actions  = SUM('PersonnelActions'[IsOpen])
Rework Count  = SUM('PersonnelActions'[IsRework])
Rework Rate   = DIVIDE([Rework Count], [Total Actions])
% of Workload = DIVIDE([Total Actions], CALCULATE([Total Actions], ALL('PersonnelActions')))
```

4. **One page only**, top-to-bottom for a non-technical reader:
   - **Headline text box:** "East Campus Research Operations handles ~39% of all personnel actions with an over-capacity analyst and a support role vacant since March 2026."
   - **Bar chart:** Actions by Department (sorted descending — East Campus towers). Color East Campus dark red, others grey.
   - **Two cards / small bar:** Jordan Lee 133% utilization; Vacant position 15 hrs/wk unfilled.
   - **Card:** $129,500 available Salaries & Benefits balance (East Campus).
   - **Recommendation box (bottom):** "Fill/reallocate the existing vacancy toward East Campus. Confirm funding before approving net-new FTE."
5. Keep it monochrome + one accent color, big fonts, no slicers the leader won't use. **Resist adding more pages** — the brief rewards a defensible point, not a busy dashboard. Export to PDF for submission.

---

## PHASE 3 — PowerPoint (D2 + D4)

Use the drafted text below verbatim or lightly edited. Plain language, one idea per line.

### Slide — D2: Recommendation

**Title:** Recommendation — Prioritize administrative support for East Campus Research Operations

**Bottom line:** The data supports directing additional administrative support to East Campus (Dept 041200) — most defensibly by **filling/reallocating the existing vacant support position**, not by approving net-new FTE yet, pending two funding clarifications.

**Evidence strong enough to act on**
- Workload concentration: East Campus handles **9 of 23** personnel actions (~39%); no other department exceeds 4.
- Staff over capacity: lead analyst is at **~133%** of available hours and flagged "over capacity"; a second assigned staffer is at ~108%.
- Existing vacancy: a **15 hr/week** support role assigned to East Campus has been **vacant since March 2026** — budgeted capacity not being delivered.

**Evidence NOT strong enough to act on alone**
- Per-department rework rates — samples of 2–4 actions are too small to compare reliably.
- The exact funding figure — a known balance-formula error plus two pending transfers (+$35K, −$12K) could move it.
- Whether the spike is temporary or sustained — only ~5 months of data.

**Recommended action:** Fill the vacant role and direct it primarily to East Campus. Revisit net-new FTE only after funding restrictions and pending transfers are confirmed.

### Slide — D4: Questions for stakeholders before deciding

- **Funding:** Will the pending transfers (T-5002 +$35K, T-5010 −$12K) change the available salary balance? Can Temporary Support / Grant / Auxiliary funds legally support a **recurring** position?
- **Vacancy:** Is the position vacant since March 2026 approved to be filled, and was it intended for East Campus?
- **Workload trend:** Is the elevated East Campus volume a temporary spike or an ongoing trend?
- **Data integrity:** Why do some actions reference employee IDs absent from payroll (e.g., 000000999, 000333444)? Are we processing work on invalid records?
- **Capacity:** Are the self-reported weekly hours validated, or should we time-track briefly before deciding?
- **Quality:** Are the rework/error fields reliable, given the flag-vs-category conflicts?

---

## PHASE 4 — D5 Data Quality, Assumptions & Limitations

Put the detailed log on the **D5_DataQuality** Excel tab; lift a short version onto a slide.

**Issues found (with location)**
- Duplicate row: A-1004 appears twice ("Duplicate extract row").
- Leading-zero loss: Dept "41200" in personnel + funding files (should be 041200).
- Name inconsistency: "East Campus Research Ops" vs "East Campus Research Operations."
- Available-balance formula doesn't tie on the Temporary Support row ($12,000 shown vs $15,000 calculated).
- Rework-flag vs error-category conflicts: A-1007 (flag N but error present), A-1023 (flag Y but blank).
- Employee ID format/validity: "ABC123456" (text), and IDs 000000999 / 000333444 not in payroll.
- Action on an inactive employee: A-1017 (employee 000111222 marked Inactive).
- Salary stored as text: "62,750" in payroll.

**Cleaning steps taken**
- Removed the duplicate by Action ID. Padded all Dept IDs to 6 digits. Standardized the department name. Recomputed available balance to expose the non-tying row. Trusted the rework flag over error category and footnoted the conflicts.

**Assumptions**
- Second A-1004 row is a true duplicate. "41200" = "041200." The two East Campus names are the same unit. Rework flag is the source of truth for rework.

**Limitations**
- Small sample (23 actions) — per-department rates are indicative, not conclusive. Capacity hours are self-reported, not time-tracked. Funding restrictions and the temporary-vs-sustained question can't be answered from these extracts.

**Fields not to rely on without clarification**
- Error category (conflicts with flag), Available balance (formula issue + pending transfers), self-reported weekly hours, and any funding-restriction adequacy.

---

## Submission checklist
- [ ] `BA_WorkSample.xlsx` — D1 summary matches the answer key, D5 log complete, AI_Log filled
- [ ] Power BI file + exported PDF (D3)
- [ ] PowerPoint deck (D2 + D4, plus optional D5 summary slide)
- [ ] One sentence in your cover note on how you used AI and what you verified
