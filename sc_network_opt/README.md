# Supply Chain Network Optimisation & Production Scheduler

A **Mixed Integer Linear Programme (MILP)** that maximises profit across a multi-echelon supply chain — simultaneously scheduling production, routing distribution, and managing inventory pre-build across plants, CFAs, and customers over a 12-month horizon.

Built in pure Python using **scipy HiGHS** as the branch-and-bound solver. Results are delivered as a formatted, colour-coded Excel workbook.

---

## Results on the included sample

| Metric | Value |
|---|---|
| Revenue | Rs 5,02,51,14,667 |
| RMPM Cost | Rs 1,93,40,55,992 |
| Production Cost | Rs 34,15,10,948 |
| Transport Cost | Rs 26,56,76,153 |
| Handling Cost | Rs 36,64,560 |
| Changeover Cost | Rs 17,83,075 |
| Holding Cost | Rs 1,09,69,911 |
| Stockout Penalty | Rs 2,57,46,667 |
| **Net Profit** | **Rs 2,44,17,07,361** |
| Profit Margin | 48.6% |
| Demand Fulfilment | 97.5% |
| Cases Produced | 68,15,200 |
| Avg Line Utilisation | 59.0% |
| Solve time | < 1 second (HiGHS MILP) |

---

## Network structure

```
Plant_A ──► CFA_North ──► Cust_Delhi, Cust_Chandigarh
        ──► CFA_South ──► Cust_Chennai, Cust_Bangalore
        ──► CFA_East  ──► Cust_Kolkata, Cust_Bhubaneswar
        ──► CFA_West  ──► Cust_Mumbai, Cust_Ahmedabad

Plant_B ──► CFA_North ──► Cust_Delhi, Cust_Chandigarh
        ──► CFA_South ──► Cust_Chennai, Cust_Bangalore
        ──► CFA_East  ──► Cust_Kolkata, Cust_Bhubaneswar
        ──► CFA_West  ──► Cust_Mumbai, Cust_Ahmedabad
```

Node types are auto-detected from name prefix:

| Prefix | Type | Role in model |
|---|---|---|
| `Plant_` | Production node | Has lines, batching, inventory; can pre-build |
| `CFA_` | Intermediate warehouse | Carries inventory; incurs handling cost |
| `Cust_` | Customer / depot | Demand node only; no inventory variable |

---

## Model formulation

**Solver:** `scipy.optimize.milp` — HiGHS branch-and-bound

### Decision variables

| Variable | Domain | Description |
|---|---|---|
| `batches[plant, line, sku, month]` | Integer ≥ 0 | Batches produced per line per period |
| `setup[plant, line, sku, month]` | Binary {0,1} | Line setup indicator |
| `flow[source, dest, sku, month]` | Continuous ≥ 0 | Cases moved on each open lane |
| `inventory[node, sku, month]` | Continuous ≥ 0 | Closing stock at plant or CFA |
| `unmet[customer, sku, month]` | Continuous ≥ 0 | Unmet demand (penalised in objective) |

### Objective — maximise profit (minimise negated profit)

```
Profit = Revenue
       − RMPM Cost          (plant / line / SKU specific)
       − Production Cost    (plant / line / SKU specific)
       − Transport Cost     (per open lane, per case)
       − Handling Cost      (per case arriving at plant or CFA)
       − Changeover Cost    (per setup event)
       − Holding Cost       (per case-month in inventory, node / SKU specific)
       − Stockout Penalty   (per unmet case)
```

### Constraints

| # | Constraint | Description |
|---|---|---|
| C1 | Inventory balance | `opening + produced + inflows − outflows = closing` at every plant and CFA |
| C2 | Demand satisfaction | `Σ inflows into customer + unmet = demand` per customer / SKU / month |
| C3 | Line capacity | `Σ batches × hours_per_batch ≤ available_hours` per line per month |
| C4 | Setup linking + min-run | `batches ≤ BIG_M × setup`; if setup=1, batches ≥ `min_batches` (from `Min_run_hrs`) |
| C5 | Safety stock floor | `closing_inventory ≥ ss_days × daily_demand` at every inventory node |
| C6 | Pre-build cap | Plant inventory ≤ total demand over next `Max_prebuild_months` |

**Hours per batch** (the C3 coefficient) is derived directly from your line data:

```
hours_per_batch = (Batch_size_cases / Run_rate_cases_per_hr) + CIP_hrs
```

This ensures capacity is consumed accurately — a line running at 250 cases/hr with a 6,000-case batch takes 24 run hours plus CIP, not an assumed constant.

---

## Repository structure

```
sc-network-optimisation/
├── Network_optimization.py   # Solver — edit INPUT_FILE here to run
├── sample_input.xlsx         # Working example (2 plants, 4 CFAs, 8 customers, 4 SKUs, 12 months)
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Installation

```bash
git clone https://github.com/<your-username>/sc-network-optimisation.git
cd sc-network-optimisation
pip install -r requirements.txt
```

Python 3.9 or higher is required.

---

## Usage

### Step 1 — Point the script to your data

Open `Network_optimization.py` and edit the two lines at the top of the Configuration section:

```python
INPUT_FILE  = "sample_input.xlsx"   # ← change to your file
OUTPUT_FILE = None                   # ← None = auto-name beside input
```

### Step 2 — Run

**Terminal:**
```bash
python Network_optimization.py
```

**Jupyter / IPython:**
```python
%run Network_optimization.py
```

The output workbook is saved in the same folder as the input, named `<input_name>_output.xlsx`.

### Step 3 — Open the output workbook

The output opens directly in Excel with a Dashboard sheet, colour-coded KPI tiles, and a tab for each plan element.

---

## Input workbook — 8 required sheets

The workbook must contain these sheets with exact names. Column names are also case-sensitive.

### SKU_Parameters

| Column | Type | Notes |
|---|---|---|
| `SKU` | Text | Unique code — must match exactly across all sheets |
| `Revenue_per_case_INR` | Number | Net price at customer gate |
| `Description` | Text | Informational only |

### Demand

One row per **Customer × SKU**. Month columns (Jan … Dec) hold cases demanded.

| Column | Type | Notes |
|---|---|---|
| `Customer` | Text | Must start with `Cust_` |
| `SKU` | Text | Must match SKU_Parameters |
| `Jan` … `Dec` | Integer | Use 0, not blank, for zero-demand months |

### Production_Lines

One row per **Plant × Line × SKU × Month**.  
Setting `Line_SKU_active = 0` for a month removes that cell entirely from the MILP — no variable, no constraint.

| Column | Type | Notes |
|---|---|---|
| `Plant` | Text | Must start with `Plant_` |
| `Line` | Text | Line identifier within the plant |
| `SKU` | Text | Must match SKU_Parameters |
| `Month` | Text | Must match Demand column headers |
| `Line_SKU_active` | 0 / 1 | 0 = blocked this period (maintenance, campaign, wrong format) |
| `Batch_size_cases` | Integer | MoQ per run — production is an integer multiple of this |
| `Run_rate_cases_per_hr` | Number | Actual line speed for this SKU |
| `Standard_batch_run_hrs` | Number | = Batch_size / Run_rate (informational; model recomputes) |
| `Min_run_hrs` | Number | Minimum block per setup — prevents 1-case runs after a costly CIP |
| `CIP_hrs` | Number | Clean-in-place time per setup |
| `Production_cost_per_case_INR` | Number | Variable manufacturing cost (labour, energy) — excludes RMPM |
| `RMPM_Cost_per_case_INR` | Number | Raw material + packaging cost — varies by line (yield, format) |
| `Changeover_cost_INR` | Number | One-time cost per setup event |

### Available_Hours

One row per **Plant × Line**. Month columns hold available production hours.

| Column | Type | Notes |
|---|---|---|
| `Plant` | Text | Must match Production_Lines |
| `Line` | Text | Must match Production_Lines |
| `Jan` … `Dec` | Number | Set 0 for planned shutdown months |

### Handling_Costs

One row per **Node × SKU**. Covers every plant and CFA.

| Column | Type | Notes |
|---|---|---|
| `Node` | Text | Plant or CFA name |
| `SKU` | Text | Must match SKU_Parameters |
| `Handling_cost_per_case_INR` | Number | Inbound receiving cost per case when stock arrives |
| `Inventory_holding_cost_per_case_per_month_INR` | Number | Finance + warehouse cost per case held per month |

### Safety_Stock

One row per **Node × SKU**.

| Column | Type | Notes |
|---|---|---|
| `Node` | Text | Plant or CFA name |
| `SKU` | Text | Must match SKU_Parameters |
| `Safety_stock_days` | Integer | Minimum days of cover. Plants: 2–4 days. CFAs: 5–14 days. |

### Cost_Parameters

Two-row parameter table. Edit the `Value` column only — do not rename the Parameter rows.

| Parameter | Description |
|---|---|
| `Stockout_penalty_per_case` | Penalty per unmet case. Set above revenue to force fulfilment. |
| `Max_prebuild_months` | Months ahead the solver may produce for future demand. |

### Network

One row per **Source → Destination × SKU × Month**.  
`Lane_open = 0` closes a lane for that period — the solver creates no flow variable for it.  
If multiple rows exist for the same (Source, Destination, SKU, Month), the **lowest cost** row wins.

| Column | Type | Notes |
|---|---|---|
| `Source` | Text | `Plant_` or `CFA_` node |
| `Destination` | Text | `CFA_` or `Cust_` node |
| `SKU` | Text | Must match SKU_Parameters |
| `Month` | Text | Must match Demand column headers |
| `Transport_cost_per_case_INR` | Number | All-in freight cost per case |
| `Lane_open` | 0 / 1 | 0 = closed this period (road block, permit, seasonal contract) |
| `Mode` | Text | Road / Rail / Air (informational only) |
| `Notes` | Text | Optional audit trail |

> **Critical:** every `Cust_` node must have at least one open lane in every month × SKU combination — either `Plant → Cust` direct or `Plant → CFA` + `CFA → Cust`. Missing lanes cause infeasibility and are reported in the diagnostic workbook.

---

## Output workbook — 11 sheets

| Sheet | Contents |
|---|---|
| **Dashboard** | 12 KPI tiles + monthly P&L snapshot (demand, fulfilment, revenue, costs, net profit per month) |
| **Profit Bridge** | Waterfall: Revenue → RMPM → Prod Cost → Gross Profit → Transport → Handling → EBITDA → Net Profit |
| **Production Schedule** | Integer batches per line / SKU / month with full hours breakdown (run hrs, CIP hrs, total hrs, min-run flag) |
| **Network Flow** | Cases on every active lane with transport cost, handling cost, revenue, and net contribution |
| **Lane Utilisation** | All lanes — open with actual flow; closed lanes highlighted in red |
| **Prebuild Plan** | Plant-level advance production earmarked for future months with holding cost |
| **Inventory Position** | Closing stock and holding cost at every plant and CFA per SKU per month |
| **Demand Fulfilment** | Demand vs fulfilment per customer / SKU / month — green ≥ 95%, yellow 80–95%, red < 80% |
| **Capacity Utilisation** | Line utilisation split into run hours and CIP hours — green ≤ 85%, yellow ≤ 95%, red > 95% |
| **Line Eligibility Matrix** | Calendar pivot (months as columns) showing which lines can produce which SKUs each month |
| **Eligibility Detail** | Full flat table for filtering — same data as the matrix, filterable in Excel |

---

## Infeasibility diagnostics

When the solver cannot find a feasible solution it runs a three-stage diagnostic automatically:

**Stage 1 — LP relaxation**  
Drops integrality and re-solves. If the LP is also infeasible the problem is structurally broken (missing lanes, insufficient capacity). If the LP is feasible, batch sizes are too coarse to satisfy all constraints as integers.

**Stage 2 — Constraint-group relaxation**  
Relaxes each of the six constraint groups independently and tests whether feasibility is restored. Groups whose relaxation fixes the problem are flagged as root causes.

**Stage 3 — Elastic LP**  
Adds one slack variable per constraint row and minimises total slack. Reports violation magnitude and worst individual constraint per group.

The diagnostic output workbook includes:
- Root cause classification (INTEGER vs STRUCTURAL)
- Constraint-group analysis table with violation magnitudes
- Capacity check: production capacity vs demand + safety stock per month and SKU
- Isolated customers: every Customer × SKU × Month with no open reachable lane

---

## Naming conventions

| Prefix | Detected as | Behaviour |
|---|---|---|
| `Plant_` | Production node | Batching, inventory, can source lanes |
| `CFA_` | Intermediate warehouse | Inventory only, no production |
| `Cust_` | Customer / depot | Demand only, no inventory variable |

Month names and SKU codes must be **identical** across all sheets (case-sensitive).

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `KeyError: 'Parameter'` | Template banner rows picked up as header | The script auto-repairs this — ensure `Parameter` exists in the sheet |
| Infeasible — isolated customer | A customer has no open lane in some period | Add lanes or open closed lanes in the Network sheet |
| Infeasible — capacity shortage | Total production capacity < demand + safety stock | Increase Available_Hours, reduce Safety_stock_days, or add a line |
| Fulfilment < 100% (feasible) | Lines fully utilised, demand cannot all be met | Increase capacity or reduce Stockout_penalty to accept partial fulfilment |
| Zero production for a SKU | `Line_SKU_active = 0` for all months, or no hours available | Check Production_Lines and Available_Hours for that SKU |
| High stockout penalty dominates P&L | Penalty set below revenue — solver prefers unmet demand | Set `Stockout_penalty_per_case` > `Revenue_per_case_INR` |

---

## Dependencies

| Package | Minimum version | Purpose |
|---|---|---|
| `numpy` | 1.21 | Array operations |
| `pandas` | 1.3 | DataFrame handling and Excel parsing |
| `openpyxl` | 3.0 | Reading input workbook; writing formatted output |
| `scipy` | 1.7 | `milp()` with HiGHS branch-and-bound solver |

---

## License

MIT License — free to use, modify, and distribute.
