"""
Fair material budget calculator (Team A, FIBE3 x Ramboll x GBCe mini-project)

Implements the six-step chain in Chapter 10 for every material, indicator,
allocation principle, environmental-space option and demand scenario.

  1  Remaining space      = full space - occupied space            (formula 1)
  2  Spain budget         = space x population share               (formula 2)
                            x capability factor (ability to pay)   (formula 3)
  3  Building budget      = Spain budget x building share          (formula 6)
                            (needs share under human needs)        (formula 4)
                            capped by local capacity if supplied   (formula 13)
  4  Material env. budget = building budget x group share
                            x share within group                   (formula 7)
  5  Material budget (t)  = material env. budget / impact per t    (formula 9)
     Binding budget       = lowest budget across indicators        (formula 10)
  6  Overshoot ratio      = demand / binding budget                (formula 12)

Usage
  python fair_material_budget.py            reads fmb_inputs.xlsx, writes fmb_results.xlsx
                                            and three charts into fmb_charts/
  python fair_material_budget.py --test     checks the code against the worked
                                            examples published in the brief
  python fair_material_budget.py --template (re)writes fmb_inputs.xlsx with the
                                            illustrative values from the brief

All quantities are in base units (t CO2e, m3, t RME and so on) and tonnes of
material. Blank cells are treated as data gaps and reported, never filled in.
"""
import sys
from pathlib import Path

import pandas as pd

INPUTS = Path("fmb_inputs.xlsx")
RESULTS = Path("fmb_results.xlsx")
CHARTS = Path("fmb_charts")
PRINCIPLES = ["Equal per capita", "Ability to pay", "Human needs"]
SPACES = ["RSOS", "FSOS"]


# --------------------------------------------------------------------------
# Template holding the illustrative placeholder values from Chapter 10 and
# the Ramboll kick-off slides. Replace every value with verified data.
# --------------------------------------------------------------------------
def write_template(path=INPUTS):
    pop_es, pop_world = 48e6, 8.5e9
    up = pop_world / pop_es  # back-calculates a global space from a Spain figure

    parameters = pd.DataFrame(
        [
            ["population_spain", pop_es, "Ch10 formula 2 (illustrative)"],
            ["population_world", pop_world, "Ch10 formula 2 (illustrative)"],
            ["gdp_pc_spain", 30000, "Ch10 section 10.4.2 (illustrative)"],
            ["gdp_pc_world", 15000, "Ch10 section 10.4.2 (illustrative)"],
            ["capability_factor_override", 0.8,
             "Ch10 section 10.9.2 uses 0.80. Delete to use gdp_pc_world / gdp_pc_spain"],
        ],
        columns=["parameter", "value", "source_note"],
    )

    # full_space and occupied_space give RSOS. The brief only gives remaining
    # values, so occupied is 0 here and FSOS equals RSOS until real data go in.
    # budget_period_years: 1 for an annual budget, or the number of years a
    # cumulative budget must last (for example 24 for 2027 to 2050).
    indicators = pd.DataFrame(
        [
            ["Climate", "t CO2e", 10000e6, 0, 1, "Ch10 10.9.1"],
            ["Water", "m3", 900e6 * up, 0, 1, "Slide value for Spain, scaled up"],
            ["Biodiversity", "PDF.m2.yr", 1.2e6 * up, 0, 1, "Slide value for Spain, scaled up"],
            ["Toxicity", "CTUe", 240000 * up, 0, 1, "Slide value for Spain, scaled up"],
            ["Material footprint", "t RME", 160e6 * up, 0, 1, "Slide value for Spain, scaled up"],
        ],
        columns=["indicator", "unit", "full_space", "occupied_space",
                 "budget_period_years", "source_note"],
    )

    sector_shares = pd.DataFrame(
        [
            ["Climate", 0.15, 0.20, None],
            ["Water", 0.15, None, None],
            ["Biodiversity", 0.15, None, None],
            ["Toxicity", 0.15, None, None],
            ["Material footprint", 0.15, None, None],
        ],
        columns=["indicator", "building_impact_share", "needs_share",
                 "local_capacity_buildings"],
    )

    group_shares = pd.DataFrame(
        [
            ["Climate", "Mineral", 0.50],
            ["Water", "Mineral", 0.50],
            ["Material footprint", "Mineral", 0.50],
            ["Biodiversity", "Bio-based", 0.50],
            ["Toxicity", "Chemicals", 0.20],
        ],
        columns=["indicator", "group", "group_share"],
    )

    materials = pd.DataFrame(
        [
            ["Cement", "Mineral"],
            ["Steel", "Metals"],
            ["Aluminium", "Metals"],
            ["Timber", "Bio-based"],
            ["Chemical product", "Chemicals"],
        ],
        columns=["material", "group"],
    )

    # One row per indicator and material. Leave a row out, or a cell blank,
    # where no reliable figure exists. The script lists it as a data gap.
    material_factors = pd.DataFrame(
        [
            ["Climate", "Cement", 0.60, 0.6],
            ["Water", "Cement", 0.60, 120],
            ["Material footprint", "Cement", 0.50, 1.2],
            ["Biodiversity", "Timber", 0.60, 0.08],
            ["Toxicity", "Chemical product", 0.40, 0.008],
        ],
        columns=["indicator", "material", "share_within_group", "impact_per_tonne"],
    )

    demand = pd.DataFrame(
        [
            ["Baseline", "Cement", 6.0e6],
            ["Baseline", "Timber", 1.0e6],
            ["Baseline", "Chemical product", 0.5e6],
        ],
        columns=["scenario", "material", "demand_t_per_year"],
    )

    with pd.ExcelWriter(path) as xl:
        for name, df in [("parameters", parameters), ("indicators", indicators),
                         ("sector_shares", sector_shares), ("group_shares", group_shares),
                         ("materials", materials), ("material_factors", material_factors),
                         ("demand", demand)]:
            df.to_excel(xl, sheet_name=name, index=False)
    print(f"Template written to {path}")


# --------------------------------------------------------------------------
# Calculation
# --------------------------------------------------------------------------
def load(path=INPUTS):
    sheets = pd.read_excel(path, sheet_name=None)
    params = sheets["parameters"].set_index("parameter")["value"]
    return params, sheets


def capability_factor(params):
    override = params.get("capability_factor_override")
    if pd.notna(override):
        return float(override)
    # Never above 1, so a poorer-than-average country is not given extra space.
    return min(1.0, params["gdp_pc_world"] / params["gdp_pc_spain"])


def run(path=INPUTS, out=RESULTS, charts=True):
    params, s = load(path)
    pop_share = params["population_spain"] / params["population_world"]
    cap = capability_factor(params)

    ind = s["indicators"]
    sec = s["sector_shares"].set_index("indicator")
    grp = s["group_shares"]
    fac = s["material_factors"]
    mats = s["materials"]
    demand = s["demand"]

    # Every indicator x material pair, so that missing pairs show up as gaps.
    grid = ind.merge(mats, how="cross")
    grid = grid.merge(grp, on=["indicator", "group"], how="left")
    grid = grid.merge(fac, on=["indicator", "material"], how="left")

    rows = []
    for space in SPACES:
        for principle in PRINCIPLES:
            for r in grid.itertuples(index=False):
                # Step 1
                global_space = r.full_space - (r.occupied_space if space == "RSOS" else 0)
                # Step 2
                spain = global_space * pop_share
                if principle == "Ability to pay":
                    spain *= cap
                # Step 3
                col = "needs_share" if principle == "Human needs" else "building_impact_share"
                b_share = sec.at[r.indicator, col] if r.indicator in sec.index else float("nan")
                building = spain * b_share
                local = sec.at[r.indicator, "local_capacity_buildings"] if r.indicator in sec.index else float("nan")
                if pd.notna(local):
                    building = min(building, local)
                # Step 4
                mat_env = building * r.group_share * r.share_within_group
                # Step 5
                cumulative_t = mat_env / r.impact_per_tonne
                annual_t = cumulative_t / r.budget_period_years
                rows.append(dict(
                    space=space, principle=principle, indicator=r.indicator, unit=r.unit,
                    group=r.group, material=r.material,
                    global_space=global_space, spain_share=pop_share,
                    capability_factor=cap if principle == "Ability to pay" else 1.0,
                    spain_budget=spain, building_share=b_share, building_budget=building,
                    group_share=r.group_share, share_within_group=r.share_within_group,
                    material_env_budget=mat_env, impact_per_tonne=r.impact_per_tonne,
                    budget_period_years=r.budget_period_years,
                    cumulative_budget_t=cumulative_t, annual_budget_t=annual_t,
                    annual_budget_kg_per_capita=annual_t * 1000 / params["population_spain"],
                ))
    chain = pd.DataFrame(rows)
    solved = chain.dropna(subset=["annual_budget_t"])

    # Step 5b: binding budget is the lowest tonnage across indicators.
    keys = ["space", "principle", "material"]
    idx = solved.groupby(keys)["annual_budget_t"].idxmin()
    binding = solved.loc[idx, keys + ["group", "indicator", "annual_budget_t",
                                      "cumulative_budget_t", "annual_budget_kg_per_capita"]]
    binding = binding.rename(columns={"indicator": "binding_indicator",
                                      "annual_budget_t": "binding_annual_budget_t",
                                      "cumulative_budget_t": "binding_cumulative_budget_t"})
    counts = solved.groupby(keys)["indicator"].nunique().rename("indicators_assessed")
    binding = binding.merge(counts, on=keys).reset_index(drop=True)

    # Indicator x material matrix of annual budgets, one block per case.
    matrix = solved.pivot_table(index=["space", "principle", "material"],
                                columns="indicator", values="annual_budget_t").reset_index()

    # Step 6: overshoot and transformation gap for each demand scenario.
    over = binding.merge(demand, on="material", how="inner")
    over["overshoot_ratio"] = over["demand_t_per_year"] / over["binding_annual_budget_t"]
    over["reduction_needed"] = (1 - 1 / over["overshoot_ratio"]).clip(lower=0)
    over["transformation_gap_t"] = (over["demand_t_per_year"] - over["binding_annual_budget_t"]).clip(lower=0)
    over["status"] = over["overshoot_ratio"].apply(lambda x: "Overshoot" if x > 1 else "Within budget")
    over = over[["space", "principle", "scenario", "material", "binding_indicator",
                 "binding_annual_budget_t", "demand_t_per_year", "overshoot_ratio",
                 "reduction_needed", "transformation_gap_t", "status"]]

    # Allocation corridor: binding annual budget under each principle.
    corridor = binding.pivot_table(index=["space", "material"], columns="principle",
                                   values="binding_annual_budget_t")
    corridor = corridor.reindex(columns=PRINCIPLES)
    corridor["corridor_low_t"] = corridor[PRINCIPLES].min(axis=1)
    corridor["corridor_high_t"] = corridor[PRINCIPLES].max(axis=1)
    corridor = corridor.reset_index()

    # Checks: shares must sum to 100% per indicator, and gaps are listed.
    g = grp.groupby("indicator")["group_share"].sum().rename("sum_of_group_shares")
    m = (fac.merge(mats, on="material").groupby(["indicator", "group"])["share_within_group"]
         .sum().rename("sum_of_shares_within_group").reset_index())
    share_check = m.merge(g, on="indicator", how="outer")
    share_check["unallocated_group_share"] = 1 - share_check["sum_of_group_shares"]
    share_check["unallocated_within_group"] = 1 - share_check["sum_of_shares_within_group"]

    base = chain[(chain.space == "RSOS")]
    gaps = base[base["annual_budget_t"].isna()][
        ["principle", "indicator", "material", "building_share", "group_share",
         "share_within_group", "impact_per_tonne"]].copy()
    gaps["missing"] = gaps.apply(
        lambda x: ", ".join(c for c in ["building_share", "group_share",
                                        "share_within_group", "impact_per_tonne"] if pd.isna(x[c])), axis=1)
    gaps = gaps[["principle", "indicator", "material", "missing"]]

    with pd.ExcelWriter(out) as xl:
        binding.to_excel(xl, sheet_name="binding_budgets", index=False)
        over.to_excel(xl, sheet_name="overshoot", index=False)
        corridor.to_excel(xl, sheet_name="corridor", index=False)
        matrix.to_excel(xl, sheet_name="budget_by_indicator", index=False)
        chain.to_excel(xl, sheet_name="calculation_chain", index=False)
        share_check.to_excel(xl, sheet_name="share_checks", index=False)
        gaps.to_excel(xl, sheet_name="data_gaps", index=False)

    if charts:
        make_charts(solved, binding, over, demand)

    pd.set_option("display.width", 200)
    pd.set_option("display.float_format", lambda v: f"{v:,.3f}")
    show = over[over.space == "RSOS"].copy()
    for c in ["binding_annual_budget_t", "demand_t_per_year", "transformation_gap_t"]:
        show[c] = show[c] / 1e6
    print(f"Population share {pop_share:.4%}   capability factor {cap:.2f}")
    print("\nBinding budgets and overshoot, RSOS, in Mt per year")
    print(show.drop(columns=["space", "scenario"]).to_string(index=False))
    print(f"\n{len(gaps)} data gaps listed in sheet data_gaps")
    print(f"Results written to {out}")
    return chain, binding, over


# --------------------------------------------------------------------------
# Charts
# --------------------------------------------------------------------------
BLUE, ORANGE, AQUA, RED, GREY, INK = "#2a78d6", "#eb6834", "#1baf7a", "#e34948", "#b9b8b2", "#0b0b0b"
P_COL = dict(zip(PRINCIPLES, [BLUE, ORANGE, AQUA]))


def _style(ax):
    for side in ["top", "right", "left"]:
        ax.spines[side].set_visible(False)
    ax.tick_params(length=0, labelsize=9)
    ax.grid(axis="x", color="#e6e5e0", linewidth=0.8)
    ax.set_axisbelow(True)


def make_charts(solved, binding, over, demand, space="RSOS", folder=CHARTS):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    folder.mkdir(exist_ok=True)
    b = binding[binding.space == space]
    d = demand[demand.scenario == demand.scenario.iloc[0]].set_index("material")["demand_t_per_year"]
    mats = sorted(b.material.unique())

    # 1  Allocation corridor: binding budget under each principle, one panel per material
    fig, axes = plt.subplots(len(mats), 1, figsize=(8, 1.9 * len(mats) + 0.6), squeeze=False)
    for ax, m in zip(axes[:, 0], mats):
        sub = b[b.material == m].set_index("principle").reindex(PRINCIPLES).dropna(subset=["binding_annual_budget_t"])
        vals = sub["binding_annual_budget_t"] / 1e6
        ax.barh(sub.index, vals, color=[P_COL[p] for p in sub.index], height=0.55)
        for p, v, ind in zip(sub.index, vals, sub["binding_indicator"]):
            ax.text(v, p, f"  {v:,.2f} Mt  ({ind})", va="center", fontsize=9, color=INK)
        top = vals.max()
        if m in d.index:
            ax.axvline(d[m] / 1e6, color=INK, linestyle="--", linewidth=1)
            ax.text(d[m] / 1e6, -0.42, f" demand {d[m] / 1e6:,.2f} Mt", ha="left", va="top", fontsize=8.5, color=INK)
            top = max(top, d[m] / 1e6)
        ax.set_xlim(0, top * 1.45)
        ax.invert_yaxis()
        ax.set_title(m, loc="left", fontsize=11, fontweight="bold")
        _style(ax)
    axes[-1, 0].set_xlabel("Binding fair budget, Mt per year")
    fig.suptitle(f"Allocation corridor by material ({space})", x=0.01, ha="left", fontsize=13)
    fig.tight_layout()
    fig.savefig(folder / "1_corridor.png", dpi=160)
    plt.close(fig)

    # 2  Budget allowed by each indicator, equal per capita. The lowest bar is binding.
    s = solved[(solved.space == space) & (solved.principle == PRINCIPLES[0])]
    fig, axes = plt.subplots(len(mats), 1, figsize=(8, 1.7 * len(mats) + 0.6), squeeze=False)
    for ax, m in zip(axes[:, 0], mats):
        sub = s[s.material == m].sort_values("annual_budget_t")
        vals = sub["annual_budget_t"] / 1e6
        cols = [RED] + [GREY] * (len(sub) - 1)
        ax.barh(sub["indicator"], vals, color=cols, height=0.55)
        for i, (ind, v) in enumerate(zip(sub["indicator"], vals)):
            ax.text(v, ind, f"  {v:,.2f} Mt" + ("  binding" if i == 0 else ""), va="center", fontsize=9, color=INK)
        ax.set_xlim(0, vals.max() * 1.35)
        ax.invert_yaxis()
        ax.set_title(m, loc="left", fontsize=11, fontweight="bold")
        _style(ax)
    axes[-1, 0].set_xlabel("Tonnes allowed by each indicator, Mt per year")
    fig.suptitle(f"Budget by indicator, equal per capita ({space})", x=0.01, ha="left", fontsize=13)
    fig.tight_layout()
    fig.savefig(folder / "2_budget_by_indicator.png", dpi=160)
    plt.close(fig)

    # 3  Overshoot ratio, log scale, with the limit at 1
    o = over[over.space == space]
    if len(o):
        omats = sorted(o.material.unique())
        fig, ax = plt.subplots(figsize=(8, 0.8 * len(omats) + 1.6))
        for p in PRINCIPLES:
            sub = o[o.principle == p]
            ax.scatter(sub["overshoot_ratio"], sub["material"], s=70, color=P_COL[p], label=p,
                       edgecolor="white", linewidth=1.5, zorder=3)
        ax.axvline(1, color=INK, linewidth=1)
        ax.text(1, len(omats) - 0.45, " fair budget = 1", fontsize=8.5, color=INK)
        ax.set_xscale("log")
        ax.set_ylim(-0.6, len(omats) - 0.3)
        ax.set_xlabel("Overshoot ratio = demand / binding budget (log scale)")
        ax.legend(frameon=False, fontsize=9, ncol=3, loc="lower center", bbox_to_anchor=(0.5, 1.0))
        _style(ax)
        fig.suptitle(f"Overshoot ratio by material ({space})", x=0.01, ha="left", fontsize=13)
        fig.tight_layout()
        fig.savefig(folder / "3_overshoot.png", dpi=160)
        plt.close(fig)
    print(f"Charts written to {folder}/")


# --------------------------------------------------------------------------
# Self-test against the worked examples in Chapter 10 and the Ramboll slides
# --------------------------------------------------------------------------
def self_test():
    import tempfile
    tmp = Path(tempfile.mkdtemp())
    write_template(tmp / "in.xlsx")
    chain, binding, over = run(tmp / "in.xlsx", tmp / "out.xlsx", charts=False)
    c = chain[chain.space == "RSOS"].set_index(["principle", "indicator", "material"])["annual_budget_t"] / 1e6
    # (principle, indicator, material, published value in Mt, where it is published)
    cases = [
        ("Equal per capita", "Climate", "Cement", 4.2, "Ch10 10.9.1 and 10.21"),
        ("Ability to pay", "Climate", "Cement", 3.4, "Ch10 10.9.2"),
        ("Human needs", "Climate", "Cement", 5.6, "Ch10 10.9.3"),
        ("Equal per capita", "Water", "Cement", 0.34, "Ramboll slide, water"),
        ("Equal per capita", "Material footprint", "Cement", 5.0, "Ramboll slide, footprint"),
        ("Equal per capita", "Biodiversity", "Timber", 0.68, "Ramboll slide, biodiversity"),
        ("Equal per capita", "Toxicity", "Chemical product", 0.36, "Ramboll slide, toxicity"),
    ]
    print("\nSelf-test. Tolerance 2% because the brief rounds Spain's share to 0.56%.")
    ok = True
    for p, i, m, want, src in cases:
        got = c[(p, i, m)]
        passed = abs(got - want) / want < 0.02
        ok &= passed
        print(f"  {'PASS' if passed else 'FAIL'}  {p:17s} {i:19s} {m:17s} got {got:6.3f}  published {want:5.2f}  {src}")
    # Structural checks that hold for any dataset
    b = binding[binding.space == "RSOS"].set_index(["principle", "material"])["binding_annual_budget_t"]
    low = chain[chain.space == "RSOS"].groupby(["principle", "material"])["annual_budget_t"].min().dropna()
    checks = {
        "binding budget equals the lowest indicator budget": bool((b.sort_index() == low.sort_index()).all()),
        "ability to pay never exceeds equal per capita": bool(
            (c["Ability to pay"].dropna() <= c["Equal per capita"].dropna() + 1e-9).all()),
        "overshoot ratio equals demand / budget": bool(
            ((over.demand_t_per_year / over.binding_annual_budget_t - over.overshoot_ratio).abs() < 1e-9).all()),
    }
    for name, passed in checks.items():
        ok &= passed
        print(f"  {'PASS' if passed else 'FAIL'}  {name}")
    print("\nAll checks passed." if ok else "\nSome checks FAILED.")
    return ok


if __name__ == "__main__":
    if "--test" in sys.argv:
        sys.exit(0 if self_test() else 1)
    if "--template" in sys.argv or not INPUTS.exists():
        write_template()
    if "--template" not in sys.argv:
        run()
