<div align="center">

# UHTM Dataset — IMI Project

### Ultra-High Temperature Materials · Physics-Informed Feature Engineering

[![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![pandas](https://img.shields.io/badge/pandas-2.x-150458?style=flat-square&logo=pandas&logoColor=white)](https://pandas.pydata.org)
[![openpyxl](https://img.shields.io/badge/openpyxl-3.x-217346?style=flat-square)](https://openpyxl.readthedocs.io)
[![Samples](https://img.shields.io/badge/Samples-200-0891b2?style=flat-square)]()
[![Features](https://img.shields.io/badge/Features-48-7c3aed?style=flat-square)]()
[![Targets](https://img.shields.io/badge/Targets-3-dc2626?style=flat-square)]()
[![Seed](https://img.shields.io/badge/Random%20Seed-42-f59e0b?style=flat-square)]()

<br/>

> A collaboratively engineered **200-sample × 48-feature** dataset of refractory ceramics and composites,  
> built by a 4-member team using physics-informed feature engineering across 8 material property domains.

</div>

---
## 📋 Website : https://imi-project.vercel.app/

## 📋 Table of Contents

- [Project Overview](#-project-overview)
- [Dataset Summary](#-dataset-summary)
- [Repository Structure](#-repository-structure)
- [Data Pipeline](#-data-pipeline)
- [Team Contributions](#-team-contributions)
  - [Aadi — Member 1](#-aadi--member-1-thermodynamic--electronic)
  - [Krish — Member 2](#-krish--member-2-mechanical--thermal-transport)
  - [Salan — Member 3](#-salan--member-3-oxidation--microstructural)
  - [Niranjan — Member 4](#-niranjan--member-4-phasecomposite--ml-descriptors--targets)
- [Feature Groups at a Glance](#-feature-groups-at-a-glance)
- [Material Classes](#-material-classes)
- [Target Variables](#-target-variables)
- [Setup & Usage](#-setup--usage)
- [Output Files](#-output-files)

---

## 🧪 Project Overview

This project constructs a structured, ML-ready dataset for **Ultra-High Temperature Materials (UHTMs)** — refractory ceramics used in hypersonic re-entry vehicles, thermal protection systems (TPS), and aerospace leading-edge components.

The dataset covers **10 material families** (carbides, borides, nitrides) and their composites/doped variants across **8 physical property domains**, with 3 supervised regression targets. All features are derived from 5 literature-anchored material properties using validated physical relations plus realistic Gaussian noise (`np.random.seed(42)`).

**Key design principles:**
- Each of the 4 team members owns a distinct physical domain → clean separation of concerns
- All features trace back to 5 shared anchors (`Tₘ`, `ρ`, `vₑ`, `ΔEN`, `Pₛ`) → internally consistent dataset
- Reproducible: fixed seed, shared base file, merge validation enforced programmatically

---

## 📊 Dataset Summary

| Property | Value |
|---|---|
| Total Samples | 200 (100 experimental + 100 synthetic) |
| Feature Columns | 48 (F01–F48) |
| Target Columns | 3 (T1–T3) |
| Metadata Columns | 5 (Sample_ID, Material_System, Source_Type, Synthesis_Method, Crystal_Structure) |
| Total Columns | 56 |
| Material Families | 10 base + composites + doped variants |
| Random Seed | 42 (all scripts) |
| Final Output | `UHTM_final_200x48.xlsx` + `UHTM_final_200x48.csv` |

---

## 🗂 Repository Structure

```
IMI-Project-main/
│
├── Aadi-Dev/                          # Member 1 — Thermodynamic + Electronic
│   ├── dataset.py                     # ★ Generates F01–F12
│   ├── Aadi.xlsx                      # Output: 200 × 17 (meta + 12 features)
│   └── UHTM_base_200.xlsx             # Base reference copy
│
├── Krish/                             # Member 2 — Mechanical + Thermal + Infrastructure
│   ├── Intro.py                       # Branch onboarding note
│   └── LAB EVALUATION/
│       ├── Krishh.py                  # ★ Generates F13–F24
│       ├── Krishh.xlsx                # Output: 200 × 17
│       ├── UHTM_base_200.xlsx
│       └── Merge/
│           ├── Base.py                # ★★ Generates UHTM_base_200.xlsx (run first)
│           ├── mergeAll.py            # ★★ Final merge of all 4 member files
│           ├── AadiDev.xlsx           # Member 1 snapshot for merge
│           ├── Krishh.xlsx            # Member 2 snapshot for merge
│           ├── Niranjan.xlsx          # Member 4 snapshot for merge
│           ├── Salan.xlsx             # Member 3 snapshot for merge
│           └── UHTM_final_200x48.xlsx # ★★ FINAL MERGED DATASET
│
├── Niranjan/                          # Member 4 — Phase + ML Descriptors + Targets
│   ├── Niranjan.py                    # ★ Generates F37–F48 + T1, T2, T3
│   ├── Niranjan.xlsx                  # Output: 200 × 20
│   └── UHTM_base_200.xlsx
│
├── Salan/                             # Member 3 — Oxidation + Microstructural
│   ├── lab evaluation/
│   │   ├── Salan.py                   # ★ Generates F25–F36
│   │   ├── Salan_member3.xlsx         # Output: 200 × 17
│   │   └── UHTM_base_200.xlsx
│   └── Backup_datasets/
│       ├── UHTM_Complete.csv
│       ├── UHTM_Complete.xlsx
│       └── completefile.py
│
├── UHTM_final_200x48.csv              # ★★ ML-ready CSV (root copy)
└── UHTM_final_200x48.xlsx             # ★★ Final annotated Excel (root copy)
```

---

## 🔄 Data Pipeline

```
┌─────────────────────────────────────────────────────────────────────────┐
│                                                                         │
│   STEP 1: Base.py                                                       │
│   ─────────────                                                         │
│   Krish generates UHTM_base_200.xlsx                                    │
│   200 rows × 10 cols (5 meta + 5 hidden anchors: Tₘ, ρ, vₑ, ΔEN, Pₛ)  │
│                          │                                              │
│          ┌───────────────┼────────────────┐                             │
│          ▼               ▼                ▼              ▼              │
│   STEP 2a: Aadi   2b: Krish       2c: Salan      2d: Niranjan          │
│   dataset.py      Krishh.py       Salan.py        Niranjan.py          │
│   F01–F12         F13–F24         F25–F36         F37–F48 + T1–T3      │
│   Aadi.xlsx       Krishh.xlsx     Salan_m3.xlsx   Niranjan.xlsx        │
│          │               │                │              │              │
│          └───────────────┴────────────────┴──────────────┘             │
│                                   │                                     │
│                                   ▼                                     │
│   STEP 3: mergeAll.py                                                   │
│   ────────────────────                                                  │
│   Validates 200 rows + Sample_ID match across all 4 files               │
│   Horizontally joins all feature groups                                 │
│                                   │                                     │
│                                   ▼                                     │
│   STEP 4: UHTM_final_200x48.xlsx / .csv                                 │
│   200 rows × 56 cols (5 meta + 48 features + 3 targets)                 │
│   + Summary Stats sheet + Feature Legend sheet                          │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## 👥 Team Contributions

---

### 🔵 Aadi — Member 1: Thermodynamic & Electronic

**Script:** `Aadi-Dev/dataset.py` &nbsp;|&nbsp; **Output:** `Aadi.xlsx` &nbsp;|&nbsp; **Features:** `F01–F12`

Aadi is responsible for the foundational material properties spanning thermodynamic stability and quantum electronic structure — the two groups that most directly determine a material's suitability as a UHTM candidate.

#### Group A — Thermodynamic / Structural (F01–F06)

| Feature | Name | Formula / Physics | Unit |
|---|---|---|---|
| **F01** | Melting Point | Literature anchor `Tₘ` + 2% noise | K |
| **F02** | Debye Temperature | `θ_D ≈ 300 + 0.15·Tₘ` (Debye-Grüneisen scaling) | K |
| **F03** | Cohesive Energy | `E_coh = vₑ·0.82 + ΔEN·1.1` (metallic + ionic/covalent) | eV/atom |
| **F04** | Formation Enthalpy | `ΔHf = -(ΔEN·45 + vₑ·8)` (exothermic = stable) | kJ/mol |
| **F05** | Lattice Parameter a | `a ≈ 3.2 + ρ^(-0.3)·0.5` (inverse power law with density) | Å |
| **F06** | Grüneisen Parameter | `γ ≈ 0.4 + ΔEN·0.3 + vₑ·0.05` (anharmonicity index) | — |

#### Group B — Electronic Structure (F07–F12)

| Feature | Name | Formula / Physics | Unit |
|---|---|---|---|
| **F07** | Band Gap | `max(0, 0.5 - vₑ·0.05)` — metallic carbides/nitrides ≈ 0 | eV |
| **F08** | DOS at Fermi Level | `N(Ef) ≈ vₑ·0.55 + 1.2` (more valence e⁻ → higher DOS) | states/eV |
| **F09** | Bader Charge Transfer | `Δq ≈ ΔEN·0.6` (Pauling-type charge transfer) | e⁻ |
| **F10** | Fermi Velocity | `v_F = 1×10⁶ · (vₑ/8.5)` (free-electron model) | m/s |
| **F11** | Valence Electron Density | `n = ρ·vₑ·1.8×10²²` (electrons per unit volume) | ×10²²/cm³ |
| **F12** | Work Function | `φ ≈ 3.5 + ΔEN·0.8 - vₑ·0.05` (surface escape energy) | eV |

> **Physical significance:** F01 is the primary UHTM selection criterion. F07 distinguishes metallic from semiconducting behaviour at high T. F11 governs metallic bonding strength and shear modulus. F12 controls thermionic emission.

---

### 🟣 Krish — Member 2: Mechanical & Thermal Transport

**Scripts:** `Krishh.py`, `Base.py`, `mergeAll.py` &nbsp;|&nbsp; **Output:** `Krishh.xlsx` &nbsp;|&nbsp; **Features:** `F13–F24`

Krish owns both mechanical integrity and thermal transport features — the two most critical property groups for structural aerospace applications. Krish also **authored the base dataset generator (`Base.py`)** and the **final merge script (`mergeAll.py`)**, serving as the project's data infrastructure lead.

#### Group C — Mechanical Properties (F13–F18)

| Feature | Name | Formula / Physics | Unit |
|---|---|---|---|
| **F13** | Young's Modulus | `E ≈ 200 + Tₘ·0.04 + vₑ·15` (Gilman-Cohen stiffness relation) | GPa |
| **F14** | Vickers Hardness | `H_v ≈ 15 + ΔEN·8 + vₑ·1.2` (bond character + electron count) | GPa |
| **F15** | Fracture Toughness K_Ic | `K_Ic ≈ 2.5 + 1.5/ΔEN` (ionic bonds = more brittle) | MPa√m |
| **F16** | Compressive Strength | `σ_c ≈ 0.6·E` (empirical ratio for dense ceramics) | GPa |
| **F17** | Poisson's Ratio | `ν ≈ 0.18 + ΔEN·0.02` (covalent ≈ 0.18; ionic → higher) | — |
| **F18** | Flexural Strength | `σ_f = 300 + E·0.8 + H·12 - porosity·15` | MPa |

#### Group D — Thermal Transport (F19–F24)

| Feature | Name | Formula / Physics | Unit |
|---|---|---|---|
| **F19** | Thermal Conductivity | `κ = 20 + vₑ·4 - ΔEN·3` (electronic + ionic scattering) | W/m·K |
| **F20** | Coeff. Thermal Expansion | `CTE = 6.5 + ΔEN·0.8 - vₑ·0.3` | ×10⁻⁶/K |
| **F21** | Specific Heat Capacity | `Cp ≈ 180 + 800/ρ` (inverse density, Dulong-Petit) | J/kg·K |
| **F22** | Thermal Diffusivity | `α = κ / (ρ·Cp)` (exact thermodynamic definition) | m²/s |
| **F23** | Max Service Temperature | `T_max ≈ 0.75·Tₘ` (standard refractory engineering rule) | K |
| **F24** | Thermal Shock Resistance | `R = σ_f·κ / (E·CTE)` (Hasselman R-parameter) | W/m |

> **Infrastructure contributions:** `Base.py` defines all 10 material anchors from literature (Materials Project, JARVIS-DFT, Fahrenholtz & Hilmas 2012). `mergeAll.py` validates row counts and Sample_ID alignment before joining, preventing silent merge errors.

---

### 🔴 Salan — Member 3: Oxidation Stability & Microstructural

**Script:** `Salan/lab evaluation/Salan.py` &nbsp;|&nbsp; **Output:** `Salan_member3.xlsx` &nbsp;|&nbsp; **Features:** `F25–F36`

Salan handles the chemical stability and process-microstructure features — the groups that determine how a UHTM behaves over time in oxidising environments and how its properties are influenced by the synthesis route.

#### Group E — Oxidation / Chemical Stability (F25–F30)

| Feature | Name | Formula / Physics | Unit |
|---|---|---|---|
| **F25** | Oxidation Onset Temperature | `T_ox = 800 + ΔEN·150 + Tₘ·0.05` | K |
| **F26** | Parabolic Rate Constant k_p | `k_p = 1×10⁻¹² · exp(ΔEN·0.8)` (Arrhenius) | kg²/m⁴s |
| **F27** | Oxidation Activation Energy | `Ea = 120 + ΔEN·30` (diffusion barrier) | kJ/mol |
| **F28** | Gravimetric Parabolic Rate | `k_p_grav = 1×10⁻¹⁰ · ΔEN` (TGA standard unit) | g²/cm⁴s |
| **F29** | Oxide Layer Stability Index | `OLS = ΔEN·0.7 + vₑ·0.15` (protective scale adherence) | — |
| **F30** | Oxygen Diffusivity in Oxide | `D_O = 1×10⁻¹⁴ · exp(-ΔEN·1.2)` | m²/s |

#### Group F — Microstructural Features (F31–F36)

| Feature | Name | Formula / Physics | Unit |
|---|---|---|---|
| **F31** | Average Grain Size | `d = 2 + (Pₛ^-0.3)·10` (Hall-Petch: smaller grains → harder) | μm |
| **F32** | Relative Density | `ρ_rel = min(99.9, 92 + Pₛ·0.18)` | % |
| **F33** | Porosity | `P = max(0.1, 100 - ρ_rel)` (complement of relative density) | % |
| **F34** | Crystallite Size (XRD) | `D_xrd = grain_size_μm · 1000 · 0.08` (Scherrer ~8% of grain) | nm |
| **F35** | Dislocation Density | `ρ_disl = 1×10¹²/d^1.5` (inverse power law with grain size) | ×10¹²/m² |
| **F36** | Grain Boundary Energy | `γ_gb = 0.3 + ΔEN·0.25 + vₑ·0.02` | J/m² |

> **Physical significance:** F25 is the primary chemical stability criterion. F26 determines oxide scale growth rate — smaller k_p = better protection. F31–F33 directly link sintering pressure (Pₛ anchor) to microstructure, closing the process–property loop.

---

### 🟢 Niranjan — Member 4: Phase/Composite · ML Descriptors · Targets

**Script:** `Niranjan/Niranjan.py` &nbsp;|&nbsp; **Output:** `Niranjan.xlsx` &nbsp;|&nbsp; **Features:** `F37–F48` + **Targets:** `T1–T3`

Niranjan handles the highest-level features: composite system descriptors, dimensionless ML merit indices designed for Pareto optimisation, and all three supervised learning targets. This is the most interdependent feature block — many features in Groups G and H re-derive quantities from Groups A–F using the same seed to ensure cross-member consistency.

#### Group G — Phase / Composite Features (F37–F42)

| Feature | Name | Formula / Physics | Unit |
|---|---|---|---|
| **F37** | Phase Stability Index | `PSI = Tₘ/4000 + E_coh/12` (normalised thermodynamic stability) | — |
| **F38** | Secondary Phase Vol. Fraction | 0% monolithic; 5–20% composites (index-dependent) | % |
| **F39** | Interfacial Energy | `γ_int = 0.5 + ΔEN·0.3` (ionic mismatch → delamination risk) | J/m² |
| **F40** | CTE Mismatch Index | `ΔCTE = |CTE - 5.0|·0.3` (relative to SiC CTE ~5×10⁻⁶/K) | — |
| **F41** | Solid Solution Distortion δ | `δ = ΔEN·0.12 + (vₑ%3)·0.05` (Hume-Rothery lattice distortion) | — |
| **F42** | Wettability Index | `W = 1 - f_ion = vₑ·0.15/(ΔEN + vₑ·0.15)` (covalent fraction) | — |

#### Group H — ML Descriptor Features (F43–F48)

| Feature | Name | Formula / Physics | Unit |
|---|---|---|---|
| **F43** | Thermal Merit Index | `TMI = κ/(CTE·ρ)` (Ashby figure of merit for TPS panels) | W/kg |
| **F44** | Toughness-Stiffness Index | `TSI = K_Ic·√E` (combined crack resistance and stiffness) | GPa·√GPa |
| **F45** | Oxidation Merit Score | `OMS = T_ox/(k_p_norm + 1)` (Bayesian optimisation reward signal) | — |
| **F46** | Bond Ionicity Fraction | `f_ion = ΔEN/(ΔEN + vₑ·0.15)` (Pauling ionicity) | — |
| **F47** | Structural Stability Index | `SSI = (Tₘ/4000)·(E_coh/10)·(1 - P/100)` | — |
| **F48** | Creep Resistance Parameter | `CR = (Tₘ/3000)·(E/400)·(1/d)^0.3` (grain size + stiffness) | — |

---

## 🎯 Target Variables

All three targets are defined and computed by **Niranjan (Member 4)**.

| Target | Name | Type | Description |
|---|---|---|---|
| **T1** | Flexural Strength | Continuous (MPa) | Multi-factor structural target. Driven by Young's Modulus, Vickers Hardness, and Porosity. Primary design metric for load-bearing structures. |
| **T2** | Oxidation Resistance Score | Continuous (0–10) | Composite weighted score: onset temperature (×3) + rate constant (×2) + stability index (×2) + CTE match (×1). |
| **T3** | Thermal Shock Cycles | Integer (cycles) | Predicted cycles-to-failure under rapid thermal cycling. Driven by K_Ic, CTE, and composite mismatch index F40. |

---

## 🗃 Feature Groups at a Glance

| Group | Features | Member | Domain | Colour (in Excel) |
|---|---|---|---|---|
| A — Thermodynamic | F01–F06 | Aadi | Phase stability, bonding energetics | Teal |
| B — Electronic | F07–F12 | Aadi | DFT / quantum properties | Blue |
| C — Mechanical | F13–F18 | Krish | Structural integrity | Purple |
| D — Thermal Transport | F19–F24 | Krish | Heat transport & diffusion | Orange |
| E — Oxidation | F25–F30 | Salan | Chemical stability | Red |
| F — Microstructural | F31–F36 | Salan | Process–structure–property | Green |
| G — Phase/Composite | F37–F42 | Niranjan | Multi-phase composite systems | Indigo |
| H — ML Descriptors | F43–F48 | Niranjan | Pareto / reward signals for inverse design | Teal-dark |
| Targets | T1–T3 | Niranjan | Supervised regression targets | Dark Red |

---

## 🧱 Material Classes

Base anchor values sourced from: *Materials Project (mp-\*)*, *JARVIS-DFT*, Fahrenholtz & Hilmas (2012), Cedillos-Barraza et al. (2016), Opeka et al. J. Eur. Ceram. Soc.

| Material | Crystal | T_m (K) | ρ (g/cm³) | v_e | ΔEN |
|---|---|---|---|---|---|
| HfC | FCC | 3900 | 12.20 | 8 | 1.3 |
| ZrC | FCC | 3420 | 6.73 | 8 | 1.3 |
| TaC | FCC | 3880 | 14.30 | 9 | 1.1 |
| HfB₂ | HEX | 3380 | 10.50 | 6 | 0.9 |
| ZrB₂ | HEX | 3245 | 6.09 | 6 | 0.9 |
| TiC | FCC | 3160 | 4.93 | 8 | 1.5 |
| NbC | FCC | 3600 | 7.79 | 9 | 1.2 |
| HfN | FCC | 3385 | 13.80 | 9 | 1.6 |
| ZrN | FCC | 2980 | 7.09 | 9 | 1.6 |
| TaN | HEX | 3090 | 16.30 | 10 | 1.4 |

**Composite variants** (experimental, index 50–99): `HfC-SiC`, `ZrB₂-SiC`, `HfB₂-SiC`, `TaC-HfC`, `ZrC-TiC`, `HfC-TaC`, `ZrB₂-ZrC`, `HfB₂-MoSi₂`, `TiC-TiB₂`, `NbC-HfC`

**Doped variants** (synthetic, index 150–199): `HfC:Y`, `ZrC:La`, `TaC:W`, `HfB₂:Al`, `ZrB₂:Y`, `TiC:Nb`, `NbC:Ta`, `HfN:Zr`, `ZrN:Hf`, `TaN:Nb`

---

## ⚙️ Setup & Usage

### Prerequisites

```bash
pip install pandas numpy openpyxl
```

### Step 1 — Generate the base file (run once)

```bash
# From Krish/LAB EVALUATION/Merge/
python Base.py
# Output: UHTM_base_200.xlsx
# Contains: 200 rows × 10 cols (5 metadata + 5 physical anchors)
```

### Step 2 — Each member generates their features independently

```bash
# Member 1 — Aadi
cd Aadi-Dev/
python dataset.py
# Output: Aadi.xlsx (200 × 17)

# Member 2 — Krish
cd "Krish/LAB EVALUATION/"
python Krishh.py
# Output: Krishh.xlsx (200 × 17)

# Member 3 — Salan
cd "Salan/lab evaluation/"
python Salan.py
# Output: Salan_member3.xlsx (200 × 17)

# Member 4 — Niranjan
cd Niranjan/
python Niranjan.py
# Output: Niranjan.xlsx (200 × 20: 12 features + 3 targets + 5 meta)
```

### Step 3 — Merge all outputs

```bash
# Copy all member xlsx files into Krish/LAB EVALUATION/Merge/
# Rename if needed: AadiDev.xlsx, Krishh.xlsx, Salan.xlsx, Niranjan.xlsx

cd "Krish/LAB EVALUATION/Merge/"
python mergeAll.py
# Validates: 200 rows + Sample_ID match across all 4 files
# Output: UHTM_final_200x48.xlsx
#   Sheet 1 — UHTM_Full_Dataset  (200 × 56, colour-coded by group)
#   Sheet 2 — Summary Stats       (describe() for all F and T columns)
#   Sheet 3 — Feature Legend       (group → member → domain mapping)
```

### Reproducibility

All scripts use `np.random.seed(42)` at the top level. As long as the base file is generated first and member scripts are run with the same seed, all outputs are deterministic.

---

## 📦 Output Files

| File | Location | Description |
|---|---|---|
| `UHTM_base_200.xlsx` | `Krish/LAB EVALUATION/Merge/` | Base dataset with material anchors. Input to all 4 member scripts. |
| `Aadi.xlsx` | `Aadi-Dev/` | Member 1 output: F01–F12 |
| `Krishh.xlsx` | `Krish/LAB EVALUATION/` | Member 2 output: F13–F24 |
| `Salan_member3.xlsx` | `Salan/lab evaluation/` | Member 3 output: F25–F36 |
| `Niranjan.xlsx` | `Niranjan/` | Member 4 output: F37–F48 + T1–T3 |
| `UHTM_final_200x48.xlsx` | Root + `Merge/` | ★ Final merged dataset with summary and legend sheets |
| `UHTM_final_200x48.csv` | Root | ★ ML-ready flat CSV export |

---

## 👤 Authors

| Member | Features | Scripts |
|---|---|---|
| **Aadi** | F01–F12 (Thermodynamic + Electronic) | `Aadi-Dev/dataset.py` |
| **Krish** | F13–F24 (Mechanical + Thermal) + Base + Merge | `Krishh.py`, `Base.py`, `mergeAll.py` |
| **Salan** | F25–F36 (Oxidation + Microstructural) | `Salan/lab evaluation/Salan.py` |
| **Niranjan** | F37–F48 (Phase + ML Descriptors) + T1–T3 | `Niranjan/Niranjan.py` |

---

<div align="center">

*IMI Project · 2026 · Ultra-High Temperature Materials Dataset*

</div>
