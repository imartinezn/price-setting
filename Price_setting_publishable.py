#!/usr/bin/env python
# coding: utf-8


# Code Price-setting tipping points in decarbonising EU electricity markets
# Author: I.M. (2026) with support from Large Language Models (LLM)


# %% Cell 0 - CONTROL FILE
# Cell 0 - CONTROL FILE LOADER (v3, harvested 11-sheet format)
# Reads control_file.xlsx and rebuilds every parameter structure the pipeline uses.
# The pipeline REFUSES to run without the control file - no silent defaults.
# 20260925: _num() applied to ALL numeric sheet reads (locale-proof: accepts
#           comma- or dot-decimal text and real numbers); duplicate hydro_params
#           block merged into one; CONTROL_FILE switchable via env var.
#
# Exposes:
#   CFG (SimpleNamespace)   scalars from 'config' + .paths dict + .zones DataFrame
#   ZONES_EU_ONLY, ALL_ZONES, ZONE_TO_COUNTRY
#   NEIGHBOURS, RESERVOIR_FILLING_PROXIES, ZONES_EXCLUDE_HYDRO_RESERVOIR
#   BASE_TO_CAT, CATEGORIES_9
#   THERMAL_PARAMS, GEN_MIN_THRESHOLDS, ZONE_THRESHOLDS, get_generation_threshold
#   HYDRO_RESERVOIR_PARAMS, PUMPED_STORAGE_PARAMS   (also on CFG)
#   GAS_HUB_MAP, GAS_TRANSPORT_ADDERS, IBERIAN_GAS_CAP
#   efficiency_maps            ({TECH}_EFFICIENCY_MAP dicts, zone -> {min,max,avg})
#   TECH_COLORS

import ast
import os
import re
from pathlib import Path
from types import SimpleNamespace
import pandas as pd
import gc

def _num(v):
    """Locale-proof numeric parse: accepts numbers and comma- or dot-decimal text.
    NOTE: control-file convention - plain numbers, no thousands separators."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    return float(str(v).strip().replace(',', '.'))



CONTROL_FILE = Path(os.environ.get('CONTROL_FILE', 'control_file.xlsx'))
if not CONTROL_FILE.exists():
    raise FileNotFoundError(f"Control file not found: {CONTROL_FILE.resolve()}")

_x = pd.ExcelFile(CONTROL_FILE)
_need = {'config', 'paths', 'zones', 'neighbours', 'categories', 'thermal_params',
         'gen_thresholds', 'hydro_params', 'gas', 'efficiencies', 'style'}
_missing = _need - set(_x.sheet_names)
if _missing:
    raise ValueError(f"Control file missing sheets: {sorted(_missing)}")

# ---- config ------------------------------------------------------------------------
_c = pd.read_excel(_x, 'config')
_cast = {'float': _num,                                   # <- locale-proof
         'int': lambda v: int(round(_num(v))),            # <- locale-proof
         'str': str,
         'bool': lambda v: str(v).strip().lower() in ('true', '1', 'yes'),
         'date': lambda v: pd.Timestamp(v),
         'list_int': lambda v: [int(i) for i in ast.literal_eval(str(v))]}
CFG = SimpleNamespace()
for _, r in _c.iterrows():
    setattr(CFG, str(r['key']).strip(), _cast[str(r['type']).strip()](r['value']))
for _k, _d in [('OUTPUT_HEADLINE_ONLY', False)]:
    if not hasattr(CFG, _k):
        setattr(CFG, _k, _d)

# ---- paths -------------------------------------------------------------------------
_p = pd.read_excel(_x, 'paths')
CFG.paths = {str(r['key']).strip(): Path(str(r['path']).strip()) for _, r in _p.iterrows()}
OUTPUT_DIR = CFG.paths.get('OUTPUT_DIR', Path('output'))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ---- base dir: project root for reference/
import sys
CFG.BASE_DIR = Path(CFG.paths.get('BASE_DIR', CONTROL_FILE.resolve().parent))
BASE_DIR = CFG.BASE_DIR
if str(CFG.BASE_DIR) not in sys.path:
    sys.path.insert(0, str(CFG.BASE_DIR))


# ---- zones -------------------------------------------------------------------------
CFG.zones = pd.read_excel(_x, 'zones')
CFG.zones['active_from'] = pd.to_datetime(CFG.zones['active_from'])
CFG.zones['active_to'] = pd.to_datetime(CFG.zones['active_to'], errors='coerce')
ALL_ZONES = CFG.zones['zone'].tolist()
ZONES_EU_ONLY = CFG.zones.loc[CFG.zones['eu_scope'] == True, 'zone'].tolist()
ZONES_EU_ALL = ALL_ZONES          # processing universe incl. non-EU neighbour zones
ZONE_TO_COUNTRY = dict(zip(CFG.zones['zone'], CFG.zones['country']))
RESERVOIR_FILLING_PROXIES = {r['zone']: r['reservoir_proxy']
                             for _, r in CFG.zones.iterrows()
                             if isinstance(r['reservoir_proxy'], str) and r['reservoir_proxy'].strip()}

_hr = CFG.zones.set_index('zone')['hydro_reservoir'].astype(str).str.upper() == 'TRUE'
HYDRO_RESERVOIR_ZONES         = _hr[_hr].index.tolist()
ZONES_EXCLUDE_HYDRO_RESERVOIR = _hr[~_hr].index.tolist()
_badp = [z for z in RESERVOIR_FILLING_PROXIES if z not in HYDRO_RESERVOIR_ZONES]
if _badp:
    print(f" reservoir_proxy set on zones marked hydro_reservoir=FALSE: {_badp}")


# ---- neighbours --------------------------------------------------------------------
_n = pd.read_excel(_x, 'neighbours')
NEIGHBOURS = {}
for _, r in _n.iterrows():
    NEIGHBOURS.setdefault(str(r['zone']).strip(), []).append(str(r['neighbour']).strip())
_unknown = sorted({z for z in list(NEIGHBOURS) + sum(NEIGHBOURS.values(), [])
                   if z not in ALL_ZONES})
if _unknown:
    print(f" neighbours references zones absent from zones sheet: {_unknown} "
          f"(candidates will never match unless data exists)")

# ---- categories --------------------------------------------------------------------
_cat = pd.read_excel(_x, 'categories')
BASE_TO_CAT = dict(zip(_cat['base_tech'].str.strip(), _cat['category'].str.strip()))
CATEGORIES_9 = ['Coal', 'Lignite', 'Gas', 'Other Fossil Fuels', 'Hydro Reservoir',
                'Hydro RoR', 'Nuclear', 'Biomass', 'VRE']
_bad = sorted(set(BASE_TO_CAT.values()) - set(CATEGORIES_9))
if _bad:
    raise ValueError(f"categories sheet maps to unknown categories: {_bad}")

# 20261021 added to reflect gen_tech_map now in control_file
# ---- gen_tech_map: raw generation column names -> base technologies ----------------
from collections import OrderedDict
_gm = pd.read_excel(_x, 'gen_tech_map')
ENTSOE_TO_STANDARD_MAPPING = OrderedDict(
    (str(r['entsoe_name']).strip(), str(r['base_tech']).strip())
    for _, r in _gm.iterrows()
    if pd.notna(r['entsoe_name']) and pd.notna(r['base_tech'])
    and str(r['entsoe_name']).strip() not in ('', 'HOW TO USE'))
if not ENTSOE_TO_STANDARD_MAPPING:
    raise ValueError("gen_tech_map sheet: no mapping rows found")
_nocat = sorted(set(ENTSOE_TO_STANDARD_MAPPING.values()) - set(BASE_TO_CAT))
if _nocat:
    print(f" gen_tech_map base techs without a category (kept in data, "
          f"never aggregated into the 9 categories): {_nocat}")

# ---- thermal params ----------------------------------------------------------------
_tp = pd.read_excel(_x, 'thermal_params')
THERMAL_PARAMS = {}
for _, r in _tp.iterrows():
    THERMAL_PARAMS.setdefault(str(r['fuel']).strip(), {})[str(r['parameter']).strip()] = _num(r['value'])

# ---- generation thresholds ---------------------------------------------------------
_g = pd.read_excel(_x, 'gen_thresholds')
GEN_MIN_THRESHOLDS = {str(r['tech']).strip(): _num(r['threshold_mw'])
                      for _, r in _g[_g['scope'] == 'default'].iterrows()}
ZONE_THRESHOLDS = {}
for _, r in _g[~_g['scope'].isin(['default', 'srmc_calc'])].iterrows():
    ZONE_THRESHOLDS.setdefault(str(r['scope']).strip(), {})[str(r['tech']).strip()] = _num(r['threshold_mw'])
if 'default' not in GEN_MIN_THRESHOLDS:
    GEN_MIN_THRESHOLDS['default'] = 10.0

def get_generation_threshold(tech_key, zone=None):
    if zone and zone in ZONE_THRESHOLDS and tech_key in ZONE_THRESHOLDS[zone]:
        return ZONE_THRESHOLDS[zone][tech_key]
    return GEN_MIN_THRESHOLDS.get(tech_key, GEN_MIN_THRESHOLDS['default'])

SRMC_CALC_THRESHOLDS = {str(r['tech']).strip(): _num(r['threshold_mw'])
                        for _, r in _g[_g['scope'] == 'srmc_calc'].iterrows()}

# ---- hydro / pumped params (single block: bare globals + CFG attributes) -----------
_h = pd.read_excel(_x, 'hydro_params')
HYDRO_RESERVOIR_PARAMS, PUMPED_STORAGE_PARAMS = {}, {}
for _, r in _h.iterrows():
    tgt = HYDRO_RESERVOIR_PARAMS if str(r['group']).strip() == 'HYDRO_RESERVOIR_PARAMS' \
          else PUMPED_STORAGE_PARAMS
    tgt[str(r['parameter']).strip()] = _num(r['value'])
for _grp, _d in [('HYDRO_RESERVOIR_PARAMS', HYDRO_RESERVOIR_PARAMS),
                 ('PUMPED_STORAGE_PARAMS', PUMPED_STORAGE_PARAMS)]:
    if not _d:
        raise RuntimeError(f"hydro_params sheet: no rows for {_grp}")
    setattr(CFG, _grp, _d)

# ---- gas ---------------------------------------------------------------------------
_gas = pd.read_excel(_x, 'gas')
GAS_HUB_MAP, GAS_TRANSPORT_ADDERS, IBERIAN_GAS_CAP = {}, {}, {}
for _, r in _gas.iterrows():
    g, k, v = str(r['group']).strip(), str(r['key']).strip(), r['value']
    if g == 'hub_map':
        GAS_HUB_MAP[k] = ast.literal_eval(str(v))          # repr strings - NOT _num
    elif g == 'transport_adder':
        GAS_TRANSPORT_ADDERS[k] = _num(v)
    elif g == 'iberian_cap':
        _s = re.sub(r"Timestamp\('([^']*)'\)", r"'\1'", str(v))   # repr - NOT _num
        try:
            _d = ast.literal_eval(_s)
            if isinstance(_d, dict):
                for _dk in ('start', 'end'):
                    if _dk in _d and isinstance(_d[_dk], str):
                        _d[_dk] = pd.Timestamp(_d[_dk])
            IBERIAN_GAS_CAP[k] = _d
        except (ValueError, SyntaxError):
            IBERIAN_GAS_CAP[k] = _s

# ---- commodities -------------------------------------------------------------------
_cm = pd.read_excel(_x, 'commodities')
CO2_SYMBOLS, COMMODITY_SYMBOLS = {}, {}
for _, r in _cm.iterrows():
    _t = CO2_SYMBOLS if str(r['group']).strip() == 'co2' else COMMODITY_SYMBOLS
    _v = str(r['value']).strip()
    try:
        _t[str(r['key']).strip()] = ast.literal_eval(_v)   # dicts/numbers (repr) - NOT _num
    except (ValueError, SyntaxError):
        _t[str(r['key']).strip()] = _v                     # bare symbol strings

# ---- static SRMC ranges (renewables O&M) -------------------------------------------
_ss = pd.read_excel(_x, 'srmc_static')
SRMC_STATIC = {str(r['tech']).strip(): (_num(r['srmc_min']), _num(r['srmc_max']))
               for _, r in _ss.iterrows()}
for _t in ('hydro_reservoir', 'hydro_pumped', 'nuclear'):
    SRMC_STATIC[_t] = None        # dynamic techs: formulas, not static ranges

# ---- efficiencies ------------------------------------------------------------------
_e = pd.read_excel(_x, 'efficiencies')
efficiency_maps = {}
for _, r in _e.iterrows():
    mname = f"{str(r['technology']).strip().upper()}_EFFICIENCY_MAP"
    efficiency_maps.setdefault(mname, {})[str(r['zone']).strip()] = {
        'min': _num(r['eff_min']), 'max': _num(r['eff_max']),
        'avg': _num(r['eff_avg']) if pd.notna(r['eff_avg'])
               else (_num(r['eff_min']) + _num(r['eff_max'])) / 2}
globals().update(efficiency_maps)          # expose GAS_EFFICIENCY_MAP etc. directly

# ---- efficiency accessor (ported from research Cell 3; part-load penalty removed) --
OIL_SHALE_EFFICIENCY_MAP = LIGNITE_EFFICIENCY_MAP.copy()   # oil shale uses lignite values

def get_efficiency(zone, fuel_type):
    """Return {'min','max','avg'} fleet efficiency for a zone and fuel."""
    _maps = {
        'gas': GAS_EFFICIENCY_MAP,
        'coal': COAL_EFFICIENCY_MAP,
        'lignite': LIGNITE_EFFICIENCY_MAP,
        'oil': OIL_EFFICIENCY_MAP,
        'oil_shale': OIL_SHALE_EFFICIENCY_MAP,
        'chp': CHP_EFFICIENCY_MAP,
        'biomass': BIOMASS_EFFICIENCY_MAP,
        'biogas': BIOGAS_EFFICIENCY_MAP,
    }
    eff_map = _maps.get(fuel_type, GAS_EFFICIENCY_MAP)

    # SMART FALLBACK: DE_AT_LU -> DE_LU -> DEFAULT
    if zone not in eff_map and zone == 'DE_AT_LU':
        zone = 'DE_LU'
    eff = eff_map.get(zone, eff_map.get('DEFAULT'))
    if eff is None:
        raise KeyError(f"No efficiency for zone={zone}, fuel={fuel_type} "
                       f"and no DEFAULT row in the efficiencies sheet")
    return eff

# ---- style -------------------------------------------------------------------------
_s = pd.read_excel(_x, 'style')
TECH_COLORS = dict(zip(_s.iloc[:, 0].astype(str).str.strip(),
                       _s.iloc[:, 1].astype(str).str.strip()))
CATEGORY_COLORS = {c: TECH_COLORS[c] for c in CATEGORIES_9 if c in TECH_COLORS}
_missing_col = [c for c in CATEGORIES_9 if c not in CATEGORY_COLORS]
if _missing_col:
    raise KeyError(f"style tab missing colors for categories: {_missing_col}")

# ---- shim: translation layer between CFG (control-file config) and the legacy ------
# ---- bare-name globals the production cells were written against -------------------
# 20261017: single full-period input file per data type (prices/generation/load/
# flows + RESERVOIR_TIME_SERIES); per-period DATA_PATHS_2015_2018 / DATA_PATHS_2026
# dicts removed together with their 'PREFIX:key' rows in the paths sheet.
DATA_PATHS = {k: Path(v) for k, v in CFG.paths.items()}


_need_paths = {'prices', 'generation', 'load', 'flows', 'reservoir',
               'OUTPUT_DIR', 'REFERENCE_DIR'}
_missp = _need_paths - set(CFG.paths)
if _missp:
    raise ValueError(f"paths sheet missing keys: {sorted(_missp)} "
                     f"(single-time-series layout expected)")

END_2026 = CFG.END_2026
START_DATE = CFG.START_DATE

UNIFORM_TOLERANCE    = CFG.UNIFORM_TOLERANCE
PRICE_TOLERANCE      = CFG.PRICE_TOLERANCE
FLOW_THRESHOLD       = CFG.FLOW_THRESHOLD
NIGHT_HOURS          = CFG.NIGHT_HOURS
OUTPUT_HEADLINE_ONLY = CFG.OUTPUT_HEADLINE_ONLY

YEAR_H = CFG.YEAR_HEADLINE

print("=" * 70)
print(f" Control file loaded: {CONTROL_FILE.name}")
print(f"   {CFG.START_DATE.date()} -> {CFG.END_2026.date()} | headline {CFG.YEAR_HEADLINE} | "
      f"headline-only: {CFG.OUTPUT_HEADLINE_ONLY}")
print(f"   zones {len(ALL_ZONES)} ({len(ZONES_EU_ONLY)} EU) | neighbour pairs "
      f"{sum(len(v) for v in NEIGHBOURS.values())} | proxies {len(RESERVOIR_FILLING_PROXIES)}")
print(f"   fuels {len(THERMAL_PARAMS)} | thresholds {len(GEN_MIN_THRESHOLDS)}+"
      f"{sum(len(v) for v in ZONE_THRESHOLDS.values())} overrides | "
      f"efficiency maps {len(efficiency_maps)} ({sum(len(v) for v in efficiency_maps.values())} zone-tech)")
print(f"   tolerances: uniform {CFG.UNIFORM_TOLERANCE} EUR | price {CFG.PRICE_TOLERANCE} EUR | "
      f"flow {CFG.FLOW_THRESHOLD} MW | fractional co-marginal: {CFG.COMARGINAL_FRACTIONAL}")
print(f"   hydro params: reservoir {len(HYDRO_RESERVOIR_PARAMS)} | pumped {len(PUMPED_STORAGE_PARAMS)}")
print("=" * 70)


# In[ ]:


# %% Cell 1 - IMPORTS and TIMER

# %% ==========================================================================
# Cell 0 - START TIMER
import time
from datetime import datetime

NOTEBOOK_START = time.time()
NOTEBOOK_START_TIME = datetime.now()

print("=" * 80)
print(f"📅 Notebook started: {NOTEBOOK_START_TIME.strftime('%Y-%m-%d %H:%M:%S')}")
print("=" * 80)


# ## 1. Import Required Libraries
# Import pandas, numpy, plotly, pathlib, datetime, spgci for Platts API, and configure display settings.

# In[2]:


# Import Required Libraries
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
from datetime import datetime, timedelta
import warnings

# Attempt to import spgci for Platts API (optional - will use fallback if unavailable)
try:
    import spgci
    SPGCI_AVAILABLE = True
except ImportError:
    SPGCI_AVAILABLE = False
    print("Warning: spgci not available. Will use fallback commodity prices.")

# Configure display settings
pd.set_option('display.max_columns', 50)
pd.set_option('display.max_rows', 100)
pd.set_option('display.width', 1000)
pd.set_option('display.float_format', '{:.2f}'.format)

warnings.filterwarnings('ignore', category=FutureWarning)

print("Libraries imported successfully")
print(f"Platts API (spgci) available: {SPGCI_AVAILABLE}")


# ## 2. Define Global Configuration (Paths, All BZN, Neighbours)
# Set BASE_DIR, define paths to new ENTSO-E CSVs, create ALL_ZONES list, and define NEIGHBOURS dictionary for all 47 bidding zones.

# In[3]:


# Base directory for data files
BASE_DIR   = CFG.paths['BASE_DIR']
DATA_DIR   = CFG.paths['DATA_DIR']


# ========================================
# UPDATED: PATHS TO ENTSO-E CSV FILES
# ========================================

# 2019-2025 period


# In[ ]:


# %% CELL 3 - GENERATION LOAD 
# %% ==========================================================================
#  Loads the single full-period generation file (see paths sheet: 'generation')
# Robust handling of ALL zone naming patterns
# Handles temporal format transitions (base + Actual Aggregated)
# Handles compound zones (DE_LU, DE_AT_LU, IT_NORD, etc.)
# Preserves DatetimeIndex
# DE_AT_LU treated as separate zone (2015-2018 Sep)
# NOW INCLUDES 2026 DATA

from pathlib import Path
import pandas as pd
import numpy as np
from collections import OrderedDict

print("=" * 80)
print("LOADING AND NORMALIZING ENTSOE GENERATION DATA (2015-2026)")
print("=" * 80)

# ========================================
# ENTSOE TO STANDARD MAPPING
# ========================================



# COMPREHENSIVE ZONE LIST (includes DE_AT_LU for 2015-2018)
KNOWN_ZONES = {
    # Single zones
    'AT', 'BE', 'BG', 'CH', 'CY', 'CZ', 'EE', 'ES', 'FI', 'FR', 'GB', 'GR', 'HR', 
    'HU', 'LT', 'LV', 'NL', 'PL', 'PT', 'RO', 'RS', 'SI', 'SK', 'AL', 'BA', 'ME', 'MK', 'MT', 'UA',
    
    # Compound zones (underscore)
    'DE_LU', 'DE_AT_LU', 'IE_SEM',  # Added DE_AT_LU
    
    # Nordic zones with numbers
    'DK_1', 'DK_2', 'NO_1', 'NO_2', 'NO_3', 'NO_4', 'NO_5', 'SE_1', 'SE_2', 'SE_3', 'SE_4',
    
    # Italian zones
    'IT_NORD', 'IT_CNOR', 'IT_CSUD', 'IT_SUD', 'IT_SICI', 'IT_SARD', 
    'IT_BRNN', 'IT_FOGN', 'IT_ROSN', 'IT_CALA',
}


def extract_zone_from_parts(parts):
    """
    Extract zone name from column name parts.
    
    Handles:
    - Simple zones: AT, FR, BE
    - Numbered zones: DK_1, NO_2, SE_4
    - Compound zones: DE_LU, DE_AT_LU, IT_NORD, IE_SEM
    
    Strategy: Try longest match first (4 parts, 3 parts, 2 parts, 1 part)
    """
    if len(parts) < 1:
        return None, None
    
    # Try 4-part zone (e.g., DE_AT_LU from ['DE', 'AT', 'LU', 'extra'])
    if len(parts) >= 4:
        zone_4 = f"{parts[-4]}_{parts[-3]}_{parts[-2]}_{parts[-1]}"
        if zone_4 in KNOWN_ZONES:
            return zone_4, parts[:-4]
    
    # Try 3-part zone (e.g., DE_AT_LU, IT_NORD from ['DE', 'AT', 'LU'])
    if len(parts) >= 3:
        zone_3 = f"{parts[-3]}_{parts[-2]}_{parts[-1]}"
        if zone_3 in KNOWN_ZONES:
            return zone_3, parts[:-3]
    
    # Try 2-part zone (e.g., DE_LU, DK_1, IT_NORD)
    if len(parts) >= 2:
        zone_2 = f"{parts[-2]}_{parts[-1]}"
        if zone_2 in KNOWN_ZONES:
            return zone_2, parts[:-2]
        
        # Also check for number pattern (NO_2, SE_4, etc.)
        if parts[-2].replace('-', '').isalpha() and parts[-1].isdigit():
            return zone_2, parts[:-2]
    
    # Try 1-part zone (e.g., AT, FR, BE)
    if len(parts) >= 1:
        zone_1 = parts[-1]
        if zone_1 in KNOWN_ZONES:
            return zone_1, parts[:-1]
        
        # Check if it looks like a zone (2-4 uppercase letters)
        if len(zone_1) <= 4 and zone_1.replace('-', '').isalpha():
            return zone_1, parts[:-1]
    
    return None, None


def normalize_entsoe_column_name(col_name):
    """
    Convert ENTSOE column name to standardized 'gen_{tech}_{zone}' format.
    
    Handles ALL zone naming patterns (including DE_AT_LU)
    Handles temporal format variations (base + Actual Aggregated)
    Excludes non-generation consumption columns
    
    Returns:
        Standardized column name, or None if should be excluded
    """
    # Skip metadata columns
    if col_name in ['DATE', 'YEAR', 'MONTH', 'DAY', 'HOUR']:
        return col_name
    
    # Already standardized
    if col_name.startswith('gen_'):
        return col_name
    
    # Identify column type
    is_consumption = 'Actual Consumption' in col_name
    is_actual_aggregated = 'Actual Aggregated' in col_name
    
    # Handle consumption columns
    if is_consumption:
        col_lower = col_name.lower()
        # Only keep pumped hydro and energy storage consumption
        if not ('hydro pumped' in col_lower or 'pumped storage' in col_lower or 'energy storage' in col_lower):
            return None  # Exclude other consumption
    
    # Parse column name
    parts = col_name.split('_')
    
    if len(parts) < 2:
        return None
    
    # Remove special keywords
    parts_cleaned = [p for p in parts 
                    if p not in ['Actual', 'Aggregated', 'Consumption', 'actual', 'aggregated', 'consumption']]
    
    if len(parts_cleaned) < 2:
        return None
    
    # Extract zone (robust matching)
    zone, tech_parts = extract_zone_from_parts(parts_cleaned)
    
    if not zone or not tech_parts:
        return None
    
    # Reconstruct technology name
    tech_name = '_'.join(tech_parts)
    
    # Map to standard technology name
    standard_tech = None
    
    # Try exact match
    tech_name_spaces = tech_name.replace('_', ' ')
    
    for entsoe_name, standard_name in ENTSOE_TO_STANDARD_MAPPING.items():
        if tech_name_spaces == entsoe_name or tech_name == entsoe_name:
            standard_tech = standard_name
            break
    
    if not standard_tech:
        # Fuzzy matching
        tech_lower = tech_name.lower().replace('_', ' ')
        
        if 'fossil gas' in tech_lower and 'coal-derived' not in tech_lower and 'derived' not in tech_lower:
            standard_tech = 'gas'
        elif 'hard coal' in tech_lower:
            standard_tech = 'coal'
        elif 'brown coal' in tech_lower or 'lignite' in tech_lower:
            standard_tech = 'lignite'
        # Oil shale (MUST check before regular oil to avoid false matches)
        elif 'oil shale' in tech_lower or 'oilshale' in tech_lower or 'oil-shale' in tech_lower:
            standard_tech = 'oil_shale'
        # Regular fossil oil (multiple patterns)
        elif 'fossil oil' in tech_lower or tech_lower == 'oil' or tech_lower.startswith('oil '):
            standard_tech = 'oil'
        elif 'coal-derived' in tech_lower or 'coal derived' in tech_lower:
            standard_tech = 'coal_derived_gas'
        elif 'peat' in tech_lower:
            standard_tech = 'peat'
        elif 'nuclear' in tech_lower:
            standard_tech = 'nuclear'
        elif 'hydro pumped' in tech_lower or 'pumped storage' in tech_lower:
            standard_tech = 'hydro_pumped'
        elif 'hydro water reservoir' in tech_lower or 'water reservoir' in tech_lower:
            standard_tech = 'hydro_reservoir'
        elif 'run-of-river' in tech_lower or 'run of river' in tech_lower:
            standard_tech = 'hydro_ror'
        elif 'wind onshore' in tech_lower:
            standard_tech = 'wind_onshore'
        elif 'wind offshore' in tech_lower:
            standard_tech = 'wind_offshore'
        elif 'solar' in tech_lower:
            standard_tech = 'solar'
        elif 'biomass' in tech_lower:
            standard_tech = 'biomass'
        elif 'waste' in tech_lower:
            standard_tech = 'waste'
        elif 'geothermal' in tech_lower:
            standard_tech = 'geothermal'
        elif 'other renewable' in tech_lower:
            standard_tech = 'other_renewable'
        elif 'marine' in tech_lower:
            standard_tech = 'marine'
        elif 'energy storage' in tech_lower:
            standard_tech = 'energy_storage'
        else:
            standard_tech = 'other'
    
    # Add "_consumption" suffix if this is a consumption column
    if is_consumption:
        standard_tech = f"{standard_tech}_consumption"
    
    return f"gen_{standard_tech}_{zone}"


def preprocess_entsoe_dataframe(df_raw, verbose=True):
    """
    Preprocess raw ENTSOE DataFrame to standardize column names.
    
    Handles temporal format transitions (sums non-overlapping data)
    Preserves DatetimeIndex
    """
    if verbose:
        print(f"\n🔄 Preprocessing ENTSOE data...")
        print(f"   Input shape: {df_raw.shape}")
    
    # Save original index
    original_index = df_raw.index.copy()
    
    # Normalize ALL columns
    column_mapping = {}
    consumption_columns_kept = []
    
    for col in df_raw.columns:
        normalized = normalize_entsoe_column_name(col)
        
        if normalized is not None and normalized != col and not col.startswith('gen_'):
            column_mapping[col] = normalized
            
            if '_consumption' in normalized:
                consumption_columns_kept.append(normalized)
    
    if verbose:
        print(f"   Columns to rename: {len(column_mapping)}")
        
        if consumption_columns_kept:
            print(f"    Consumption columns preserved: {len(set(consumption_columns_kept))}")
    
    # Apply renaming
    df_processed = df_raw.rename(columns=column_mapping)
    
    # Handle duplicates by summing (temporal transitions don't overlap)
    # Memory-safe version (20261011): compute all merges first, drop once -
    # the old per-column loop copied the whole frame on every iteration.
    duplicate_cols = df_processed.columns[df_processed.columns.duplicated()].unique()

    if len(duplicate_cols) > 0:
        if verbose:
            print(f"\n     Found {len(duplicate_cols)} duplicate columns (summing)")

        merged = {c: df_processed.loc[:, df_processed.columns == c].sum(axis=1)
                  for c in duplicate_cols}
        df_processed = df_processed.loc[:, ~df_processed.columns.isin(set(duplicate_cols))]
        for c, s in merged.items():
            df_processed[c] = s
    
    # Restore original DatetimeIndex
    df_processed.index = original_index
    
    if verbose:
        print(f"\n   Output shape: {df_processed.shape}")
        print(f"   Index type: {type(df_processed.index).__name__}")
        
        gen_cols = [c for c in df_processed.columns if c.startswith('gen_')]
        print(f"   Generation columns: {len(gen_cols)}")
    
    return df_processed


# ========== FULL PERIOD (single merged input file) ==========
print(f"\n{'='*80}")
print(f"LOADING GENERATION DATA (FULL PERIOD)")
print(f"{'='*80}")

ENTSOE_GEN_FILE = DATA_PATHS['generation']

print(f"\n📂 Loading from configured path:")
print(f"   {ENTSOE_GEN_FILE}")

if not ENTSOE_GEN_FILE.exists():
    raise FileNotFoundError(f"ENTSOE generation file not found at: {ENTSOE_GEN_FILE}")

print(f"   File found")
df_gen_raw = pd.read_csv(ENTSOE_GEN_FILE, index_col=0, parse_dates=True)
print(f"   Raw shape: {df_gen_raw.shape}")
print(f"   Index type (initial): {type(df_gen_raw.index).__name__}")

# Fix timezone issues
if not isinstance(df_gen_raw.index, pd.DatetimeIndex):
    print(f"     Converting index to DatetimeIndex...")
    try:
        df_gen_raw.index = pd.to_datetime(df_gen_raw.index, utc=True).tz_localize(None)
    except:
        df_gen_raw.index = pd.to_datetime(df_gen_raw.index)
else:
    if df_gen_raw.index.tz is not None:
        print(f"     Removing timezone from index...")
        df_gen_raw.index = df_gen_raw.index.tz_localize(None)

print(f"   Date range: {df_gen_raw.index.min()} to {df_gen_raw.index.max()}")

# Normalize
df_gen = preprocess_entsoe_dataframe(df_gen_raw, verbose=True)
del df_gen_raw
gc.collect()

df_gen = df_gen.sort_index()


print(f"   Combined shape: {df_gen.shape}")
print(f"   Date range: {df_gen.index.min()} to {df_gen.index.max()}")
print(f"   Years covered: {sorted(df_gen.index.year.unique())}")


# Verify
if not isinstance(df_gen.index, pd.DatetimeIndex):
    raise TypeError("DatetimeIndex was not preserved during preprocessing")

print(f"\n{'='*80}")
print(f"DATA LOADED AND NORMALIZED (2015-2026)")
print(f"{'='*80}")

# Check sample years
gen_cols = [c for c in df_gen.columns if c.startswith('gen_')]

for year in [2015, 2018, 2019, 2025, 2026]:
    mask_year = (df_gen.index.year == year)
    if mask_year.sum() > 0:
        total_year = df_gen.loc[mask_year, gen_cols].sum().sum()
        print(f"\n📊 Year {year}:")
        print(f"   Rows: {mask_year.sum():,}")
        print(f"   Total generation: {total_year:,.0f} MWh")

# Sample zone check across periods
test_zones = ['DE_AT_LU', 'DE_LU', 'AT', 'DK_1', 'FR']
print(f"\n📍 Sample zones:")

for zone in test_zones:
    zone_cols = [c for c in gen_cols if c.endswith(f'_{zone}')]
    
    if zone_cols:
        # Check 2015, 2018, 2019, and 2026
        total_2015 = df_gen.loc[df_gen.index.year == 2015, zone_cols].sum().sum() if (df_gen.index.year == 2015).any() else 0
        total_2018 = df_gen.loc[df_gen.index.year == 2018, zone_cols].sum().sum() if (df_gen.index.year == 2018).any() else 0
        total_2019 = df_gen.loc[df_gen.index.year == 2019, zone_cols].sum().sum() if (df_gen.index.year == 2019).any() else 0
        total_2026 = df_gen.loc[df_gen.index.year == 2026, zone_cols].sum().sum() if (df_gen.index.year == 2026).any() else 0
        
        status_2015 = "OK" if total_2015 > 0 else "—"
        status_2018 = "OK" if total_2018 > 0 else "—"
        status_2019 = "OK" if total_2019 > 0 else "—"
        status_2026 = "OK" if total_2026 > 0 else "—"
        
        print(f"   {zone}: {len(zone_cols)} cols")
        print(f"      2015: {status_2015} {total_2015:,.0f} MWh")
        print(f"      2018: {status_2018} {total_2018:,.0f} MWh")
        print(f"      2019: {status_2019} {total_2019:,.0f} MWh")
        print(f"      2026: {status_2026} {total_2026:,.0f} MWh")
    else:
        print(f"   {zone}: NO COLUMNS")

print(f"\n{'='*80}")
print(f"CELL 3.5 COMPLETE - df_gen ready for use (2015-2026)")
print(f"{'='*80}")


# In[ ]:


# %% CELL 4 - COLUMN MAPPINGS (research Cell 4)
# %% ==========================================================================
# Cell 4: Column Name Mappings for ENTSO-E Data
# Pattern definitions for hybrid column logic (temporal transitions)
# Works with normalized data from Cell 3.5
# Handles DE_AT_LU, DE_LU, and AT as separate zones

import pandas as pd
import numpy as np
from collections import OrderedDict

print("=" * 80)
print("GENERATION COLUMN PATTERN DEFINITIONS (2015-2025)")
print("=" * 80)

# ========================================
# PATTERN DICTIONARY
# ========================================

GEN_COL_PATTERNS = {
    'gas': {
        'new_patterns': [
            'Fossil Gas_Actual Aggregated',
            'Fossil Gas_Actual_Aggregated',
            'Fossil_Gas_Actual_Aggregated',
            'Fossil_Gas_Actual Aggregated',
        ],
        'old_patterns': [
            'Fossil Gas',
            'Fossil_Gas',
        ]
    },
    
    'coal': {
        'new_patterns': [
            'Fossil Hard coal_Actual Aggregated',
            'Fossil Hard Coal_Actual Aggregated',
            'Fossil_Hard_coal_Actual_Aggregated',
            'Fossil_Hard_Coal_Actual_Aggregated',
        ],
        'old_patterns': [
            'Fossil Hard coal',
            'Fossil Hard Coal',
            'Fossil_Hard_coal',
        ]
    },
    
    'lignite': {
        'new_patterns': [
            'Fossil Brown coal/Lignite_Actual Aggregated',
            'Fossil_Brown_coal_Lignite_Actual_Aggregated',
        ],
        'old_patterns': [
            'Fossil Brown coal/Lignite',
            'Fossil_Brown_coal_Lignite',
        ]
    },
    
    'oil': {
        'new_patterns': [
            'Fossil Oil_Actual Aggregated',
            'Fossil_Oil_Actual_Aggregated',
        ],
        'old_patterns': [
            'Fossil Oil',
            'Fossil_Oil',
            'Oil',
        ]
    },
    
    # Oil shale patterns
    'oil_shale': {
        'new_patterns': [
            'Fossil Oil shale_Actual Aggregated',
            'Fossil_Oil_shale_Actual_Aggregated',
            'Fossil_Oilshale_Actual_Aggregated',
        ],
        'old_patterns': [
            'Fossil Oil shale',
            'Fossil_Oil_shale',
            'Fossil_Oilshale',
            'Oil shale',
            'Oil_shale',
            'Oilshale',
        ]
    },
    
    'peat': {
        'new_patterns': [
            'Fossil Peat_Actual Aggregated',
        ],
        'old_patterns': [
            'Fossil Peat',
            'Peat',
        ]
    },
    
    'coal_derived_gas': {
        'new_patterns': [
            'Fossil Coal-derived gas_Actual Aggregated',
        ],
        'old_patterns': [
            'Fossil Coal-derived gas',
            'Coal-derived gas',
        ]
    },
    
    'nuclear': {
        'new_patterns': [
            'Nuclear_Actual Aggregated',
        ],
        'old_patterns': [
            'Nuclear',
        ]
    },
    
    'hydro_reservoir': {
        'new_patterns': [
            'Hydro Water Reservoir_Actual Aggregated',
            'Hydro_Water_Reservoir_Actual_Aggregated',
        ],
        'old_patterns': [
            'Hydro Water Reservoir',
            'Hydro_Water_Reservoir',
            'Water Reservoir',
        ]
    },
    
    'hydro_ror': {
        'new_patterns': [
            'Hydro Run-of-river and poundage_Actual Aggregated',
            'Hydro_Run-of-river_and_poundage_Actual_Aggregated',
        ],
        'old_patterns': [
            'Hydro Run-of-river and poundage',
            'Hydro_Run-of-river_and_poundage',
            'Run-of-river and poundage',
        ]
    },
    
    'hydro_pumped': {
        'new_patterns': [
            'Hydro Pumped Storage_Actual Aggregated',
            'Hydro_Pumped_Storage_Actual_Aggregated',
        ],
        'old_patterns': [
            'Hydro Pumped Storage',
            'Hydro_Pumped_Storage',
            'Pumped Storage',
        ]
    },
    
    'hydro_pumped_consumption': {
        'new_patterns': [
            'Hydro Pumped Storage_Actual Consumption',
            'Hydro_Pumped_Storage_Actual_Consumption',
        ],
        'old_patterns': [
            'Hydro Pumped Storage Consumption',
        ]
    },
    
    'energy_storage': {
        'new_patterns': [
            'Energy storage_Actual Aggregated',
            'Energy_storage_Actual_Aggregated',
        ],
        'old_patterns': [
            'Energy storage',
            'Energy_storage',
        ]
    },
    
    'energy_storage_consumption': {
        'new_patterns': [
            'Energy storage_Actual Consumption',
            'Energy_storage_Actual_Consumption',
        ],
        'old_patterns': [
            'Energy storage Consumption',
            'Energy_storage_Consumption',
        ]
    },
    
    'wind_onshore': {
        'new_patterns': [
            'Wind Onshore_Actual Aggregated',
            'Wind_Onshore_Actual_Aggregated',
        ],
        'old_patterns': [
            'Wind Onshore',
            'Wind_Onshore',
        ]
    },
    
    'wind_offshore': {
        'new_patterns': [
            'Wind Offshore_Actual Aggregated',
            'Wind_Offshore_Actual_Aggregated',
        ],
        'old_patterns': [
            'Wind Offshore',
            'Wind_Offshore',
        ]
    },
    
    'solar': {
        'new_patterns': [
            'Solar_Actual Aggregated',
        ],
        'old_patterns': [
            'Solar',
        ]
    },
    
    'biomass': {
        'new_patterns': [
            'Biomass_Actual Aggregated',
        ],
        'old_patterns': [
            'Biomass',
        ]
    },
    
    'geothermal': {
        'new_patterns': [
            'Geothermal_Actual Aggregated',
        ],
        'old_patterns': [
            'Geothermal',
        ]
    },
    
    'marine': {
        'new_patterns': [
            'Marine_Actual Aggregated',
        ],
        'old_patterns': [
            'Marine',
        ]
    },
    
    'waste': {
        'new_patterns': [
            'Waste_Actual Aggregated',
        ],
        'old_patterns': [
            'Waste',
        ]
    },
    
    'other_renewable': {
        'new_patterns': [
            'Other renewable_Actual Aggregated',
            'Other_renewable_Actual_Aggregated',
        ],
        'old_patterns': [
            'Other renewable',
            'Other_renewable',
        ]
    },
    
    'other': {
        'new_patterns': [
            'Other_Actual Aggregated',
        ],
        'old_patterns': [
            'Other',
        ]
    },
}


# ========================================
# HELPER FUNCTIONS
# ========================================

def find_matching_columns(df, tech_type, zone=None, format_type='both'):
    """
    Find ALL matching columns for a technology type.
    Works with normalized data from Cell 3.5
    Handles temporal transitions (OLD→NEW format)
    Supports DE_AT_LU, DE_LU, and AT as separate zones
    
    Args:
        df: DataFrame with generation data
        tech_type: Technology type (e.g., 'gas', 'coal', 'nuclear')
        zone: Zone code (e.g., 'DE_AT_LU', 'DE_LU', 'AT', 'FR')
        format_type: 'both', 'new', or 'old'
    
    Returns:
        dict with 'new' and 'old' keys containing matching column names
    """
    patterns_dict = GEN_COL_PATTERNS.get(tech_type, {'new_patterns': [], 'old_patterns': []})
    
    found_columns = {'new': [], 'old': []}
    
    # PRIORITY 1: Check for already-normalized columns (from Cell 3.5)
    if zone:
        normalized_col = f"gen_{tech_type}_{zone}"
        if normalized_col in df.columns:
            # Data already normalized - return immediately
            found_columns['old'].append(normalized_col)
            return found_columns
    
    # PRIORITY 2: Search using patterns (for non-normalized data)
    if format_type in ['both', 'new']:
        for pattern in patterns_dict.get('new_patterns', []):
            if zone:
                # Zone-specific search
                possible_cols = [
                    f"{pattern}_{zone}",
                    f"{pattern.replace(' ', '_')}_{zone}",
                ]
                for col_name in possible_cols:
                    if col_name in df.columns:
                        found_columns['new'].append(col_name)
            else:
                # Search all columns matching pattern
                for col in df.columns:
                    if pattern in col:
                        found_columns['new'].append(col)
    
    if format_type in ['both', 'old']:
        for pattern in patterns_dict.get('old_patterns', []):
            if zone:
                # Zone-specific search
                possible_cols = [
                    f"{pattern}_{zone}",
                    f"{pattern.replace(' ', '_')}_{zone}",
                ]
                for col_name in possible_cols:
                    if col_name in df.columns:
                        found_columns['old'].append(col_name)
            else:
                # Search all columns matching pattern
                for col in df.columns:
                    if pattern in col:
                        found_columns['old'].append(col)
    
    # Remove duplicates while preserving order
    found_columns['new'] = list(dict.fromkeys(found_columns['new']))
    found_columns['old'] = list(dict.fromkeys(found_columns['old']))
    
    return found_columns


def find_price_column(df, zone):
    """
    Find price column for a specific zone.
    Handles DE_AT_LU, DE_LU, and AT
    """
    # Normalized price column (from Cell 3.x)
    normalized_patterns = [f'price_{zone}', f'Price_{zone}']
    
    # Legacy patterns
    legacy_patterns = [
        f'DA_price_{zone}',
        f'{zone}_price',
        f'{zone}_Price',
        f'price_{zone.upper()}',
        f'Price_{zone.upper()}',
    ]
    
    all_patterns = normalized_patterns + legacy_patterns
    
    for pattern in all_patterns:
        for col in df.columns:
            if pattern.lower() == col.lower() or pattern == col:
                return col
    
    # Fuzzy match as last resort
    zone_upper = zone.upper()
    for col in df.columns:
        if 'price' in col.lower() and zone_upper in col.upper():
            return col
    
    return None


def find_load_column(df, zone):
    """
    Find load column for a specific zone.
    Handles DE_AT_LU, DE_LU, and AT
    """
    # Normalized load column
    normalized_patterns = [f'Load_{zone}', f'load_{zone}']
    
    # Legacy patterns
    legacy_patterns = [
        f'Actual_Total_Load_{zone}',
        f'{zone}_load',
        f'{zone}_Load',
        f'load_{zone.upper()}',
        f'Load_{zone.upper()}',
    ]
    
    all_patterns = normalized_patterns + legacy_patterns
    
    for pattern in all_patterns:
        for col in df.columns:
            if pattern.lower() == col.lower() or pattern == col:
                return col
    
    # Fuzzy match as last resort
    zone_upper = zone.upper()
    for col in df.columns:
        if 'load' in col.lower() and zone_upper in col.upper():
            return col
    
    return None


def find_flow_columns(df, zone):
    """
    Find all flow columns related to a zone.
    Handles DE_AT_LU, DE_LU, and AT
    Handles both 'Flow_X_to_Y' and 'X > Y' formats
    """
    imports, exports = [], []
    zone_upper = zone.upper()
    
    for col in df.columns:
        col_upper = col.upper()
        
        # Import patterns (flow TO this zone)
        import_patterns = [
            f'_TO_{zone_upper}',
            f'> {zone_upper}',
            f'>{zone_upper}',
            f'-{zone_upper}',
            f'_{zone_upper}$',  # Ends with zone
        ]
        
        for pattern in import_patterns:
            if pattern in col_upper:
                imports.append(col)
                break
        
        # Export patterns (flow FROM this zone)
        export_patterns = [
            f'{zone_upper}_TO_',
            f'{zone_upper} >',
            f'{zone_upper}>',
            f'{zone_upper}-',
            f'^{zone_upper}_',  # Starts with zone
        ]
        
        for pattern in export_patterns:
            if pattern in col_upper or col_upper.startswith(zone_upper + '_'):
                exports.append(col)
                break
    
    return {'imports': list(set(imports)), 'exports': list(set(exports))}


print("Column pattern definitions loaded")
print(f"   Technologies defined: {len(GEN_COL_PATTERNS)}")
print(f"   Includes: gas, coal, lignite, oil, oil_shale, peat, nuclear, hydro variants, renewables")
print(f"   Consumption patterns: hydro_pumped_consumption, energy_storage_consumption")
print(f"   Zone handling: DE_AT_LU (2015-2018), DE_LU (2019+), AT (2019+) as separate zones")


# In[ ]:


# %% CELL 5 - PRICES/LOAD/FLOWS (research Cell 5)
# %% ==========================================================================
# Cell 5: Load ENTSO-E Data with Hybrid Column Processing (2015-2026)
# Works with pre-normalized df_gen from Cell 3.5
# Handles DE_AT_LU, DE_LU, and AT as separate zones
# FX conversion DISABLED - prices already converted in source CSV
# NOW INCLUDES 2026 DATA

import requests
import io
from pathlib import Path
import pandas as pd
import numpy as np

print("="*80)
print("LOADING ENTSO-E DATA (PRICES, LOAD, FLOWS) - 2015-2026")
print("="*80)


def load_entsoe_data_combined(data_paths_2015, data_paths_2019, data_paths_2026):
    """
    Load all ENTSO-E data files from all three periods and combine them.
    Returns DataFrames with DATETIME INDEX
    Checks if generation data already loaded in Cell 3.5
    Combines 2015-2018, 2019-2025, and 2026 data
    """
    data = {}
    
    datetime_cols = ['datetime', 'DateTime', 'Datetime', 'timestamp', 'Timestamp', 'time', 'Time', 'date', 'Date']
    
    for key in ['prices', 'generation', 'load', 'flows']:
        print(f"\n{'='*60}")
        print(f"Loading {key} data...")
        print(f"{'='*60}")
        
        #   SPECIAL HANDLING: Check if generation data already loaded in Cell 3.5
        if key == 'generation':
            if 'df_gen' in globals() and df_gen is not None and not df_gen.empty:
                print(f"    Generation data already loaded and normalized in Cell 3.5")
                print(f"  ✓ Using existing df_gen: {df_gen.shape[0]:,} rows × {df_gen.shape[1]} columns")
                data[key] = df_gen
                continue
            else:
                print(f"    df_gen not found - will load and combine from files")
        
        # Get paths for all three periods
        path_2015 = data_paths_2015.get(key)
        path_2019 = data_paths_2019.get(key)
        path_2026 = data_paths_2026.get(key)
        
        dfs_to_combine = []
        
        # Load 2015-2018 data
        if path_2015 and Path(path_2015).exists():
            print(f"\n  Loading 2015-2018 file: {path_2015}")
            try:
                df_preview = pd.read_csv(path_2015, nrows=5)
                datetime_col = None
                
                for col in df_preview.columns:
                    if col.lower() in [c.lower() for c in datetime_cols]:
                        datetime_col = col
                        break
                    if 'date' in col.lower() or 'time' in col.lower():
                        datetime_col = col
                        break
                
                if datetime_col:
                    df_2015 = pd.read_csv(path_2015, index_col=datetime_col, parse_dates=True)
                else:
                    df_2015 = pd.read_csv(path_2015, index_col=0, parse_dates=True)
                
                # Ensure datetime index
                if not isinstance(df_2015.index, pd.DatetimeIndex):
                    try:
                        df_2015.index = pd.to_datetime(df_2015.index, utc=True).tz_localize(None)
                    except:
                        df_2015.index = pd.to_datetime(df_2015.index)
                else:
                    if df_2015.index.tz is not None:
                        df_2015.index = df_2015.index.tz_localize(None)
                
                df_2015 = df_2015.sort_index()
                dfs_to_combine.append(df_2015)
                
                print(f"    ✓ Loaded {len(df_2015):,} rows, {len(df_2015.columns)} columns")
                print(f"    ✓ Date range: {df_2015.index.min()} to {df_2015.index.max()}")
                
            except Exception as e:
                print(f"     ERROR loading 2015-2018 {key}: {e}")
        else:
            print(f"\n    2015-2018 file not found: {path_2015}")
        
        # Load 2019-2025 data
        if path_2019 and Path(path_2019).exists():
            print(f"\n  Loading 2019-2025 file: {path_2019}")
            try:
                df_preview = pd.read_csv(path_2019, nrows=5)
                datetime_col = None
                
                for col in df_preview.columns:
                    if col.lower() in [c.lower() for c in datetime_cols]:
                        datetime_col = col
                        break
                    if 'date' in col.lower() or 'time' in col.lower():
                        datetime_col = col
                        break
                
                if datetime_col:
                    df_2019 = pd.read_csv(path_2019, index_col=datetime_col, parse_dates=True)
                else:
                    df_2019 = pd.read_csv(path_2019, index_col=0, parse_dates=True)
                
                # Ensure datetime index
                if not isinstance(df_2019.index, pd.DatetimeIndex):
                    try:
                        df_2019.index = pd.to_datetime(df_2019.index, utc=True).tz_localize(None)
                    except:
                        df_2019.index = pd.to_datetime(df_2019.index)
                else:
                    if df_2019.index.tz is not None:
                        df_2019.index = df_2019.index.tz_localize(None)
                
                df_2019 = df_2019.sort_index()
                dfs_to_combine.append(df_2019)
                
                print(f"    ✓ Loaded {len(df_2019):,} rows, {len(df_2019.columns)} columns")
                print(f"    ✓ Date range: {df_2019.index.min()} to {df_2019.index.max()}")
                
            except Exception as e:
                print(f"     ERROR loading 2019-2025 {key}: {e}")
        else:
            print(f"\n    2019-2025 file not found: {path_2019}")
        
        # ========== LOAD 2026 DATA (NEW) ==========
        if path_2026 and Path(path_2026).exists():
            print(f"\n  Loading 2026 file: {path_2026}")
            try:
                df_preview = pd.read_csv(path_2026, nrows=5)
                datetime_col = None
                
                for col in df_preview.columns:
                    if col.lower() in [c.lower() for c in datetime_cols]:
                        datetime_col = col
                        break
                    if 'date' in col.lower() or 'time' in col.lower():
                        datetime_col = col
                        break
                
                if datetime_col:
                    df_2026 = pd.read_csv(path_2026, index_col=datetime_col, parse_dates=True)
                else:
                    df_2026 = pd.read_csv(path_2026, index_col=0, parse_dates=True)
                
                # Ensure datetime index
                if not isinstance(df_2026.index, pd.DatetimeIndex):
                    try:
                        df_2026.index = pd.to_datetime(df_2026.index, utc=True).tz_localize(None)
                    except:
                        df_2026.index = pd.to_datetime(df_2026.index)
                else:
                    if df_2026.index.tz is not None:
                        df_2026.index = df_2026.index.tz_localize(None)
                
                df_2026 = df_2026.sort_index()
                #  H1 2026 ONLY: curtail to 1 Jan - 30 Jun 2026. FIXED applied on 27/08/2026
                df_2026 = df_2026.loc[:'2026-06-30']
                dfs_to_combine.append(df_2026)
                
                print(f"    ✓ Loaded {len(df_2026):,} rows, {len(df_2026.columns)} columns")
                print(f"    ✓ Date range: {df_2026.index.min()} to {df_2026.index.max()}")
                
            except Exception as e:
                print(f"     ERROR loading 2026 {key}: {e}")
        else:
            print(f"\n    2026 file not found: {path_2026}")
        
        # Combine all periods
        if len(dfs_to_combine) >= 2:
            print(f"\n  🔗 Combining {len(dfs_to_combine)} periods...")
            df_combined = pd.concat(dfs_to_combine, axis=0, sort=False)
            df_combined = df_combined.sort_index()
            # df_combined = df_combined.fillna(0)
            #  FIX 27/08/2026: do NOT fillna(0). Missing prices/load/generation must stay
            # NaN so gap hours are dropped in Cell 10 (dropna on 'price'),
            # exactly like the colleague's loader. fillna(0) fabricated
            # 0 €/MWh hours in 2015 that direct-matched wind/solar (SRMC 0-15)
            
            print(f"    ✓ Combined: {len(df_combined):,} rows, {len(df_combined.columns)} columns")
            print(f"    ✓ Date range: {df_combined.index.min()} to {df_combined.index.max()}")
            
            data[key] = df_combined
            
        elif len(dfs_to_combine) == 1:
            print(f"\n    Only one period available for {key}")
            data[key] = dfs_to_combine[0]
            
        else:
            print(f"\n   No data files found for {key}")
            data[key] = pd.DataFrame()
    
    return data


def identify_available_zones(data):
    """
    Identify which zones have data available across all datasets.
     Works with normalized 'gen_{tech}_{zone}' format from Cell 3.5
     Handles DE_AT_LU, DE_LU, and AT as separate zones
    """
    availability = {zone: {'prices': False, 'generation': False, 'load': False, 'flows': False} 
                   for zone in ALL_ZONES}
    
    # Check prices
    if not data.get('prices', pd.DataFrame()).empty:
        for zone in ALL_ZONES:
            col = find_price_column(data['prices'], zone)
            if col is not None:
                availability[zone]['prices'] = True
    
    #  Check generation - works with normalized 'gen_{tech}_{zone}' columns
    if not data.get('generation', pd.DataFrame()).empty:
        df_gen_check = data['generation']
        
        for zone in ALL_ZONES:
            # Check if any normalized columns exist for this zone
            zone_cols = [col for col in df_gen_check.columns 
                        if col.startswith('gen_') and col.endswith(f'_{zone}')]
            
            if len(zone_cols) > 0:
                # Check if any column has actual data (not all zeros)
                has_data = False
                for col in zone_cols[:5]:  # Check first 5 columns
                    if df_gen_check[col].sum() > 0:
                        has_data = True
                        break
                
                availability[zone]['generation'] = has_data
    
    # Check load
    if not data.get('load', pd.DataFrame()).empty:
        for zone in ALL_ZONES:
            col = find_load_column(data['load'], zone)
            if col is not None:
                availability[zone]['load'] = True
    
    # Check flows
    if not data.get('flows', pd.DataFrame()).empty:
        for zone in ALL_ZONES:
            flow_cols = find_flow_columns(data['flows'], zone)
            if len(flow_cols['imports']) > 0 or len(flow_cols['exports']) > 0:
                availability[zone]['flows'] = True
    
    # Summarize
    complete_zones = [z for z, v in availability.items() 
                     if v['prices'] and v['generation']]
    partial_zones = [z for z, v in availability.items() 
                    if v['prices'] or v['generation']]
    
    print(f"\n{'='*80}")
    print(f"ZONE AVAILABILITY SUMMARY (2015-2026)")
    print(f"{'='*80}")
    print(f"\n  Complete data (prices + generation): {len(complete_zones)} zones")
    print(f"  Partial data: {len(partial_zones)} zones")
    
    # Show detailed availability
    print(f"\n📊 Detailed Availability:")
    for zone in sorted(availability.keys()):
        v = availability[zone]
        status = []
        if v['prices']: status.append('P')
        if v['generation']: status.append('G')
        if v['load']: status.append('L')
        if v['flows']: status.append('F')
        
        status_str = '|'.join(status) if status else 'NONE'
        
        if v['prices'] and v['generation']:
            print(f"   {zone:10s}: {status_str}")
        elif v['prices'] or v['generation']:
            print(f"    {zone:10s}: {status_str}")
    
    return availability, complete_zones


# ========================================
# MAIN EXECUTION
# ========================================

print(f"\n{'='*80}")
print(f"USING DATA PATHS FROM CELL 2:")
print(f"{'='*80}")

print(f"\ninput files:")
for key, path in DATA_PATHS.items():
    exists_str = "✓" if Path(path).exists() else "✗"
    print(f"  {exists_str} {key:12s}: {path}")

# CHECK: Is generation data already loaded from Cell 3.5?
print(f"\n{'='*80}")
print(f"CHECKING FOR PRE-LOADED GENERATION DATA")
print(f"{'='*80}")

if 'df_gen' in globals() and df_gen is not None and not df_gen.empty:
    print(f"\nGeneration data already loaded from Cell 3.5")
    print(f"   Shape: {df_gen.shape}")
    print(f"   Date range: {df_gen.index.min()} to {df_gen.index.max()}")
    
    # Verify normalization
    sample_cols = [c for c in df_gen.columns[:10] if '_' in c]
    normalized = all(c.startswith('gen_') for c in sample_cols)
    print(f"   Columns normalized: {normalized}")
    
    if normalized:
        print(f"    Ready to use!")
    else:
        print(f"     Columns may not be normalized - check Cell 3.5")
else:
    print(f"\n  Generation data not found from Cell 3.5")
    print(f"   Will load fresh from files (slower)")

# Load other datasets (prices, load, flows) - combines all three periods
print(f"\n{'='*80}")
print(f"LOADING AND COMBINING DATASETS (2015-2026)")
print(f"{'='*80}")

entsoe_data = load_entsoe_data_combined({}, DATA_PATHS, {})   # single merged file per type

# Extract dataframes
df_prices = entsoe_data.get('prices', pd.DataFrame())
df_gen = entsoe_data.get('generation', pd.DataFrame())  # Will be pre-loaded if Cell 3.5 ran
df_load = entsoe_data.get('load', pd.DataFrame())
df_flows = entsoe_data.get('flows', pd.DataFrame())

# 20260910 Added to cap inouyt ENTSO-E data to 1H 2026. Comment or delete to have year up to date or full 20026 when the data is complete
# Cap all inputs at H1 2026 (END_2026 from Cell 2)
for _name in ['df_prices', 'df_gen', 'df_load', 'df_flows']:
    _df = globals().get(_name)
    if _df is not None and len(_df) > 0:
        globals()[_name] = _df.loc[:END_2026]
        print(f"   ✂️ {_name}: capped at {END_2026} -> ends {globals()[_name].index.max()}")

# ========================================
# FX CONVERSION DISABLED
# ========================================

print("\n" + "="*80)
print(" FX CONVERSION SKIPPED - PRICES ALREADY CONVERTED IN SOURCE CSV")
print("="*80)

# ========================================
# IDENTIFY AVAILABLE ZONES
# ========================================

zone_availability, complete_zones = identify_available_zones(entsoe_data)

# ========================================
# VERIFY GENERATION FORMAT
# ========================================

if not df_gen.empty:
    print(f"\n{'='*80}")
    print(f"GENERATION DATA FORMAT VERIFICATION")
    print(f"{'='*80}")
    
    # Check if already normalized
    sample_cols = [c for c in df_gen.columns[:10] if '_' in c]
    already_normalized = all(c.startswith('gen_') for c in sample_cols)
    
    if already_normalized:
        print(f"\n Generation data already normalized (from Cell 3.5)")
        print(f"   No additional processing needed")
        
        df_gen_unified = df_gen
        
        # Show sample
        print(f"\n   Sample columns:")
        for col in sample_cols[:5]:
            print(f"     • {col}")
        
        # Count zones
        zones_in_data = set()
        for col in df_gen.columns:
            if col.startswith('gen_'):
                parts = col.split('_')
                if len(parts) >= 3:
                    # Handle compound zones (DE_AT_LU, IT_NORD)
                    if len(parts) >= 5:
                        zone = '_'.join(parts[2:])
                    else:
                        zone = '_'.join(parts[2:])
                    
                    # Remove _consumption suffix if present
                    if zone.endswith('_consumption'):
                        zone = '_'.join(zone.split('_')[:-1])
                    zones_in_data.add(zone)
        
        print(f"\n   Zones with generation data: {len(zones_in_data)}")
        
        # Check for DE_AT_LU, DE_LU, AT
        special_zones = ['DE_AT_LU', 'DE_LU', 'AT']
        print(f"\n   Zone transition check:")
        for zone in special_zones:
            if zone in zones_in_data:
                zone_cols = [c for c in df_gen.columns if c.startswith('gen_') and c.endswith(f'_{zone}')]
                total = df_gen[zone_cols].sum().sum()
                print(f"     {zone}: {len(zone_cols)} cols, {total:,.0f} MWh")
            else:
                print(f"     — {zone}: not found")
        
    else:
        print(f"\n  Generation data NOT normalized")
        print(f"   Running normalization now (should have been done in Cell 3.5)...")
        
        # Fallback: normalize now
        df_gen = preprocess_entsoe_dataframe(df_gen, verbose=True)
        df_gen_unified = df_gen
    
    # Update in data dict
    entsoe_data['generation'] = df_gen_unified
    df_gen = df_gen_unified

else:
    print(f"\n No generation data loaded")

# ========================================
# FINAL SUMMARY
# ========================================

print(f"\n{'='*80}")
print(f" DATA LOADING COMPLETE (2015-2026)")
print(f"{'='*80}")

print(f"\n📊 Dataset Summary:")
print(f"   Prices:     {len(df_prices):>8,} rows × {len(df_prices.columns):>4} columns")
print(f"   Generation: {len(df_gen):>8,} rows × {len(df_gen.columns):>4} columns")
print(f"   Load:       {len(df_load):>8,} rows × {len(df_load.columns):>4} columns")
print(f"   Flows:      {len(df_flows):>8,} rows × {len(df_flows.columns):>4} columns")

if not df_gen.empty:
    print(f"\n📅 Date ranges:")
    print(f"   Prices:     {df_prices.index.min()} to {df_prices.index.max()}")
    print(f"   Generation: {df_gen.index.min()} to {df_gen.index.max()}")
    print(f"   Load:       {df_load.index.min()} to {df_load.index.max()}")
    if not df_flows.empty:
        print(f"   Flows:      {df_flows.index.min()} to {df_flows.index.max()}")

print(f"\n📍 Zone Summary:")
print(f"   Total zones configured: {len(ALL_ZONES)}")
print(f"   Zones with complete data: {len(complete_zones)}")

if complete_zones:
    print(f"\n   Complete zones:")
    # Print in rows of 5
    for i in range(0, len(complete_zones), 5):
        zone_group = complete_zones[i:i+5]
        print(f"     {', '.join(zone_group)}")

print(f"\n{'='*80}")
print(f"Ready for analysis!")
print(f"{'='*80}")


# In[11]:


# FINAL VERIFICATION: Check entire pipeline (2015-2026)
# Updated to verify DE_AT_LU, DE_LU, and AT zones
# NOW INCLUDES 2026 VERIFICATION

print("="*80)
print("FINAL VERIFICATION - COMPLETE PIPELINE CHECK (2015-2026)")
print("="*80)

# 1. Check df_gen structure
print(f"\n1. GENERATION DATA:")
print(f"   Shape: {df_gen.shape}")
gen_cols_normalized = [col for col in df_gen.columns if col.startswith('gen_')]
print(f"   Columns with 'gen_' prefix: {len(gen_cols_normalized)}/{len(df_gen.columns)}")
print(f"   All columns normalized: {len(gen_cols_normalized) == len(df_gen.columns)}")
print(f"   Sample columns: {list(df_gen.columns[:5])}")
print(f"   Date range: {df_gen.index.min()} to {df_gen.index.max()}")

# 2. Check zone transition (DE_AT_LU, DE_LU, AT)
print(f"\n2. ZONE TRANSITION CHECK:")

# DE_AT_LU
deatlu_cols = [col for col in df_gen.columns if col.endswith('_DE_AT_LU')]
print(f"\n   DE_AT_LU (should have 2015-2018 data):")
print(f"     Total columns: {len(deatlu_cols)}")
if deatlu_cols:
    deatlu_2015 = df_gen.loc[df_gen.index.year == 2015, deatlu_cols].sum().sum() if (df_gen.index.year == 2015).any() else 0
    deatlu_2019 = df_gen.loc[df_gen.index.year == 2019, deatlu_cols].sum().sum() if (df_gen.index.year == 2019).any() else 0
    deatlu_2026 = df_gen.loc[df_gen.index.year == 2026, deatlu_cols].sum().sum() if (df_gen.index.year == 2026).any() else 0
    print(f"     2015: {deatlu_2015:,.0f} MWh {'' if deatlu_2015 > 0 else ''}")
    print(f"     2019: {deatlu_2019:,.0f} MWh {' (should be 0)' if deatlu_2019 == 0 else ''}")
    print(f"     2026: {deatlu_2026:,.0f} MWh {'(should be 0)' if deatlu_2026 == 0 else ''}")
    print(f"     Sample: {deatlu_cols[:3]}")

# DE_LU
delu_cols = [col for col in df_gen.columns if col.endswith('_DE_LU')]
print(f"\n   DE_LU (should have 2019+ data):")
print(f"     Total columns: {len(delu_cols)}")
if delu_cols:
    delu_2015 = df_gen.loc[df_gen.index.year == 2015, delu_cols].sum().sum() if (df_gen.index.year == 2015).any() else 0
    delu_2019 = df_gen.loc[df_gen.index.year == 2019, delu_cols].sum().sum() if (df_gen.index.year == 2019).any() else 0
    delu_2026 = df_gen.loc[df_gen.index.year == 2026, delu_cols].sum().sum() if (df_gen.index.year == 2026).any() else 0
    print(f"     2015: {delu_2015:,.0f} MWh {'(should be 0)' if delu_2015 == 0 else ''}")
    print(f"     2019: {delu_2019:,.0f} MWh {'' if delu_2019 > 0 else ''}")
    print(f"     2026: {delu_2026:,.0f} MWh {' if delu_2026 > 0 else '— (no data yet)'}")
    print(f"     Sample: {delu_cols[:3]}")

# AT
at_cols = [col for col in df_gen.columns if col.endswith('_AT') and 'DE_AT_LU' not in col]
print(f"\n   AT (should have 2019+ data):")
print(f"     Total columns: {len(at_cols)}")
if at_cols:
    at_2015 = df_gen.loc[df_gen.index.year == 2015, at_cols].sum().sum() if (df_gen.index.year == 2015).any() else 0
    at_2019 = df_gen.loc[df_gen.index.year == 2019, at_cols].sum().sum() if (df_gen.index.year == 2019).any() else 0
    at_2026 = df_gen.loc[df_gen.index.year == 2026, at_cols].sum().sum() if (df_gen.index.year == 2026).any() else 0
    print(f"     2015: {at_2015:,.0f} MWh {'(should be 0)' if at_2015 == 0 else ''}")
    print(f"     2019: {at_2019:,.0f} MWh {' if at_2019 > 0 else ''}")
    print(f"     2026: {at_2026:,.0f} MWh {'' if at_2026 > 0 else '— (no data yet)'}")
    print(f"     Sample: {at_cols[:3]}")

# 3. Check Austria specifically
print(f"\n3. AUSTRIA DATA (2019+):")
print(f"   Total columns: {len(at_cols)}")
print(f"   Has gen_gas_AT: {'gen_gas_AT' in at_cols}")
print(f"   Has gen_coal_derived_gas_AT: {'gen_coal_derived_gas_AT' in at_cols}")
print(f"   Has gen_nuclear_AT: {'gen_nuclear_AT' in at_cols}")

# 4. Check if Cell 4 patterns would work
print(f"\n4. TESTING CELL 4 PATTERN MATCHING:")

# Test gas for AT
test_result_at = find_matching_columns(df_gen, 'gas', 'AT')
print(f"\n   Gas columns found for AT:")
print(f"     OLD: {test_result_at['old']}")
print(f"     NEW: {test_result_at['new']}")

if len(test_result_at['old']) == 1 and test_result_at['old'][0] == 'gen_gas_AT':
    print(f"    CORRECT: Found only gen_gas_AT")
elif len(test_result_at['old']) == 0 and len(test_result_at['new']) == 0:
    print(f"     No matches - this is expected since columns already have 'gen_' prefix")
else:
    print(f"    UNEXPECTED: Found {len(test_result_at['old']) + len(test_result_at['new'])} columns")

# Test gas for DE_AT_LU
test_result_deatlu = find_matching_columns(df_gen, 'gas', 'DE_AT_LU')
print(f"\n   Gas columns found for DE_AT_LU:")
print(f"     OLD: {test_result_deatlu['old']}")
print(f"     NEW: {test_result_deatlu['new']}")

if 'gen_gas_DE_AT_LU' in test_result_deatlu['old']:
    print(f"    CORRECT: Found gen_gas_DE_AT_LU")

# Test gas for DE_LU
test_result_delu = find_matching_columns(df_gen, 'gas', 'DE_LU')
print(f"\n   Gas columns found for DE_LU:")
print(f"     OLD: {test_result_delu['old']}")
print(f"     NEW: {test_result_delu['new']}")

if 'gen_gas_DE_LU' in test_result_delu['old']:
    print(f"   CORRECT: Found gen_gas_DE_LU")

# 5. Check complete zones
print(f"\n5. COMPLETE ZONES:")
print(f"   Total zones with complete data: {len(complete_zones)}")
print(f"   Zones (first 15): {complete_zones[:15]}")

# Check if DE_AT_LU, DE_LU, AT are in complete zones
special_zones_status = []
for zone in ['DE_AT_LU', 'DE_LU', 'AT']:
    if zone in complete_zones:
        special_zones_status.append(f"{zone}")
    else:
        special_zones_status.append(f"— {zone}")
print(f"   Special zones: {', '.join(special_zones_status)}")

# 6. Check prices
print(f"\n6. PRICE DATA:")
print(f"   Shape: {df_prices.shape}")
print(f"   Date range: {df_prices.index.min()} to {df_prices.index.max()}")

# Test price columns for all three zones
for zone in ['DE_AT_LU', 'DE_LU', 'AT']:
    price_col = find_price_column(df_prices, zone)
    if price_col:
        mean_price = df_prices[price_col].mean()
        print(f"   {zone}: column='{price_col}', mean={mean_price:.2f} EUR/MWh")
    else:
        print(f"   {zone}:  No price column found")

# Test PL price
pl_price_col = find_price_column(df_prices, 'PL')
if pl_price_col:
    pl_mean = df_prices[pl_price_col].mean()
    print(f"   PL: column='{pl_price_col}', mean={pl_mean:.2f} EUR/MWh (should be 80-100)")
else:
    print(f"   PL:  No price column found")

# 7. Check load data
print(f"\n7. LOAD DATA:")
print(f"   Shape: {df_load.shape}")
print(f"   Date range: {df_load.index.min()} to {df_load.index.max()}")

# Test load columns for all three zones
for zone in ['DE_AT_LU', 'DE_LU', 'AT']:
    load_col = find_load_column(df_load, zone)
    if load_col:
        mean_load = df_load[load_col].mean()
        print(f"   {zone}: column='{load_col}', mean={mean_load:.2f} MW")
    else:
        print(f"   {zone}:  No load column found")

# 8. Year-by-year summary
print(f"\n8. YEAR-BY-YEAR DATA SUMMARY:")

# UPDATED: Include 2026
years_to_check = [2015, 2018, 2019, 2025, 2026]
for year in years_to_check:
    mask = df_gen.index.year == year
    if mask.sum() > 0:
        total_gen = df_gen.loc[mask, gen_cols_normalized].sum().sum()
        print(f"   {year}: {mask.sum():>5,} rows, {total_gen:>15,.0f} MWh")
    else:
        print(f"   {year}: — No data")

print("\n" + "="*80)
print("VERIFICATION COMPLETE")
print("="*80)

# Final status
issues = []
if len(gen_cols_normalized) != len(df_gen.columns):
    issues.append("Some generation columns not normalized")
if 'DE_AT_LU' not in [c.split('_')[-1] for c in deatlu_cols] if deatlu_cols else True:
    issues.append("DE_AT_LU data missing")
if 'DE_LU' not in [c.split('_')[-1] for c in delu_cols] if delu_cols else True:
    issues.append("DE_LU data missing")
if len(at_cols) == 0:
    issues.append("AT data missing")

if issues:
    print(f"\n  Issues found:")
    for issue in issues:
        print(f"   • {issue}")
else:
    print(f"\n ALL CHECKS PASSED - Ready for analysis!")


# In[ ]:


# Cell 6a - USER COMMODITY DATA LOADER (no-Platts mode)
# When CFG.COMMODITY_SOURCE == 'file', reads USER_COMMODITY_FILE and builds the
# same three dataframes the Platts cells produce. EUR units throughout; see the
# template's readme sheet for the column contract.
COMMODITY_SOURCE = getattr(CFG, 'COMMODITY_SOURCE', 'platts').strip().lower()
if COMMODITY_SOURCE == 'file':
    _uf = CFG.paths['USER_COMMODITY_FILE']
    if not Path(_uf).exists():
        raise FileNotFoundError(f"COMMODITY_SOURCE='file' but USER_COMMODITY_FILE not found: {_uf}")
    def _read_user_sheet(sheet, required):
        df = pd.read_excel(_uf, sheet, header=1)          # row 1 is the note line
        df = df.dropna(subset=['Date'])
        df['Date'] = pd.to_datetime(df['Date'])
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"user commodity file, sheet '{sheet}': missing columns {missing}")
        return df.set_index('Date').sort_index()
    USER_GAS_HUBS = _read_user_sheet('gas_hubs', ['TTF'])
    USER_CO2      = _read_user_sheet('co2', ['EUA'])
    USER_COAL = _read_user_sheet('coal', ['coal_eur_t'])
    USER_OIL  = _read_user_sheet('oil',  ['oil_eur_t'])
    # Separate tabs may have different date grids -> outer union; NaN where a
    # commodity has no quote that day (same behaviour as blanks in one sheet).
    USER_COAL_OIL = USER_COAL[['coal_eur_t']].join(USER_OIL[['oil_eur_t']], how='outer')
    print(f"User commodity data loaded: gas hubs {list(USER_GAS_HUBS.columns)}, "
          f"co2 {len(USER_CO2):,} days, coal {USER_COAL['coal_eur_t'].notna().sum():,} days, "
          f"oil {USER_OIL['oil_eur_t'].notna().sum():,} days")
else:
    print("Commodity source: Platts API")


# In[ ]:


# %% CELL 6 - GAS PRICES (research Cell 6)
# %% ==========================================================================
# Cell 6: Fetch gas prices from Platts API with comprehensive hub mapping (2015-2026)
#  Updated to include DE_AT_LU zone and 2015-2018 period
#  Includes Iberian Exception on PVB
#  NOW EXTENDS TO 2026

import spgci as ci
import pandas as pd
import numpy as np

# ========== PLATTS API CREDENTIALS from control file (skipped in user-file mode) ==========
if COMMODITY_SOURCE != 'file':
    def _cfg_cred(name):
        v = str(getattr(CFG, name, '') or '').strip()
        return '' if v.lower() in ('', 'nan', 'none') else v

    PLATTS_USER = _cfg_cred('PLATTS_USER')
    PLATTS_PASS = _cfg_cred('PLATTS_PASS')
    PLATTS_KEY  = _cfg_cred('PLATTS_KEY')
    if not (PLATTS_USER and PLATTS_PASS and PLATTS_KEY):
        raise RuntimeError(
            "Platts credentials missing. Fill PLATTS_USER / PLATTS_PASS / PLATTS_KEY in the "
            "'config' tab of the control file, or run without Platts: set "
            "COMMODITY_SOURCE = 'file' in the control file and fill "
            "user_commodity_prices.xlsx (see cover sheet).")
    ci.set_credentials(PLATTS_USER, PLATTS_PASS, PLATTS_KEY)
    mdd = ci.MarketData()
    

def extract_gas_prices_for_zone(zone_suffix, hub_info, start_date='2015-01-01', end_date='2026-12-31'):
    """
    Extract gas hub prices for a zone using Platts API.
     Handles DE_AT_LU (2015-2018 old German hubs)
     Handles DE_LU (old hubs 2019 + THE from Oct 2021)
    NOW EXTENDS TO 2026
    
    Args:
        zone_suffix: Bidding zone code
        hub_info: Dict with 'hub' and 'symbol' keys
        start_date: Start date for extraction
        end_date: End date for extraction
    
    Returns:
        DataFrame with columns ['Date', 'Daily_Gas_Price_{zone}']
    """
    gas_hub = hub_info['hub']
    gas_symbol = hub_info['symbol']
    
    # ========== SPECIAL CASE: DE_AT_LU (2015-2018 only) ==========
    if zone_suffix == 'DE_AT_LU':
        excel_file = DATA_DIR / 'BO_Hub (Prices, Volumes) and Pipe (Prices).xlsx'
        
        if excel_file.exists():
            try:
                df = pd.read_excel(excel_file, sheet_name='all MS_D_t', header=1)
                df.columns = df.columns.str.strip()
                
                date_col = '7.Full Date(Hub Prices_1)'
                df['date'] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
                
                #  Filter 2015-01-01 to 2018-09-30 (DE_AT_LU period)
                start = pd.Timestamp('2015-01-01')
                end = pd.Timestamp('2018-09-30')
                mask = (df['date'] >= start) & (df['date'] <= end)
                df_range = df.loc[mask].copy()
                
                # Average of GasPool and NCG
                df_range['gas_price'] = df_range[['DE GASPOOL average price', 'DE NCG avg price']].mean(axis=1, skipna=True)
                gas_prices = df_range[['date', 'gas_price']].dropna(subset=['gas_price'])
                
                print(f"  ✓ DE_AT_LU: Loaded {len(gas_prices)} days (2015-2018, GasPool + NCG)")
                
            except Exception as e:
                print(f"   DE_AT_LU: Could not load from Excel: {e}")
                gas_prices = pd.DataFrame(columns=['date', 'gas_price'])
        else:
            print(f"   DE_AT_LU: Excel file not found at {excel_file}")
            gas_prices = pd.DataFrame(columns=['date', 'gas_price'])
    
    # ========== SPECIAL CASE: DE_LU (2019 onwards, including 2026) ==========
    elif zone_suffix == 'DE_LU':
        excel_file = DATA_DIR / 'BO_Hub (Prices, Volumes) and Pipe (Prices).xlsx'
        
        # Load old German hubs (2019-01-01 to 2021-09-29)
        if excel_file.exists():
            try:
                df = pd.read_excel(excel_file, sheet_name='all MS_D_t', header=1)
                df.columns = df.columns.str.strip()
                
                date_col = '7.Full Date(Hub Prices_1)'
                df['date'] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
                
                #  Filter 2019-01-01 to 2021-09-29
                start = pd.Timestamp('2019-01-01')
                end = pd.Timestamp('2021-09-29')
                mask = (df['date'] >= start) & (df['date'] <= end)
                df_range = df.loc[mask].copy()
                
                # Average of GasPool and NCG
                df_range['gas_price'] = df_range[['DE GASPOOL average price', 'DE NCG avg price']].mean(axis=1, skipna=True)
                old_hubs = df_range[['date', 'gas_price']].dropna(subset=['gas_price'])
                
                print(f"  ✓ DE_LU: Loaded {len(old_hubs)} days of old hubs (2019-Sep 2021)")
                
            except Exception as e:
                print(f"   DE_LU: Could not load old hubs: {e}")
                old_hubs = pd.DataFrame(columns=['date', 'gas_price'])
        else:
            print(f"   DE_LU: Excel file not found")
            old_hubs = pd.DataFrame(columns=['date', 'gas_price'])
        
        #  Fetch THE prices from Oct 2021 onwards (now includes 2026!)
        if SPGCI_AVAILABLE:
            try:
                mdd = spgci.MarketData()
                df_new = mdd.get_assessments_by_symbol_historical(
                    symbol=[gas_symbol], 
                    assess_date_gte='2021-10-01',
                    assess_date_lte=end_date,  #  Include up to end_date
                    bate='c', 
                    paginate=True
                )
                df_new = df_new.drop(['bate', 'isCorrected', 'modDate'], axis=1, errors='ignore')
                df_new = df_new.pivot_table(values='value', index='assessDate', columns='symbol').reset_index()
                df_new.rename(columns={'assessDate': 'date', gas_symbol: 'gas_price'}, inplace=True)
                df_new['date'] = pd.to_datetime(df_new['date'])
                
                print(f"  ✓ DE_LU: Fetched {len(df_new)} days of THE prices (Oct 2021 - {end_date})")
                
            except Exception as e:
                print(f"   DE_LU: Error fetching THE prices: {e}")
                df_new = pd.DataFrame(columns=['date', 'gas_price'])
        else:
            df_new = pd.DataFrame(columns=['date', 'gas_price'])
        
        # Concatenate old + new
        gas_prices = pd.concat([
            old_hubs[['date', 'gas_price']], 
            df_new[['date', 'gas_price']]
        ], ignore_index=True)
    
    # ========== STANDARD CASE: Fetch from Platts API (including 2026) ==========
    else:
        if SPGCI_AVAILABLE:
            try:
                mdd = spgci.MarketData()
                df_gas = mdd.get_assessments_by_symbol_historical(
                    symbol=[gas_symbol], 
                    assess_date_gte=start_date,
                    assess_date_lte=end_date,  #  Include up to end_date
                    bate='c', 
                    paginate=True
                )
                df_gas = df_gas.drop(['bate', 'isCorrected', 'modDate'], axis=1, errors='ignore')
                df_gas = df_gas.pivot_table(values='value', index='assessDate', columns='symbol').reset_index()
                df_gas.rename(columns={'assessDate': 'date', gas_symbol: 'gas_price'}, inplace=True)
                df_gas['date'] = pd.to_datetime(df_gas['date'])
                
                gas_prices = df_gas[['date', 'gas_price']].copy()
                
            except Exception as e:
                print(f"   {zone_suffix}: API error, trying TTF fallback: {e}")
                
                # Fallback to TTF
                try:
                    df_gas = mdd.get_assessments_by_symbol_historical(
                        symbol=['GTFTX00'], 
                        assess_date_gte=start_date,
                        assess_date_lte=end_date,  #  Include up to end_date
                        bate='c', 
                        paginate=True
                    )
                    df_gas = df_gas.drop(['bate', 'isCorrected', 'modDate'], axis=1, errors='ignore')
                    df_gas = df_gas.pivot_table(values='value', index='assessDate', columns='symbol').reset_index()
                    df_gas.rename(columns={'assessDate': 'date', 'GTFTX00': 'gas_price'}, inplace=True)
                    df_gas['date'] = pd.to_datetime(df_gas['date'])
                    
                    gas_prices = df_gas[['date', 'gas_price']].copy()
                    
                except Exception as e2:
                    print(f"  ✗ {zone_suffix}: TTF fallback also failed: {e2}")
                    return pd.DataFrame(columns=['Date', f'Daily_Gas_Price_{zone_suffix}'])
        else:
            # Generate fallback if API not available
            print(f"   {zone_suffix}: Platts API not available, using fallback prices")
            date_range = pd.date_range(start=start_date, end=end_date, freq='D')
            np.random.seed(hash(zone_suffix) % 2**32)
            base_price = 30.0
            variation = np.random.normal(0, 2, len(date_range))
            gas_prices = pd.DataFrame({
                'date': date_range,
                'gas_price': base_price + variation
            })
    
    # ========== FINALIZE ==========
    if gas_prices.empty:
        return pd.DataFrame(columns=['Date', f'Daily_Gas_Price_{zone_suffix}'])
    
    gas_prices['Date'] = gas_prices['date']
    gas_prices[f'Daily_Gas_Price_{zone_suffix}'] = gas_prices['gas_price']
    gas_prices = gas_prices[['Date', f'Daily_Gas_Price_{zone_suffix}']].dropna()
    
    if not gas_prices.empty:
        print(f"  ✓ {zone_suffix:10s} ({gas_hub:5s}): {len(gas_prices):,} days from {gas_prices['Date'].min().date()} to {gas_prices['Date'].max().date()}")
    
    return gas_prices


def get_zone_gas_price(gas_prices_df, zone):
    """
    Get gas price for a specific zone from the hub prices DataFrame.
    
    Args:
        gas_prices_df: DataFrame with hub gas prices
        zone: Bidding zone code
    
    Returns:
        Series with gas prices for the zone (including transport adder)
    """
    gas_col = f'Daily_Gas_Price_{zone}'
    
    if gas_col in gas_prices_df.columns:
        price = gas_prices_df[gas_col].copy()
    elif 'Daily_Gas_Price_NL' in gas_prices_df.columns:
        # Fallback to TTF (NL)
        price = gas_prices_df['Daily_Gas_Price_NL'].copy()
    else:
        raise RuntimeError(f"No gas hub price series for zone {zone} - "
                           f"check GAS_HUB_MAP / commodity inputs")
    
    # Add transport cost
    adder = GAS_TRANSPORT_ADDERS.get(zone, 0.0)
    price = price + adder
    
    return price

def apply_iberian_gas_cap(gas_price, date, zone):
    """
    Apply Iberian gas price cap if applicable.

    Args:
        gas_price: Original gas price (EUR/MWh)
        date: Timestamp of the price
        zone: Bidding zone ('ES' or 'PT')

    Returns:
        Capped gas price if within cap period, otherwise original price
    """
    if zone not in IBERIAN_GAS_CAP:
        return gas_price

    cap_config = IBERIAN_GAS_CAP[zone]

    # Check if date is within cap period
    if cap_config['start'] <= date <= cap_config['end']:
        month_key = date.strftime('%Y-%m')
        if month_key in cap_config['monthly_caps']:
            cap_value = cap_config['monthly_caps'][month_key]
            return min(gas_price, cap_value)

    return gas_price

# ========== MAIN EXECUTION ==========
print("=" * 80)
print("EXTRACTING GAS HUB PRICES FROM PLATTS API (2015-2026)")
print("=" * 80)

# Determine date range from loaded data if available
if not df_prices.empty:
    start_date = df_prices.index.min().strftime('%Y-%m-%d')
    end_date = df_prices.index.max().strftime('%Y-%m-%d')
else:
    start_date = '2015-01-01'
    end_date = '2026-12-31'

print(f"\nFetching gas prices for period: {start_date} to {end_date}")
print(f"SPGCI (Platts API) available: {SPGCI_AVAILABLE}")
print(f"\n Special handling:")
print(f"   • DE_AT_LU: Uses old German hubs (2015-2018)")
print(f"   • DE_LU: Uses old German hubs (2019) + THE (Oct 2021 - 2026)")
print()

# 20261012 modified to include gas prices from file
# Extract gas prices for all zones
gas_price_dfs = {}

if COMMODITY_SOURCE == 'file':
    # ---- user-supplied hub prices (no Platts) --------------------------------
    _HUBKEY = {'CZ': 'VTP_CZ', 'SK': 'VTP_SK'}          # two VTPs, disambiguated
    for zone, hub_info in GAS_HUB_MAP.items():
        try:
            if zone == 'DE_LU':                          # old German hubs -> THE splice
                the = USER_GAS_HUBS.get('THE')
                old = USER_GAS_HUBS.get('DE_OLD')
                if the is None and old is None:
                    raise ValueError("user file has neither 'THE' nor 'DE_OLD' column")
                if the is None:
                    s = old
                elif old is None:
                    s = the
                else:
                    s = the.where(USER_GAS_HUBS.index >= pd.Timestamp('2021-10-01'), old)
            elif zone == 'DE_AT_LU':
                s = USER_GAS_HUBS['DE_OLD'] if 'DE_OLD' in USER_GAS_HUBS.columns \
                    else USER_GAS_HUBS['TTF']
            else:
                key = _HUBKEY.get(zone, hub_info['hub'])
                if key in USER_GAS_HUBS.columns:
                    s = USER_GAS_HUBS[key]
                else:
                    print(f"   {zone}: hub '{key}' not in user file - falling back to TTF")
                    s = USER_GAS_HUBS['TTF']
            gas_price_dfs[zone] = (s.rename(f'Daily_Gas_Price_{zone}').dropna()
                                     .rename_axis('Date').reset_index())
            print(f"  ✓ {zone}: user hub data ({len(gas_price_dfs[zone]):,} days)")
        except Exception as e:
            print(f"  ✗ {zone}: FAILED - {e}")
else:
    # ---- Platts API (original path, unchanged) -------------------------------
    for zone, hub_info in GAS_HUB_MAP.items():
        try:
            gas_price_dfs[zone] = extract_gas_prices_for_zone(zone, hub_info, start_date, end_date)
        except Exception as e:
            print(f"  ✗ {zone}: FAILED - {e}")

# ========== MERGE ALL GAS PRICES ==========
print("\n" + "=" * 80)
print("MERGING ALL GAS HUB PRICES")
print("=" * 80)

gas_prices_df = None
for zone, df_gas in gas_price_dfs.items():
    if not df_gas.empty:
        if gas_prices_df is None:
            gas_prices_df = df_gas.copy()
        else:
            gas_prices_df = gas_prices_df.merge(df_gas, on='Date', how='outer')

if gas_prices_df is not None:
    gas_prices_df = gas_prices_df.sort_values('Date').reset_index(drop=True)
    gas_prices_df = gas_prices_df.set_index('Date')
    
    # ========================================
    # APPLY IBERIAN GAS PRICE CAP
    # ========================================
    print("\n" + "=" * 80)
    print("APPLYING IBERIAN GAS PRICE CAP (ES & PT: June 2022 - Dec 2023)")
    print("=" * 80)
    
    for zone in ['ES', 'PT']:
        gas_col = f'Daily_Gas_Price_{zone}'
        if gas_col in gas_prices_df.columns:
            # Apply cap to each date
            original_prices = gas_prices_df[gas_col].copy()
            gas_prices_df[gas_col] = gas_prices_df.apply(
                lambda row: apply_iberian_gas_cap(row[gas_col], row.name, zone) 
                if pd.notna(row[gas_col]) else row[gas_col],
                axis=1
            )
            
            # Count capped days
            capped_mask = (original_prices != gas_prices_df[gas_col]) & original_prices.notna()
            capped_count = capped_mask.sum()
            
            if capped_count > 0:
                avg_reduction = (original_prices[capped_mask] - gas_prices_df.loc[capped_mask, gas_col]).mean()
                
                print(f"  ✓ {zone}: Capped {capped_count:,} days")
                print(f"     Avg reduction: {avg_reduction:.2f} €/MWh")
                print(f"     Cap period: {IBERIAN_GAS_CAP[zone]['start'].date()} to {IBERIAN_GAS_CAP[zone]['end'].date()}")
            else:
                print(f"   {zone}: No prices exceeded cap during cap period")
    
    print(f"\n✓ Merged gas prices: {gas_prices_df.shape}")
    print(f"  Date range: {gas_prices_df.index.min()} to {gas_prices_df.index.max()}")
    print(f"  Zones with data: {len([c for c in gas_prices_df.columns if 'Daily_Gas_Price' in c])}")
    
    # ========== DIAGNOSTIC: SHOW COVERAGE ==========
    print("\n" + "=" * 80)
    print("GAS PRICE DATA COVERAGE BY ZONE (2015-2026)")
    print("=" * 80)
    
    for col in sorted(gas_prices_df.columns):
        zone = col.replace('Daily_Gas_Price_', '')
        hub = GAS_HUB_MAP.get(zone, {}).get('hub', 'Unknown')
        
        non_null = gas_prices_df[col].notna().sum()
        pct = non_null / len(gas_prices_df) * 100 if len(gas_prices_df) > 0 else 0
        
        cap_note = " ( CAP APPLIED)" if zone in ['ES', 'PT'] else ""
        zone_note = " (2015-2018 only)" if zone == 'DE_AT_LU' else " (2019+ only)" if zone in ['DE_LU', 'AT'] else ""
        
        print(f"  {zone:10s} ({hub:7s}): {non_null:>6,} / {len(gas_prices_df):,} days ({pct:>5.1f}%){cap_note}{zone_note}")
    
    print("\n✓ Gas hub prices extracted successfully (2015-2026)")
    print("  - Native hubs: CEGH, ZTP, VTP (CZ), VTP (SK), TTF, PEG, PSV, THE, PVB, NBP")
    print("  - DE_AT_LU: Old German hubs (GasPool + NCG, 2015-2018)")
    print("  - DE_LU: Old German hubs (2019) + THE (Oct 2021 - 2026)")
    print("  - IE_SEM uses UK NBP (symbol: AALJI00)")
    print("  - TTF proxy for: Nordic, Baltic, Eastern Europe, Balkans")
    print("  -  IBERIAN GAS CAP applied to ES & PT (June 2022 - December 2023)")
    
else:
    print("✗ No gas price data could be loaded")
    gas_prices_df = pd.DataFrame()

print("=" * 80)

# ========================================
# EXPORT GAS PRICES FOR SENSITIVITY ANALYSIS
# ========================================

print(f"\n💾 Exporting gas prices for sensitivity analysis...")

gas_export_file = OUTPUT_DIR / 'gas_prices_2015_2026.csv'
gas_prices_df.to_csv(gas_export_file)

print(f"   ✓ Saved: {gas_export_file}")
print(f"   Rows: {len(gas_prices_df):,}")



# In[ ]:


# %% CELL 7 - EUA PRICES (research Cell 7)
# %% ==========================================================================
# Cell 7: Fetch EU ETS carbon prices from Platts API (2015-2026)
# Updated to fetch from 2015 onwards
# NOW EXTENDS TO 2026

import spgci as ci
import pandas as pd
import numpy as np

# ========== PLATTS API (skipped entirely in user-file mode) ==========
if COMMODITY_SOURCE != 'file':
    import spgci as ci
    ci.set_credentials(PLATTS_USER, PLATTS_PASS, PLATTS_KEY)   # imported in Cell 6

# [CFG-OWNED, def removed]
# EU ETS Carbon Price Configuration
# CO2_SYMBOLS = {
#     'EUA': 'EUAEA00',  # EU Allowance (2015+)
#     'UKA': 'AIDUK00',  # UK Allowance (post-Brexit, Jan 2021+) in EUR/Mt
# }

def fetch_co2_prices_platts(start_date, end_date):
    """
    Fetch EU ETS carbon prices from Platts API.

    FIXED (ported from colleague, 20/08/2026): each scheme is fetched in its
    own try/except and UKA is skipped entirely before Jan 2021 (it did not
    trade). Previously a UKA failure propagated to an outer handler that
    discarded the EUA data which had already loaded successfully ->
    empty co2_prices_df -> 80 EUR/t default in Cell 10 -> no fossil
    price-setting in 2015-2018.
    """
    if not SPGCI_AVAILABLE:
        print(" Platts API not available - please install spgci package")
        return pd.DataFrame()

    try:
        mdd = ci.MarketData()
    except Exception as e:
        print(f"✗ Could not open Platts MarketData session: {e}")
        return pd.DataFrame()

    co2_prices = {}

    for scheme, symbol in CO2_SYMBOLS.items():
        # UKA only exists from Jan 2021 - skip it for earlier periods
        if scheme == 'UKA' and pd.Timestamp(end_date) < pd.Timestamp('2021-01-01'):
            print(f"  Skipping {scheme}: period ends before Jan 2021 (UKA not traded yet)")
            continue

        # UKA: never request dates before it existed
        scheme_start = start_date
        if scheme == 'UKA':
            scheme_start = max(pd.Timestamp(start_date),
                               pd.Timestamp('2021-01-01')).strftime('%Y-%m-%d')

        print(f"  Fetching {scheme} prices ({symbol}) from {scheme_start}...")

        try:
            df_co2 = mdd.get_assessments_by_symbol_historical(
                symbol=[symbol],
                assess_date_gte=scheme_start,
                assess_date_lte=end_date,
                bate='c',
                paginate=True
            )

            if df_co2 is None or df_co2.empty:
                print(f"   No data returned for {scheme} - skipping")
                continue

            if 'value' not in df_co2.columns:
                print(f"   {scheme}: response has no 'value' column "
                      f"(cols={list(df_co2.columns)[:6]}) - skipping")
                continue

            df_co2 = df_co2.drop(['bate', 'isCorrected', 'modDate'], axis=1, errors='ignore')
            df_co2 = df_co2.pivot_table(values='value', index='assessDate', columns='symbol').reset_index()
            df_co2.rename(columns={'assessDate': 'date', symbol: scheme}, inplace=True)
            df_co2['date'] = pd.to_datetime(df_co2['date'])

            co2_prices[scheme] = df_co2.set_index('date')[scheme]

            print(f"  ✓ {scheme}: {len(df_co2)} days from "
                  f"{df_co2['date'].min().date()} to {df_co2['date'].max().date()}")

        except Exception as e:
            print(f"   {scheme} fetch failed ({str(e)[:90]}) - continuing without it")
            continue

    if not co2_prices:
        print("  ✗ No CO2 price data could be loaded")
        return pd.DataFrame()

    df = pd.DataFrame(co2_prices)

    # Keep your pre-2021 UKA <- EUA backfill (more sensible for GB/IE than the
    # bfill that would otherwise pull the first 2021 UKA value back to 2015)
    if 'UKA' in df.columns and 'EUA' in df.columns:
        first_uka = df['UKA'].first_valid_index()
        if first_uka is not None:
            df.loc[:first_uka, 'UKA'] = df.loc[:first_uka, 'UKA'].fillna(
                df.loc[:first_uka, 'EUA'])
            print(f"   Filled UKA with EUA for period before {first_uka.date()}")
    elif 'EUA' in df.columns and 'UKA' not in df.columns:
        df['UKA'] = df['EUA'].copy()
        print("   Created UKA column using EUA values (UKA not available)")

    print(f"  ✓ CO2 frame built: {df.shape[0]:,} days x {list(df.columns)}")
    return df
        

def get_zone_co2_price(co2_prices_df, zone):
    """
    Get CO2 price for a specific zone.
     Handles UK zones (GB, IE_SEM) with UKA from 2021 onwards
    Uses EUA for all other zones and UK before 2021
    
    Args:
        co2_prices_df: DataFrame with CO2 prices
        zone: Bidding zone code
    
    Returns:
        Series with CO2 prices for the zone
    """
    if co2_prices_df.empty:
        raise RuntimeError("No CO2 price data available - check EUA download/inputs")
    
    # UK uses UKA from Jan 2021, rest of Europe uses EUA
    if zone in ['GB', 'IE_SEM']:
        scheme = 'UKA'
    else:
        scheme = 'EUA'
    
    if scheme in co2_prices_df.columns:
        return co2_prices_df[scheme].fillna(method='ffill')  # Forward fill gaps
    elif 'EUA' in co2_prices_df.columns:
        # Fallback to EUA if UKA not available
        print(f"  ℹ {zone}: Using EUA price (UKA not available)")
        return co2_prices_df['EUA'].fillna(method='ffill')
    else:
        raise RuntimeError(f"{zone}: no CO2 price data - check EUA download/inputs")

# ========== MAIN EXECUTION ==========
print("=" * 80)
print("FETCHING CO2 PRICES FROM PLATTS API (2015-2026)")
print("=" * 80)

# Use date range from electricity prices if available
if 'df_prices' in globals() and not df_prices.empty:
    start_date = df_prices.index.min().strftime('%Y-%m-%d')
    end_date = df_prices.index.max().strftime('%Y-%m-%d')
else:
    start_date = '2015-01-01'
    end_date = '2026-12-31'

print(f"\nFetching CO2 prices for period: {start_date} to {end_date}")
print(f"SPGCI (Platts API) available: {SPGCI_AVAILABLE}")
print(f"\nℹ Note: UKA (UK Allowance) only available from Jan 2021 onwards")
print(f"         Using EUA for UK zones before Jan 2021\n")

# 20261012 Modified to include Co2 prices from file
if COMMODITY_SOURCE == 'file':
    co2_prices_df = USER_CO2.copy()
    # Mirror the Platts path's UKA handling: GB/IE_SEM read 'UKA', so the column
    # must exist and be EUA-backfilled before UKA trading started (Jan 2021).
    if 'UKA' not in co2_prices_df.columns or co2_prices_df['UKA'].notna().sum() == 0:
        co2_prices_df['UKA'] = co2_prices_df['EUA'].copy()
        print("  ℹ UKA empty/absent in user file - using EUA for UK zones")
    else:
        first_uka = co2_prices_df['UKA'].first_valid_index()
        co2_prices_df.loc[:first_uka, 'UKA'] = \
            co2_prices_df.loc[:first_uka, 'UKA'].fillna(co2_prices_df.loc[:first_uka, 'EUA'])
        print(f"  ℹFilled UKA with EUA before {first_uka.date()}")
    print(f" CO2 prices from user file: {len(co2_prices_df):,} days x {list(co2_prices_df.columns)}")
else:
    co2_prices_df = fetch_co2_prices_platts(start_date, end_date)

if not co2_prices_df.empty:
    print("\n" + "=" * 80)
    print("CO2 PRICE DATA SUMMARY (2015-2026)")
    print("=" * 80)
    print(f"Date range: {co2_prices_df.index.min()} to {co2_prices_df.index.max()}")
    print(f"Total days: {len(co2_prices_df)}")
    
    if 'EUA' in co2_prices_df.columns:
        eua_data = co2_prices_df['EUA'].dropna()
        print(f"\nEUA Statistics:")
        print(f"  Average: {eua_data.mean():.2f} €/tCO2")
        print(f"  Min: {eua_data.min():.2f} €/tCO2 on {eua_data.idxmin().date()}")
        print(f"  Max: {eua_data.max():.2f} €/tCO2 on {eua_data.idxmax().date()}")
        print(f"  Non-null values: {eua_data.notna().sum():,} / {len(co2_prices_df):,}")
        print(f"  Coverage: {eua_data.notna().sum() / len(co2_prices_df) * 100:.1f}%")
    
    if 'UKA' in co2_prices_df.columns:
        uka_data = co2_prices_df['UKA'].dropna()
        if len(uka_data) > 0:
            print(f"\nUKA Statistics (Jan 2021+):")
            print(f"  Average: {uka_data.mean():.2f} €/tCO2")
            print(f"  Min: {uka_data.min():.2f} €/tCO2 on {uka_data.idxmin().date()}")
            print(f"  Max: {uka_data.max():.2f} €/tCO2 on {uka_data.idxmax().date()}")
            print(f"  Non-null values: {uka_data.notna().sum():,} / {len(co2_prices_df):,}")
            print(f"  First UKA date: {uka_data.index.min().date()}")
            print(f"  Last UKA date: {uka_data.index.max().date()}")
        else:
            print(f"\n UKA column exists but has no data")
    
    # Year-by-year summary
    print(f"\n{'='*60}")
    print("YEAR-BY-YEAR EUA PRICES")
    print(f"{'='*60}")
    
    if 'EUA' in co2_prices_df.columns:
        co2_prices_df['year'] = co2_prices_df.index.year
        
        for year in sorted(co2_prices_df['year'].unique()):
            year_data = co2_prices_df[co2_prices_df['year'] == year]['EUA'].dropna()
            if len(year_data) > 0:
                print(f"{year}: {year_data.mean():>6.2f} €/tCO2 avg  |  "
                      f"{year_data.min():>6.2f} - {year_data.max():>6.2f} range  |  "
                      f"{len(year_data):>3} days")
        
        co2_prices_df.drop('year', axis=1, inplace=True)
    
    print("\n✓ CO2 prices loaded successfully (2015-2026)")
    print("  - EUA: EU Allowances for all EU zones")
    print("  - UKA: UK Allowances for GB and IE_SEM (from Jan 2021)")
    print("  - UK zones use EUA before Jan 2021")
    
else:
    print("\n✗ Failed to load CO2 prices from Platts API")
    print("  Please check API credentials and symbol codes")

print("=" * 80)

# ========================================
# EXPORT CO2 PRICES FOR SENSITIVITY ANALYSIS
# ========================================

print(f"\n💾 Exporting CO2 prices for sensitivity analysis...")

co2_export_file = OUTPUT_DIR / 'co2_prices_2015_2026.csv'
co2_prices_df.to_csv(co2_export_file)

print(f"   ✓ Saved: {co2_export_file}")
print(f"   Rows: {len(co2_prices_df):,}")


# In[ ]:


# %% CELL 8 - COAL/OIL PRICES (research Cell 8)
# %% ==========================================================================
# Cell 8: Fetch coal and oil prices from Platts API and convert using ECB FX rates (2015-2026)
#  Updated to fetch from 2015 onwards
#  NOW EXTENDS TO 2026

import spgci as ci
import pandas as pd
import numpy as np
import requests
import io
from datetime import datetime

# ========== PLATTS API (skipped entirely in user-file mode) ==========
if COMMODITY_SOURCE != 'file':
    import spgci as ci
    ci.set_credentials(PLATTS_USER, PLATTS_PASS, PLATTS_KEY)   # imported in Cell 6



def fetch_usd_eur_rates(start_date, end_date):
    """
    Fetch USD/EUR exchange rates from ECB API.
     Handles full 2015-2026 period
    
    Args:
        start_date: Start date (YYYY-MM-DD)
        end_date: End date (YYYY-MM-DD)
    
    Returns:
        DataFrame with Date and USD_EUR columns
    """
    print("\n" + "=" * 80)
    print("EXTRACTING USD/EUR EXCHANGE RATES (ECB) - 2015-2026")
    print("=" * 80)
    
    try:
        # ECB API
        entrypoint = "https://data-api.ecb.europa.eu/service/"
        key_USD = "D.USD.EUR.SP00.A"
        params = {"startPeriod": start_date, "endPeriod": end_date}
        
        resp = requests.get(
            f"{entrypoint}data/EXR/{key_USD}", 
            params=params, 
            headers={"Accept": "text/csv"}
        )
        
        df_fx = pd.read_csv(io.StringIO(resp.text))
        df_fx['TIME_PERIOD'] = pd.to_datetime(df_fx['TIME_PERIOD'])
        USD = df_fx.rename(columns={'TIME_PERIOD': 'Date', 'OBS_VALUE': 'USD_EUR'})[['Date', 'USD_EUR']]
        
        print(f"✓ USD/EUR rates: {len(USD):,} days from {USD['Date'].min().date()} to {USD['Date'].max().date()}")
        
        # Show year-by-year average
        USD['year'] = USD['Date'].dt.year
        print(f"\n  Year-by-year average USD/EUR rates:")
        for year in sorted(USD['year'].unique()):
            year_avg = USD[USD['year'] == year]['USD_EUR'].mean()
            print(f"    {year}: {year_avg:.4f}")
        USD.drop('year', axis=1, inplace=True)
        
        print("=" * 80)
        
        return USD
        
    except Exception as e:
        print(f"✗ Error fetching USD/EUR rates: {e}")
        print("  Using default rate of 1.10 (average 2015-2026)")
        print("=" * 80)
        return None


def fetch_commodity_prices_platts(start_date, end_date):
    """
    Fetch coal and oil prices from Platts API.
    Handles full 2015-2026 period
    
    Args:
        start_date: Start date for data (YYYY-MM-DD)
        end_date: End date for data (YYYY-MM-DD)
    
    Returns:
        DataFrame with commodity prices (coal in EUR/t, fuel oil in $/t)
    """
    if not SPGCI_AVAILABLE:
        print(" Platts API not available - please install spgci package")
        return pd.DataFrame()
    
    try:
        mdd = ci.MarketData()
        
        prices = {}
        
        for commodity, config in COMMODITY_SYMBOLS.items():
            symbol = config['symbol']
            bate = config['bate']
            name = config['name']
            currency = config['currency']
            
            print(f"  Fetching {name} ({symbol}, bate={bate}, {currency})...")
            
            # Use get_assessments_by_symbol_historical with specific bate
            df_commodity = mdd.get_assessments_by_symbol_historical(
                symbol=[symbol],
                assess_date_gte=start_date,
                assess_date_lte=end_date,
                bate=bate,  # Use commodity-specific bate ('u' for coal, 'c' for oil)
                paginate=True
            )
            
            if df_commodity.empty:
                print(f"   No data returned for {commodity}")
                continue
            
            # Clean and pivot data
            df_commodity = df_commodity.drop(['bate', 'isCorrected', 'modDate'], axis=1, errors='ignore')
            df_commodity = df_commodity.pivot_table(values='value', index='assessDate', columns='symbol').reset_index()
            df_commodity.rename(columns={'assessDate': 'date', symbol: commodity}, inplace=True)
            df_commodity['date'] = pd.to_datetime(df_commodity['date'])
            
            prices[commodity] = df_commodity.set_index('date')[commodity]
            
            print(f"  ✓ {commodity}: {len(df_commodity)} days from {df_commodity['date'].min().date()} to {df_commodity['date'].max().date()}")
        
        if not prices:
            print("  ✗ No commodity price data could be loaded")
            return pd.DataFrame()
        
        # Combine into single DataFrame
        df = pd.DataFrame(prices)
        
        return df
        
    except Exception as e:
        print(f"✗ Error fetching Platts commodity prices: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()


def convert_usd_to_eur(usd_tonne, fx_df):
    """
    Convert commodity price from USD/tonne to EUR/tonne using ECB FX rates.
    
    ECB provides USD/EUR rate as "How many USD = 1 EUR"
    Example: USD_EUR = 1.10 means 1 EUR = 1.10 USD
    
    To convert USD → EUR: DIVIDE by USD_EUR
    Example: $110/t ÷ 1.10 = €100/t
    
    Args:
        usd_tonne: Series with commodity price in USD/tonne
        fx_df: DataFrame with Date and USD_EUR columns
    
    Returns:
        Series with commodity price in EUR/tonne
    """
    # Merge FX rates
    commodity_df = pd.DataFrame({'usd_price': usd_tonne})
    commodity_df['Date'] = commodity_df.index
    
    commodity_df = commodity_df.merge(fx_df, on='Date', how='left')
    
    # Forward fill missing FX rates
    commodity_df['USD_EUR'] = commodity_df['USD_EUR'].fillna(method='ffill')
    
    # If still missing, use historical average (1 EUR ≈ 1.10 USD for 2015-2026)
    commodity_df['USD_EUR'] = commodity_df['USD_EUR'].fillna(1.10)
    
    #  CORRECTED: Divide by USD_EUR to convert USD → EUR
    # ECB rate means: 1 EUR = USD_EUR dollars
    # So: $X / USD_EUR = €Y
    eur_tonne = commodity_df['usd_price'] / commodity_df['USD_EUR']
    
    # Set the index back to the original index
    eur_tonne.index = usd_tonne.index
    
    return eur_tonne


# ========== MAIN EXECUTION ==========
print("=" * 80)
print("FETCHING COAL & OIL PRICES FROM PLATTS API (2015-2026)")
print("=" * 80)

# Use date range from electricity prices if available
if 'df_prices' in globals() and not df_prices.empty:
    start_date = df_prices.index.min().strftime('%Y-%m-%d')
    end_date = df_prices.index.max().strftime('%Y-%m-%d')
else:
    start_date = '2015-01-01'
    end_date = '2026-12-31'

print(f"\nFetching coal and oil prices for period: {start_date} to {end_date}")
print(f"SPGCI (Platts API) available: {SPGCI_AVAILABLE}\n")

# 20261012 Modified so it can work with commodity prices from file
if COMMODITY_SOURCE == 'file':
    # User-supplied prices, already EUR/t: no Platts, no FX conversion
    fx_rates = None
    commodity_prices_df = pd.DataFrame(index=USER_COAL_OIL.index)
    commodity_prices_df['coal_ara']   = USER_COAL_OIL['coal_eur_t']
    commodity_prices_df['coal_price'] = USER_COAL_OIL['coal_eur_t']
    commodity_prices_df['fuel_oil']   = USER_COAL_OIL['oil_eur_t']   # EUR here, not USD
    commodity_prices_df['oil_price']  = USER_COAL_OIL['oil_eur_t']
    print(f"Commodity prices from user file: {len(commodity_prices_df):,} days (EUR/t, no FX applied)")
else:
    # Fetch USD/EUR exchange rates from ECB (only needed for fuel oil)
    fx_rates = fetch_usd_eur_rates(start_date, end_date)

    # Fetch commodity prices from Platts
    commodity_prices_df = fetch_commodity_prices_platts(start_date, end_date)

if not commodity_prices_df.empty:
    print("\n" + "=" * 80)
    print("COMMODITY PRICE DATA SUMMARY (2015-2026)")
    print("=" * 80)
    
    
    # 20261012 Modified for use of commodities prices from file
    if COMMODITY_SOURCE != 'file':
        # Coal is ALREADY in EUR/tonne - no conversion needed
        if 'coal_ara' in commodity_prices_df.columns:
            commodity_prices_df['coal_price'] = commodity_prices_df['coal_ara']
            print("✓ Coal prices already in EUR/tonne (CSEUR00 symbol)")

        # Fuel oil conversion
        if 'fuel_oil' in commodity_prices_df.columns and fx_rates is not None:
            commodity_prices_df['oil_price'] = convert_usd_to_eur(
                commodity_prices_df['fuel_oil'],
                fx_rates
            )
            print("✓ Fuel oil converted from USD/tonne to EUR/tonne using ECB exchange rates")
        elif 'fuel_oil' in commodity_prices_df.columns:
            # CORRECTED: Divide by default rate
            print("   Using default FX rate (1 EUR = 1.10 USD) for fuel oil conversion")
            commodity_prices_df['oil_price'] = commodity_prices_df['fuel_oil'] / 1.10
    else:
        # COMMODITY_SOURCE == 'file': coal_price/oil_price already set in EUR by Branch C
        print("✓ User-supplied prices already EUR/t - conversion skipped")
    
    print(f"\nDate range: {commodity_prices_df.index.min()} to {commodity_prices_df.index.max()}")
    print(f"Total days: {len(commodity_prices_df)}")
    
    if 'coal_ara' in commodity_prices_df.columns:
        coal_data = commodity_prices_df['coal_ara'].dropna()
        print(f"\nCoal ARA (Daily Average) Statistics:")
        print(f"  Symbol: CSEUR00 (bate=u) - Already in EUR")
        print(f"  Average: {coal_data.mean():.2f} €/t")
        print(f"  Min: {coal_data.min():.2f} €/t on {coal_data.idxmin().date()}")
        print(f"  Max: {coal_data.max():.2f} €/t on {coal_data.idxmax().date()}")
        print(f"  Non-null values: {coal_data.notna().sum():,} / {len(commodity_prices_df):,}")
        print(f"  Coverage: {coal_data.notna().sum() / len(commodity_prices_df) * 100:.1f}%")
    
    if 'fuel_oil' in commodity_prices_df.columns:
        oil_data = commodity_prices_df['fuel_oil'].dropna()
        oil_eur_data = commodity_prices_df['oil_price'].dropna()
        print(f"\nFuel Oil 1%S (FOB Rotterdam) Statistics:")
        print(f"  Symbol: PUAAP00 (bate=c) - USD converted to EUR")
        print(f"  Average (USD): {oil_data.mean():.2f} $/t")
        print(f"  Min (USD): {oil_data.min():.2f} $/t on {oil_data.idxmin().date()}")
        print(f"  Max (USD): {oil_data.max():.2f} $/t on {oil_data.idxmax().date()}")
        print(f"  Average (EUR): {oil_eur_data.mean():.2f} €/t")
        print(f"  Non-null values: {oil_data.notna().sum():,} / {len(commodity_prices_df):,}")
        print(f"  Coverage: {oil_data.notna().sum() / len(commodity_prices_df) * 100:.1f}%")
    
    # Year-by-year summary
    print(f"\n{'='*80}")
    print("YEAR-BY-YEAR COMMODITY PRICES")
    print(f"{'='*80}")
    
    commodity_prices_df['year'] = commodity_prices_df.index.year
    
    print(f"\n{'Year':<6} | {'Coal (€/t)':>12} | {'Oil USD ($/t)':>15} | {'Oil EUR (€/t)':>15}")
    print(f"{'-'*60}")
    
    for year in sorted(commodity_prices_df['year'].unique()):
        year_data = commodity_prices_df[commodity_prices_df['year'] == year]
        
        coal_avg = year_data['coal_ara'].mean() if 'coal_ara' in year_data.columns else np.nan
        oil_usd_avg = year_data['fuel_oil'].mean() if 'fuel_oil' in year_data.columns else np.nan
        oil_eur_avg = year_data['oil_price'].mean() if 'oil_price' in year_data.columns else np.nan
        
        coal_str = f"{coal_avg:>12.2f}" if not pd.isna(coal_avg) else f"{'—':>12}"
        oil_usd_str = f"{oil_usd_avg:>15.2f}" if not pd.isna(oil_usd_avg) else f"{'—':>15}"
        oil_eur_str = f"{oil_eur_avg:>15.2f}" if not pd.isna(oil_eur_avg) else f"{'—':>15}"
        
        print(f"{year:<6} | {coal_str} | {oil_usd_str} | {oil_eur_str}")
    
    commodity_prices_df.drop('year', axis=1, inplace=True)
    
    print("\n✓ Coal and oil prices loaded successfully (2015-2026)")
    print("  - CSEUR00: Coal CIF ARA in EUR/tonne (no conversion needed)")
    print("  - PUAAP00: Fuel Oil 1%S in USD/tonne → converted to EUR/tonne")
    
else:
    print("\n✗ Failed to load commodity prices from Platts API")
    print("  Please check API credentials and symbol codes")

print("=" * 80)

# ========== VERIFICATION CHECK ==========
if COMMODITY_SOURCE != 'file' and not commodity_prices_df.empty and 'fuel_oil' in commodity_prices_df.columns and 'oil_price' in commodity_prices_df.columns:
    print("\n" + "=" * 80)
    print("CONVERSION VERIFICATION CHECK")
    print("=" * 80)
    
    # Sample a few dates to check conversion
    sample_dates = commodity_prices_df.sample(min(5, len(commodity_prices_df)))
    
    print(f"\n{'Date':<12s} | {'USD/t':>10s} | {'USD_EUR':>10s} | {'EUR/t':>10s} | {'Check':>15s}")
    print("-" * 80)
    
    for idx, row in sample_dates.iterrows():
        usd_price = row['fuel_oil']
        eur_price = row['oil_price']
        
        # Get FX rate for this date
        if fx_rates is not None:
            fx_row = fx_rates[fx_rates['Date'] == idx]
            if len(fx_row) > 0:
                fx_rate = fx_row['USD_EUR'].iloc[0]
            else:
                fx_rate = 1.10  # fallback
        else:
            fx_rate = 1.10
        
        # Calculate expected EUR price
        expected_eur = usd_price / fx_rate
        
        # Check if conversion is correct
        is_correct = abs(eur_price - expected_eur) < 0.01
        status = "✓ OK" if is_correct else "✗ ERROR"
        
        print(f"{idx.date()!s:<12s} | {usd_price:>10.2f} | {fx_rate:>10.4f} | {eur_price:>10.2f} | {status:>15s}")
    
    # Overall sanity check
    avg_usd = commodity_prices_df['fuel_oil'].mean()
    avg_eur = commodity_prices_df['oil_price'].mean()
    implied_rate = avg_usd / avg_eur
    
    print("\n" + "=" * 80)
    print("SANITY CHECK:")
    print(f"  Average USD price: {avg_usd:.2f} $/t")
    print(f"  Average EUR price: {avg_eur:.2f} €/t")
    print(f"  Implied FX rate: {implied_rate:.4f} (should be ~1.05-1.20 for 2015-2026)")
    
    if 1.0 < implied_rate < 1.25:
        print("  ✓ Conversion looks correct!")
    else:
        print("  ✗ WARNING: Implied FX rate is unusual - check conversion logic")
    
    print("=" * 80)

    # ========================================
# EXPORT COMMODITY PRICES FOR SENSITIVITY ANALYSIS
# ========================================

print(f"\n💾 Exporting commodity prices for sensitivity analysis...")

commodity_export_file = OUTPUT_DIR / 'commodity_prices_2015_2026.csv'
commodity_prices_df.to_csv(commodity_export_file)

print(f"   ✓ Saved: {commodity_export_file}")
print(f"   Rows: {len(commodity_prices_df):,}")
print(f"   Columns: {', '.join(commodity_prices_df.columns.tolist())}")


# In[18]:


commodity_prices_df


# In[ ]:


# %% ==========================================================================
# %% CELL 9 - RESERVOIR DATA (research Cell 9)
# %% ==========================================================================
# Cell 9: Load Reservoir Data from ENTSO-E (2015-2026)
#  Updated to load both 2015-2018 and 2019-2026 files
#  Handles DE_AT_LU zone transition
#  NOW EXTENDS TO 2026 (included in main 2019-2026 file)

import pandas as pd
import numpy as np
from pathlib import Path

print("=" * 80)
print("LOADING RESERVOIR DATA FROM ENTSO-E (2015-2026)")
print("=" * 80)

# ========== PATH CONFIGURATION ==========
RESERVOIR_DATA_FILE = CFG.paths['reservoir']   # <- control file, full period


def load_reservoir_data_file(reservoir_file, period_name):
    """
    Load weekly reservoir data from a single ENTSO-E CSV file.
    
    Args:
        reservoir_file: Path to CSV file
        period_name: String describing the period (e.g., "2015-2018")
    
    Returns:
        dict: {zone: DataFrame with hourly 'filling_pct' column (UTC-indexed)}
    """
    if not Path(reservoir_file).exists():
        print(f" {period_name} file not found: {reservoir_file}")
        return {}
    
    try:
        print(f"\n📂 Loading {period_name}: {reservoir_file.name}")
        
        # ========== LOAD CSV ==========
        df = pd.read_csv(reservoir_file)
        
        # Validate columns
        expected_cols = {'Date', 'Country', 'ResVol_MWh'}
        if not expected_cols.issubset(df.columns):
            print(f" Invalid CSV format! Expected: {expected_cols}, Found: {set(df.columns)}")
            return {}
        
        # ========== PARSE DATES ==========
        df['Date'] = pd.to_datetime(df['Date'], utc=False)
        
        # Localize to CET/CEST (handles DST automatically)
        df['Date'] = df['Date'].dt.tz_localize('Europe/Vienna', ambiguous='infer', nonexistent='shift_forward')
        
        # Convert to UTC (standardizes all zones)
        df['Date'] = df['Date'].dt.tz_convert('UTC')
        
        # Drop invalid dates
        df = df.dropna(subset=['Date'])
        
        if df.empty:
            print(" No valid data after date parsing")
            return {}
        
        print(f"   Loaded {len(df):,} rows covering {df['Country'].nunique()} zones")
        print(f"   Date range: {df['Date'].min()} to {df['Date'].max()}")
        
        # ========== PROCESS EACH ZONE ==========
        reservoir_data = {}
        
        for zone in HYDRO_RESERVOIR_ZONES:
            zone_data = df[df['Country'] == zone].copy()
            
            if zone_data.empty:
                continue
            
            # ========== DATA CLEANING ==========
            zone_data = zone_data[zone_data['ResVol_MWh'].notna()]
            zone_data = zone_data[zone_data['ResVol_MWh'] >= 0]
            
            if zone_data.empty:
                continue
            
            # Calculate max volume (for normalization to %)
            max_vol = zone_data['ResVol_MWh'].max()
            
            if max_vol <= 0:
                continue
            
            # Calculate filling percentage
            zone_data['filling_pct'] = (zone_data['ResVol_MWh'] / max_vol) * 100
            
            # ========== NORMALIZE TO WEEK START ==========
            # Step 1: Convert to period, then to timestamp (timezone-naive)
            zone_data['week_start'] = zone_data['Date'].dt.to_period('W-MON').dt.to_timestamp()
            
            # Step 2: Localize back to UTC (restore timezone)
            zone_data['week_start'] = zone_data['week_start'].dt.tz_localize('UTC')
            
            # Remove duplicates (keep first entry per week)
            zone_data = zone_data.drop_duplicates(subset=['week_start'], keep='first')
            
            # Sort and set index
            zone_data = zone_data.sort_values('week_start').set_index('week_start')
            
            # ========== INTERPOLATE WEEKLY → HOURLY ==========
            date_start = zone_data.index.min()
            date_end = zone_data.index.max()
            
            # Create hourly index (UTC, timezone-aware)
            hourly_index = pd.date_range(
                start=date_start,
                end=date_end,
                freq='H',
                tz='UTC'
            )
            
            # Reindex to hourly
            hourly_data = zone_data[['filling_pct']].reindex(hourly_index)
            
            # ========== INTERPOLATION ==========
            # 1. Linear interpolation between weekly values
            hourly_data['filling_pct'] = hourly_data['filling_pct'].interpolate(
                method='linear',
                limit_direction='both'
            )
            
            # 2. Backward fill at start
            hourly_data['filling_pct'] = hourly_data['filling_pct'].fillna(method='bfill')
            
            # 3. Forward fill at end
            hourly_data['filling_pct'] = hourly_data['filling_pct'].fillna(method='ffill')
            
            # 4. Fill remaining with median (for big gaps)
            if hourly_data['filling_pct'].isna().any():
                median_fill = hourly_data['filling_pct'].median()
                hourly_data['filling_pct'] = hourly_data['filling_pct'].fillna(
                    median_fill if pd.notna(median_fill) else 50.0
                )
            
            # 5. Clip to valid range [0, 100]
            hourly_data['filling_pct'] = hourly_data['filling_pct'].clip(0, 100)
            
            # ========== STORE RESULT ==========
            reservoir_data[zone] = hourly_data
            
            # Calculate statistics
            gap_weeks = len(zone_data)
            gap_hours = len(hourly_data)
            mean_fill = hourly_data['filling_pct'].mean()
            min_fill = hourly_data['filling_pct'].min()
            max_fill = hourly_data['filling_pct'].max()
            std_fill = hourly_data['filling_pct'].std()
            
            print(f"  ✓ {zone:10s}: {gap_weeks:>4} weeks → {gap_hours:>7,} hours | "
                  f"Fill: {min_fill:>5.1f}%-{mean_fill:>5.1f}%-{max_fill:>5.1f}% (σ={std_fill:.1f})")
            
            # Warn about sparse data
            if gap_weeks < 26:
                print(f"       Only {gap_weeks} weeks of data (< 6 months)")
        
        return reservoir_data
        
    except Exception as e:
        print(f" ERROR loading {period_name}: {e}")
        import traceback
        traceback.print_exc()
        return {}


def combine_reservoir_data(data_2015_2018, data_2019_2026):
    """
    Combine reservoir data from both periods.
     Handles zone transition (DE_AT_LU → DE_LU)
     2019-2026 file now includes data through 2026
    
    Args:
        data_2015_2018: dict from 2015-2018 file
        data_2019_2026: dict from 2019-2026 file (includes 2026)
    
    Returns:
        dict: Combined reservoir data for all zones
    """
    print(f"\n{'='*80}")
    print("COMBINING 2015-2018 AND 2019-2026 DATA")
    print(f"{'='*80}")
    
    combined = {}
    
    # Get all unique zones from both periods
    all_zones = set(data_2015_2018.keys()) | set(data_2019_2026.keys())
    
    for zone in sorted(all_zones):
        has_2015 = zone in data_2015_2018
        has_2019 = zone in data_2019_2026
        
        # Build list of available dataframes
        dfs_to_concat = []
        period_info = []
        
        if has_2015:
            dfs_to_concat.append(data_2015_2018[zone])
            period_info.append(f"{len(data_2015_2018[zone]):>7,} (2015-2018)")
        if has_2019:
            dfs_to_concat.append(data_2019_2026[zone])
            period_info.append(f"{len(data_2019_2026[zone]):>7,} (2019-2026)")
        
        # Concatenate all available periods
        if dfs_to_concat:
            combined[zone] = pd.concat(dfs_to_concat).sort_index()
            # FIX 31/08/2026: the 2015-2018 and 2019-2026 files overlap at 2018-12-25 for some
            #    zones (BG, RO, GR) → duplicate index breaks reindex(method='ffill') in
            #    build_zone_master. Keep the later file's value.
            combined[zone] = combined[zone][~combined[zone].index.duplicated(keep='last')]
            
            print(f"  ✓ {zone:10s}: {' + '.join(period_info)} = {len(combined[zone]):>7,} hours")
    
    return combined


# ========== LOAD DATA (single merged file, full period) ==========
print(f"\n{'='*80}")
print("LOADING RESERVOIR DATA (FULL PERIOD)")
print(f"{'='*80}")

reservoir_data = load_reservoir_data_file(RESERVOIR_DATA_FILE, "2015-2026")

# ========== SUMMARY ==========
print(f"\n{'='*80}")
print(f"FINAL RESERVOIR DATA: {len(reservoir_data)} ZONES (2015-2026)")
print(f"{'='*80}")

if reservoir_data:
    all_starts = [d.index.min() for d in reservoir_data.values()]
    all_ends = [d.index.max() for d in reservoir_data.values()]
    print(f"\n📅 Overall date range (UTC): {min(all_starts)} to {max(all_ends)}")
    
    # Show which zones have data for which periods
    print(f"\n📊 Zone Coverage:")
    for zone in sorted(reservoir_data.keys()):
        zone_start = reservoir_data[zone].index.min()
        zone_end = reservoir_data[zone].index.max()
        years_covered = zone_end.year - zone_start.year + 1
        
        # Check if it's a transition zone
        if zone == 'DE_AT_LU':
            note = " (2015-2018 only)"
        elif zone == 'DE_LU':
            note = " (2019-2026)"
        elif zone == 'AT' and zone_start.year >= 2019:
            note = " (2019-2026)"
        else:
            note = ""
        
        print(f"  {zone:10s}: {zone_start.date()} to {zone_end.date()} ({years_covered} years){note}")

# Zones without data
missing_zones = [z for z in HYDRO_RESERVOIR_ZONES if z not in reservoir_data]
if missing_zones:
    print(f"\n  ZONES WITHOUT DATA ({len(missing_zones)}):")
    for zone in missing_zones:
        print(f"     {zone}")
    print(f"   → Will use seasonal fallback in Cell 11")

# ========== DIAGNOSTIC ==========
print(f"\n{'='*80}")
print("DIAGNOSTIC: RESERVOIR DATA QUALITY")
print(f"{'='*80}")

if reservoir_data:
    print(f"\nSuccessfully loaded {len(reservoir_data)} zones with ACTUAL ENTSO-E data")
    print(f"   Cell 11 will use REAL scarcity signals")
    
    # Sample visualization - prioritize DE_AT_LU if available
    if 'DE_AT_LU' in reservoir_data:
        sample_zone = 'DE_AT_LU'
        print(f"\n📊 Sample: {sample_zone} (2015-2018 data, first 48 hours)")
    elif 'DE_LU' in reservoir_data:
        sample_zone = 'DE_LU'
        print(f"\n📊 Sample: {sample_zone} (2019-2026 data, first 48 hours)")
    else:
        sample_zone = list(reservoir_data.keys())[0]
        print(f"\n📊 Sample: {sample_zone} (first 48 hours)")
    
    sample_df = reservoir_data[sample_zone]
    print(sample_df.head(48)[['filling_pct']].to_string())
    
    # Check seasonal variation
    print(f"\n📈 Seasonal variation check ({sample_zone}):")
    sample_df['month'] = sample_df.index.month
    monthly_avg = sample_df.groupby('month')['filling_pct'].mean()
    print(monthly_avg.to_string())
    
    # Check year-by-year statistics
    print(f"\n📊 Year-by-year statistics ({sample_zone}):")
    sample_df['year'] = sample_df.index.year
    yearly_avg = sample_df.groupby('year')['filling_pct'].agg(['mean', 'min', 'max'])
    print(yearly_avg.to_string())
    
    sample_df.drop(['month', 'year'], axis=1, inplace=True, errors='ignore')
    
else:
    print(f"\n NO DATA LOADED - Cell 11 will use seasonal fallback")

print(f"{'='*80}")

# ========================================
# EXPORT RESERVOIR DATA FOR SENSITIVITY ANALYSIS
# ========================================

print(f"\n💾 Exporting reservoir data for sensitivity analysis...")

# Combine all zones into single DataFrame
reservoir_combined = []

for zone, df in reservoir_data.items():
    df_copy = df.copy()
    df_copy['zone'] = zone  # Add zone column
    reservoir_combined.append(df_copy)

if reservoir_combined:
    df_reservoir_export = pd.concat(reservoir_combined, ignore_index=False)
    
    reservoir_export_file = OUTPUT_DIR / 'reservoir_data_2015_2026.csv'
    df_reservoir_export.to_csv(reservoir_export_file)
    
    print(f"   ✓ Saved: {reservoir_export_file}")
    print(f"   Zones: {len(reservoir_data)}")
    print(f"   Rows: {len(df_reservoir_export):,}")
    print(f"   Columns: {', '.join(df_reservoir_export.columns.tolist())}")
    
    # Show sample
    print(f"\n   Sample data:")
    print(df_reservoir_export.head(10))
else:
    print(f"    No reservoir data to export")


# In[20]:


for z in ['BG', 'RO', 'GR', 'ES', 'FR']:
    r = reservoir_data.get(z)
    if r is not None:
        dups = r.index.duplicated().sum()
        print(f"{z}: {len(r)} rows, {dups} duplicate timestamps, monotonic={r.index.is_monotonic_increasing}")
        if dups:
            print(r[r.index.duplicated(keep=False)].head(6))


# In[ ]:


# %% CELL 10 - MASTER DF (research Cell 10)
# %% ==========================================================================
# Cell 10 - Build master DataFrame per zone (2015-2026)
#  Updated to handle full date range and zone transitions
#  NOW EXTENDS TO 2026

def build_zone_master(zone, df_prices, df_gen, df_load, df_flows,
                      gas_prices_df, co2_prices_df, commodity_prices_df,
                      reservoir_data=None):
    """
    Build master DataFrame for a specific zone combining all data sources.
    Uses gen_{tech}_{zone} columns from Cell 4.5
     Handles 2015-2026 date range
     Works with DE_AT_LU, DE_LU, and AT zones
    """
    master = pd.DataFrame()
    
    # 1. Extract HOME ZONE price
    price_col = find_price_column(df_prices, zone)
    if price_col and price_col in df_prices.columns:
        master['price'] = df_prices[price_col]
    
    #  ADD NEIGHBOUR PRICES
    neighbours = NEIGHBOURS.get(zone, [])
    for neighbour in neighbours:
        neighbour_price_col = find_price_column(df_prices, neighbour)
        if neighbour_price_col and neighbour_price_col in df_prices.columns:
            master[f'price_{neighbour}'] = df_prices[neighbour_price_col]
    
    # 2. Extract generation using gen_{tech}_{zone} columns
    gen_tech_patterns = {
        'gas': f'gen_gas_{zone}',
        'coal': f'gen_coal_{zone}',
        'lignite': f'gen_lignite_{zone}',
        'oil': f'gen_oil_{zone}',
        'oil_shale': f'gen_oil_shale_{zone}',
        'peat': f'gen_peat_{zone}',
        'coal_derived_gas': f'gen_coal_derived_gas_{zone}',
        'nuclear': f'gen_nuclear_{zone}',
        'biomass': f'gen_biomass_{zone}',
        'chp': f'gen_chp_{zone}',
        'hydro_reservoir': f'gen_hydro_reservoir_{zone}',
        'hydro_pumped': f'gen_hydro_pumped_{zone}',
        'hydro_pumped_consumption': f'gen_hydro_pumped_consumption_{zone}',  # Add consumption
        'hydro_ror': f'gen_hydro_ror_{zone}',
        'energy_storage': f'gen_energy_storage_{zone}',
        'wind_onshore': f'gen_wind_onshore_{zone}',
        'wind_offshore': f'gen_wind_offshore_{zone}',
        'solar': f'gen_solar_{zone}',
        'geothermal': f'gen_geothermal_{zone}',
        'marine': f'gen_marine_{zone}',
        'waste': f'gen_waste_{zone}',
        'other_renewable': f'gen_other_renewable_{zone}',
        'other': f'gen_other_{zone}',
    }
    
    #  Map columns (only if they exist)
    for tech, col_name in gen_tech_patterns.items():
        if col_name in df_gen.columns:
            master[f'gen_{tech}'] = df_gen[col_name]
        else:
            # Fill with 0 for missing technologies
            master[f'gen_{tech}'] = 0
    
    # 3. Extract load
    load_col = find_load_column(df_load, zone)
    if load_col and load_col in df_load.columns:
        master['load'] = df_load[load_col]
    
    # 4. Extract flows
    flow_cols = find_flow_columns(df_flows, zone)
    
    if flow_cols['imports']:
        import_cols = [c for c in flow_cols['imports'] if c in df_flows.columns]
        if import_cols:
            master['imports'] = df_flows[import_cols].sum(axis=1)
    
    if flow_cols['exports']:
        export_cols = [c for c in flow_cols['exports'] if c in df_flows.columns]
        if export_cols:
            master['exports'] = df_flows[export_cols].sum(axis=1)
    
    if 'imports' in master.columns or 'exports' in master.columns:
        master['net_imports'] = master.get('imports', 0) - master.get('exports', 0)
    
    # 5. Add commodity prices (resample to hourly if needed)
    if not gas_prices_df.empty:
        gas_price = get_zone_gas_price(gas_prices_df, zone)
        if len(gas_price) < len(master):
            gas_price = gas_price.resample('H').ffill()
        master['gas_price'] = gas_price.reindex(master.index)
        master['gas_price'] = master['gas_price'].fillna(method='bfill').fillna(method='ffill')
        
        # #  Fill any remaining NaNs with default
        # if master['gas_price'].isna().any():
        #     master['gas_price'] = master['gas_price'].fillna(30.0)
        # Fixed 2026/08/27 to avoid filling with defaults

    if not co2_prices_df.empty:
        co2_price = get_zone_co2_price(co2_prices_df, zone)
        if len(co2_price) < len(master):
            co2_price = co2_price.resample('H').ffill()
        master['co2_price'] = co2_price.reindex(master.index)
        master['co2_price'] = master['co2_price'].fillna(method='bfill').fillna(method='ffill')
        
        #  Fill any remaining NaNs with default
        # if master['co2_price'].isna().any():
        #     master['co2_price'] = master['co2_price'].fillna(80.0)
        # Fixed 2026/08/27 to avoid filling with defaults

    if not commodity_prices_df.empty:
        if 'coal_price' in commodity_prices_df.columns:
            coal = commodity_prices_df['coal_price']
            if len(coal) < len(master):
                coal = coal.resample('H').ffill()
            master['coal_price'] = coal.reindex(master.index)
            master['coal_price'] = master['coal_price'].fillna(method='bfill').fillna(method='ffill')
            
            # #  Fill any remaining NaNs with default
            # if master['coal_price'].isna().any():
            #     master['coal_price'] = master['coal_price'].fillna(80.0)
            # Fixed 2026/08/27 to avoid filling with defaults
        
        if 'oil_price' in commodity_prices_df.columns:
            oil = commodity_prices_df['oil_price']
            if len(oil) < len(master):
                oil = oil.resample('H').ffill()
            master['oil_price'] = oil.reindex(master.index)
            master['oil_price'] = master['oil_price'].fillna(method='bfill').fillna(method='ffill')  #  FIXED: was method['ffill')
            
            #  Fill any remaining NaNs with default
            # if master['oil_price'].isna().any():
            #     master['oil_price'] = master['oil_price'].fillna(500.0)
            # Fixed 2026/08/27 to avoid filling with defaults
        
        #  Oil shale fuel price (fixed from ERAA 2025)
        master['oil_shale_price'] = THERMAL_PARAMS['oil_shale']['fuel_price_eur_per_mwh_th']
    
    # 6. Add reservoir data if available ( FIXED: timezone alignment)
    if reservoir_data and zone in reservoir_data:
        res = reservoir_data[zone]
        # FIX 31/08/2026: guard against duplicate timestamps (reindex with method='ffill'
        #    requires a unique, sorted source index)
        res = res[~res.index.duplicated(keep='last')].sort_index()
        if not res.empty and 'filling_pct' in res.columns:
            try:
                # Try direct reindex first (works if timezones match)
                master['reservoir_filling'] = res['filling_pct'].reindex(master.index, method='ffill')
            except TypeError as e:
                # Timezone mismatch → strip timezones and reindex
                if 'Cannot compare dtypes' in str(e):
                    # Make both timezone-naive
                    res_naive = res.copy()
                    if res_naive.index.tz is not None:
                        res_naive.index = res_naive.index.tz_localize(None)
                    
                    master_naive_index = master.index
                    if master_naive_index.tz is not None:
                        master_naive_index = master_naive_index.tz_localize(None)
                    
                    # Reindex with both naive
                    master['reservoir_filling'] = res_naive['filling_pct'].reindex(
                        master_naive_index, 
                        method='ffill'
                    )
                else:
                    # Unknown error, re-raise
                    raise
    
    master['zone'] = zone
    
    #  Only drop rows where price is NaN (keep all data)
    master = master.dropna(subset=['price']) if 'price' in master.columns else master
    

    if master.index.duplicated().any():
        n = int(master.index.duplicated().sum())
        master = master[~master.index.duplicated(keep='first')].sort_index()
        print(f"   {zone}: removed {n} duplicate timestamp(s) at the file join")
    
    return master


# ========== TEST BUILD FUNCTION ==========
print("=" * 80)
print("TESTING build_zone_master FUNCTION (2015-2026)")
print("=" * 80)

if complete_zones:
    # Test with multiple zones including transition zones if available
    test_zones = []
    
    # Prioritize transition zones for testing
    if 'DE_AT_LU' in complete_zones:
        test_zones.append('DE_AT_LU')
    if 'DE_LU' in complete_zones:
        test_zones.append('DE_LU')
    if 'AT' in complete_zones:
        test_zones.append('AT')
    
    # Add first complete zone if transition zones not available
    if not test_zones:
        test_zones.append(complete_zones[0])
    
    for test_zone in test_zones[:3]:  # Test up to 3 zones
        print(f"\n{'='*80}")
        print(f"BUILDING master_test FOR {test_zone}")
        print(f"{'='*80}")
        
        master_test = build_zone_master(
            test_zone, df_prices, df_gen, df_load, df_flows,
            gas_prices_df, co2_prices_df, commodity_prices_df,
            reservoir_data
        )
        
        print(f"\nmaster_test shape: {master_test.shape}")
        print(f"   Date range: {master_test.index.min()} to {master_test.index.max()}")
        print(f"   Years covered: {master_test.index.year.min()} - {master_test.index.year.max()}")
        
        # Check generation columns
        gen_cols = [c for c in master_test.columns if c.startswith('gen_')]
        gen_cols_with_data = [c for c in gen_cols if master_test[c].sum() > 0]
        print(f"\nGeneration columns found: {len(gen_cols)}")
        print(f"   Columns with data: {len(gen_cols_with_data)}")
        print(f"   Technologies: {[c.replace('gen_', '') for c in gen_cols_with_data]}")
        
        # Check neighbour prices
        neighbour_price_cols = [c for c in master_test.columns if c.startswith('price_')]
        print(f"\n Neighbour price columns: {len(neighbour_price_cols)}")
        if neighbour_price_cols:
            print(f"   Neighbours: {[c.replace('price_', '') for c in neighbour_price_cols]}")
        
        # Check commodity prices
        commodity_cols = ['gas_price', 'co2_price', 'coal_price', 'oil_price', 'oil_shale_price']
        available_commodities = [c for c in commodity_cols if c in master_test.columns]
        print(f"\n Commodity prices: {len(available_commodities)}")
        print(f"   Available: {available_commodities}")
        
        # Check for NaNs in critical columns
        critical_cols = ['price', 'load']
        if 'price' in master_test.columns:
            price_nans = master_test['price'].isna().sum()
            print(f"\n Price column: {len(master_test) - price_nans:,} / {len(master_test):,} non-NaN values")
        
        if 'load' in master_test.columns:
            load_nans = master_test['load'].isna().sum()
            print(f" Load column: {len(master_test) - load_nans:,} / {len(master_test):,} non-NaN values")
        
        # Check reservoir data
        if 'reservoir_filling' in master_test.columns:
            res_data = master_test['reservoir_filling'].dropna()
            if len(res_data) > 0:
                print(f"\n Reservoir data: {len(res_data):,} / {len(master_test):,} hours")
                print(f"   Range: {res_data.min():.1f}% - {res_data.max():.1f}%")
        
        # Year-by-year check
        print(f"\n Year-by-year data availability:")
        master_test['year'] = master_test.index.year
        for year in sorted(master_test['year'].unique()):
            year_data = master_test[master_test['year'] == year]
            print(f"   {year}: {len(year_data):>6,} hours")
        master_test.drop('year', axis=1, inplace=True)
        
        # Sample data
        print(f"\n Sample data (first 10 rows):")
        display_cols = ['price', 'load', 'gen_gas', 'gen_nuclear', 'gen_solar', 'gas_price', 'co2_price']
        available_display = [c for c in display_cols if c in master_test.columns]
        print(master_test[available_display].head(10))
        
        # Store for potential use
        if test_zone == test_zones[0]:
            master_test_sample = master_test.copy()

print("\n" + "=" * 80)
print("build_zone_master FUNCTION TESTED SUCCESSFULLY")
print("=" * 80)
print("\nReady to build master DataFrames for all zones!")


# In[ ]:


# %% CELL 10a - HYDRO SRMC FUNCTIONS (research 10a)
# %% ==========================================================================
# Cell 10a - Hydro Reservoir SRMC Function and Hydro Pumped Storage function (2015-2026)
#  Updated to handle DE_AT_LU, DE_LU, and AT zones |  EXTENDS TO 2026
# 20260903: flexibility buffer + base SRMC high aligned across both functions
# 20260918: minimum spread aligned to 8 EUR/MWh in both; functions simplified;
#           PARAMS dicts are the single source of truth (also read by Cell 11's
#           vectorized SRMC block - do NOT remove them)

# ========== ZONES WITH NEGLIGIBLE HYDRO RESERVOIR CAPACITY ==========
ZONES_EXCLUDE_HYDRO_RESERVOIR = [
    'BE',     # Belgium: minimal reservoir capacity
    'NL',     # Netherlands: no reservoir hydro
    'DK_1',   # Denmark West: no reservoir hydro
    'DK_2',   # Denmark East: no reservoir hydro
    'GB',     # Great Britain: mostly pumped storage, not reservoir
    'IE_SEM', # Ireland: minimal reservoir capacity
]

# ========== PARAMETERS (from control file hydro_params sheet, via Cell 0) ==========
# Used here AND by Cell 11's vectorized SRMC blocks - single source of truth.
HYDRO_RESERVOIR_PARAMS = CFG.HYDRO_RESERVOIR_PARAMS
PUMPED_STORAGE_PARAMS = CFG.PUMPED_STORAGE_PARAMS

# NIGHT_HOURS comes from CFG (control file) via Cell 0

def compute_dynamic_hydro_srmc(
    reservoir_level_pct: float,
    residual_demand: float,
    residual_demand_max: float,
    gas_srmc_avg: float = None,
    base_srmc_low: float = HYDRO_RESERVOIR_PARAMS['base_srmc_low'],
    base_srmc_high: float = HYDRO_RESERVOIR_PARAMS['base_srmc_high'],
    demand_weight: float = HYDRO_RESERVOIR_PARAMS['demand_weight'],
    flexibility_buffer: float = HYDRO_RESERVOIR_PARAMS['flexibility_buffer'],
    gas_ceiling_factor: float = HYDRO_RESERVOIR_PARAMS['gas_ceiling_factor']
):
    """
    Reservoir hydro SRMC = opportunity cost, anchored to gas.

      srmc_base = gas_ceiling * [min_factor + scarcity * (max_factor - min_factor)]
                  + demand push toward the ceiling
      where gas_ceiling = gas_ceiling_factor * gas SRMC,
            scarcity    = 1 - filling/100  (empty reservoir -> bids near ceiling),
            demand push = demand_weight * (RD / RD_max) * (gas_ceiling - srmc_base).
      The bid becomes a band of +/- flexibility_buffer, clipped to
      [base_srmc_low, base_srmc_high], capped at the gas ceiling,
      with a minimum spread of 8 EUR/MWh.
    """
    reservoir_level_pct = max(0.0, min(100.0, reservoir_level_pct if pd.notna(reservoir_level_pct) else 50.0))
    residual_demand_max = max(1.0, residual_demand_max) if pd.notna(residual_demand_max) else 1.0

    reservoir_scarcity = 1.0 - (reservoir_level_pct / 100.0)
    demand_factor = max(0.0, min(1.0, residual_demand / residual_demand_max))

    if gas_srmc_avg is not None and pd.notna(gas_srmc_avg):
        gas_ceiling = gas_srmc_avg * gas_ceiling_factor
        min_factor = HYDRO_RESERVOIR_PARAMS['min_factor']
        max_factor = HYDRO_RESERVOIR_PARAMS['max_factor']
        bid_factor = min_factor + reservoir_scarcity * (max_factor - min_factor)
        srmc_base = gas_ceiling * bid_factor
        demand_push = demand_factor * (gas_ceiling - srmc_base) * demand_weight
        srmc_base = min(gas_ceiling, srmc_base + demand_push)
        srmc_base = max(base_srmc_low, min(base_srmc_high, srmc_base))
    else:
        gas_ceiling = None
        srmc_base = (base_srmc_low + reservoir_scarcity * (base_srmc_high - base_srmc_low)
                     + demand_factor * (base_srmc_high - base_srmc_low) * demand_weight)

    srmc_min = max(base_srmc_low, srmc_base - flexibility_buffer)
    srmc_max = min(base_srmc_high, srmc_base + flexibility_buffer)
    if gas_ceiling is not None:
        srmc_max = min(srmc_max, gas_ceiling)
    if srmc_max - srmc_min < 8.0:
        mid = (srmc_min + srmc_max) / 2
        srmc_min = max(base_srmc_low, mid - 4.0)
        srmc_max = min(base_srmc_high, mid + 4.0)
    return (srmc_min, srmc_max)


def compute_pumped_storage_srmc(
    hour: int,
    gas_srmc_avg: float = None,
    base_srmc_low: float = PUMPED_STORAGE_PARAMS['base_srmc_low'],
    base_srmc_high: float = PUMPED_STORAGE_PARAMS['base_srmc_high'],
    night_factor: float = PUMPED_STORAGE_PARAMS['night_factor'],
    day_factor: float = PUMPED_STORAGE_PARAMS['day_factor'],
    flexibility_buffer: float = PUMPED_STORAGE_PARAMS['flexibility_buffer']
):
    """
    Pumped storage SRMC, anchored to the daily gas benchmark with day/night gating.

      srmc_base = gas_srmc_avg * (night_factor if night else day_factor),
      clipped to [base_srmc_low, base_srmc_high]; night = NIGHT_HOURS
      (22:00-04:59). The bid becomes a band of +/- flexibility_buffer,
      clipped to the same range, minimum spread 8 EUR/MWh (all bounds as
      for hydro reservoir). Fallback without gas SRMC: 10 (night) / 80 (day).
    """
    is_night = hour in NIGHT_HOURS

    if gas_srmc_avg is not None and pd.notna(gas_srmc_avg):
        srmc_base = gas_srmc_avg * (night_factor if is_night else day_factor)
        srmc_base = max(base_srmc_low, min(base_srmc_high, srmc_base))
    else:
        srmc_base = base_srmc_low + 10.0 if is_night else base_srmc_high - 20.0

    srmc_min = max(base_srmc_low, srmc_base - flexibility_buffer)
    srmc_max = min(base_srmc_high, srmc_base + flexibility_buffer)
    if srmc_max - srmc_min < 8.0:
        mid = (srmc_min + srmc_max) / 2
        srmc_min = max(base_srmc_low, mid - 4.0)
        srmc_max = min(base_srmc_high, mid + 4.0)
    return (srmc_min, srmc_max)

# ========== SUMMARY ==========
print("=" * 80)
print("HYDRO FUNCTIONS DEFINED (2015-2026)")
print("=" * 80)
print("\n📊 Hydro reservoir parameters:")
for key, value in HYDRO_RESERVOIR_PARAMS.items():
    print(f"   • {key}: {value}")
print("\n📊 Pumped storage parameters:")
for key, value in PUMPED_STORAGE_PARAMS.items():
    print(f"   • {key}: {value}")
print(f"   • night_hours: {NIGHT_HOURS}")
print(f"\n🚫 Zones excluded from hydro reservoir ({len(ZONES_EXCLUDE_HYDRO_RESERVOIR)}):")
print(f"   {', '.join(ZONES_EXCLUDE_HYDRO_RESERVOIR)}")
print("\nZone transitions: DE_AT_LU (2015-2018) hydro_reservoir via AT data;")
print("   DE_LU and AT (2019+) separate, DE_LU + CZ via AT proxy, PL via SE_4")
print("=" * 80)


# In[ ]:


# %% CELL 11 - SRMC + 7-STEP (research Cell 11)
# %% ==========================================================================
# Cell 11: FINAL CORRECTED Dynamic SRMC Calculation and residual demand calculation (2015-2026)
#  Updated to handle DE_AT_LU, DE_LU, and AT zones
#  NOW EXTENDS TO 2026

# ========== ZONES WITH SIGNIFICANT HYDRO RESERVOIR GENERATION ==========
# Use the list from Cell 9
ZONES_WITH_HYDRO_RESERVOIR_GEN = HYDRO_RESERVOIR_ZONES



def calculate_srmc_all_techs(master_df, zone):
    """
    Calculate SRMC with CORRECTED dynamic efficiency interpolation.
     Works across 2015-2026 period
     Handles zone transitions (DE_AT_LU, DE_LU, AT)
    
    Key features:
    - Dynamic efficiency per hour based on YEARLY generation levels
    - Hydro reservoir: Shadow price from filling %
    - Pumped storage: Round-trip arbitrage cost
    - Nuclear: Uranium cost + dynamic efficiency
    - RES: Static O&M ranges
    """

    if master_df.index.duplicated().any():
        n = int(master_df.index.duplicated().sum())
        master_df = master_df[~master_df.index.duplicated(keep='first')].sort_index()
        print(f"   {zone}: removed {n} duplicate timestamp(s) at the file join")

    srmc = pd.DataFrame(index=master_df.index)
    
    # Add year column for YEARLY aggregation
    master_df['year'] = master_df.index.year
    
    # ========== STEP 1: EXTRACT FUEL PRICES ==========
    for _col in ('gas_price', 'co2_price', 'coal_price', 'oil_price'):
        if _col not in master_df.columns:
            raise KeyError(f"{_col} missing from master_df - "
                           f"commodity loading (Cells 6-8) incomplete")
    gas_price = master_df['gas_price']
    co2_price = master_df['co2_price']

    # Coal price: EUR/tonne -> EUR/MWh_thermal
    coal_price_per_tonne = master_df['coal_price']
    coal_cv = THERMAL_PARAMS['coal']['fuel_cv']  # 25.0 GJ/t
    coal_price = (coal_price_per_tonne / coal_cv) * 3.6

    # Oil price: EUR/tonne -> EUR/MWh_thermal
    oil_price_per_tonne = master_df['oil_price']
    oil_cv = THERMAL_PARAMS['oil']['fuel_cv']  # 42.0 GJ/t
    oil_price = (oil_price_per_tonne / oil_cv) * 3.6
    
    # ========== STEP 2: DYNAMIC SRMC FOR THERMAL TECHNOLOGIES ( FIXED) ==========

    thermal_techs = {
        'gas': (['gen_gas', 'gen_fossil_gas'], 'gas', 'co2_factor'),
        'coal': (['gen_coal', 'gen_fossil_hard_coal'], 'coal', 'co2_factor'),
        'lignite': (['gen_lignite', 'gen_fossil_brown_coal'], 'lignite', 'co2_factor'),
        'oil': (['gen_oil', 'gen_fossil_oil'], 'oil', 'co2_factor'),
        'oil_shale': (['gen_oil_shale'], 'oil_shale', 'co2_factor'),  
        'biomass': (['gen_biomass'], 'biomass', None),
        'chp': (['gen_other'], 'chp', 'co2_factor'),
    }


    gen_thresholds = SRMC_CALC_THRESHOLDS   # <- control file ('srmc_calc' scope)

    for tech_key, (gen_patterns, fuel_key, has_co2) in thermal_techs.items():
        # ==========  CORRECTED: FIND GENERATION COLUMN WITH ZONE SUFFIX ==========
        gen_col = None
        for pattern in gen_patterns:
            for col in master_df.columns:
                col_lower = col.lower()
                #  Match pattern AND check for zone suffix
                if pattern.lower() in col_lower:
                    # Verify it's actually a generation column (starts with 'gen_')
                    if col.startswith('gen_'):
                        gen_col = col
                        break
            if gen_col:
                break
        
        #  DIAGNOSTIC: Log if column not found (helps debug future issues)
        if gen_col is None:
            # Only print warning if zone should have this technology
            # (Don't spam for every missing tech in every zone)
            continue
        
        # Get fuel price
        if tech_key == 'gas' or tech_key == 'chp':
            fuel_price = gas_price
        elif tech_key == 'coal':
            fuel_price = coal_price
        elif tech_key == 'lignite':
            fuel_price = THERMAL_PARAMS['lignite']['fuel_price_eur_per_mwh_th']
        elif tech_key == 'oil':
            fuel_price = oil_price
        elif tech_key == 'oil_shale':  # NEW
            fuel_price = master_df.get('oil_shale_price', THERMAL_PARAMS['oil_shale']['fuel_price_eur_per_mwh_th'])
        elif tech_key == 'biomass':
            fuel_price = THERMAL_PARAMS['biomass']['fuel_price_eur_per_mwh_th']
            subsidy = THERMAL_PARAMS['biomass'].get('subsidy_eur_per_mwh', 0.0)
        else:
            fuel_price = 30.0
            subsidy = 0.0
        
        var_om = THERMAL_PARAMS[fuel_key]['var_om']
        eff_map = get_efficiency(zone, tech_key)
        
        if gen_col is None or gen_col not in master_df.columns:
            # No generation data, use average efficiency
            if has_co2:
                co2_factor = THERMAL_PARAMS[fuel_key][has_co2]
                srmc[f'srmc_{tech_key}'] = (fuel_price / eff_map['avg']) + (co2_price * co2_factor / eff_map['avg']) + var_om
            else:
                srmc[f'srmc_{tech_key}'] = (fuel_price / eff_map['avg']) + var_om
            continue
        
        # DYNAMIC EFFICIENCY BASED ON YEARLY GENERATION LEVELS
        threshold = gen_thresholds.get(tech_key, 0.1)
        active_mask = master_df[gen_col] > threshold
        
        if active_mask.sum() == 0:
            # No active hours, use average
            if has_co2:
                co2_factor = THERMAL_PARAMS[fuel_key][has_co2]
                srmc[f'srmc_{tech_key}'] = (fuel_price / eff_map['avg']) + (co2_price * co2_factor / eff_map['avg']) + var_om
            else:
                srmc[f'srmc_{tech_key}'] = (fuel_price / eff_map['avg']) + var_om
            continue
        
        #  Calculate YEARLY min/max generation (for normalization)
        yearly_stats = master_df[active_mask].groupby('year')[gen_col].agg(['min', 'max']).reset_index()
        yearly_stats.columns = ['year', f'{tech_key}_gen_min_year', f'{tech_key}_gen_max_year']
        
        # Merge with master_df
        temp_df = master_df.reset_index()
        temp_df = temp_df.merge(yearly_stats, on='year', how='left')
        temp_df = temp_df.set_index(master_df.index.name if master_df.index.name else 'index')
        
        # Calculate generation factor (0 = yearly min, 1 = yearly max)
        gen_hour = temp_df[gen_col]
        gen_min_year = temp_df[f'{tech_key}_gen_min_year']
        gen_max_year = temp_df[f'{tech_key}_gen_max_year']
        
        gen_range = gen_max_year - gen_min_year
        gen_range = gen_range.replace(0, np.nan).fillna(1.0)
        
        gen_factor = ((gen_hour - gen_min_year) / gen_range).clip(0, 1)
        
        #  INTERPOLATE EFFICIENCY
        # gen_factor = 0 (min gen) → use eff_max (best plants)
        # gen_factor = 1 (max gen) → use eff_min (worst plants)
        dynamic_eff = eff_map['max'] - gen_factor * (eff_map['max'] - eff_map['min'])
        
        #  CALCULATE DYNAMIC SRMC
        if has_co2:
            co2_factor = THERMAL_PARAMS[fuel_key][has_co2]
            dynamic_srmc = (fuel_price / dynamic_eff) + (co2_price * co2_factor / dynamic_eff) + var_om
        else:
            # FOR BIOMASS, SUBTRACT SUBSIDY
            if tech_key == 'biomass':
                dynamic_srmc = ((fuel_price / dynamic_eff) + var_om) - subsidy
                dynamic_srmc = dynamic_srmc.clip(lower=0.0)  #  Ensure non-negative
            else:
                dynamic_srmc = (fuel_price / dynamic_eff) + var_om
        
        # Initialize with average (for inactive hours)
        if has_co2:
            avg_srmc = (fuel_price / eff_map['avg']) + (co2_price * co2_factor / eff_map['avg']) + var_om
        else:
            avg_srmc = (fuel_price / eff_map['avg']) + var_om
        
        srmc[f'srmc_{tech_key}'] = avg_srmc
        srmc.loc[active_mask, f'srmc_{tech_key}'] = dynamic_srmc.loc[active_mask]
        
        # Diagnostic
        coverage = active_mask.sum() / len(master_df) * 100
        print(f"  ✓ {tech_key:10s}: Dynamic SRMC for {active_mask.sum():,} hours ({coverage:.1f}% coverage)")
    
    # ========== STEP 3: DYNAMIC SRMC FOR NUCLEAR ==========
    # Uses uranium cost + dynamic efficiency based on generation level
    nuclear_gen_col = None
    for col in master_df.columns:
        if 'gen_nuclear' in col.lower():
            nuclear_gen_col = col
            break
    
    if nuclear_gen_col and nuclear_gen_col in master_df.columns:
        uranium_cost = THERMAL_PARAMS['nuclear']['uranium_cost_eur_per_mwh_th']
        var_om = THERMAL_PARAMS['nuclear']['var_om']
        eff_min = THERMAL_PARAMS['nuclear']['eff_min']
        eff_max = THERMAL_PARAMS['nuclear']['eff_max']
        eff_avg = (eff_min + eff_max) / 2
        
        active_mask = master_df[nuclear_gen_col] > 0.1
        
        if active_mask.sum() > 0:
            # Yearly normalization
            yearly_stats = master_df[active_mask].groupby('year')[nuclear_gen_col].agg(['min', 'max']).reset_index()
            yearly_stats.columns = ['year', 'nuclear_gen_min_year', 'nuclear_gen_max_year']
            
            temp_df = master_df.reset_index()
            temp_df = temp_df.merge(yearly_stats, on='year', how='left')
            temp_df = temp_df.set_index(master_df.index.name if master_df.index.name else 'index')
            
            gen_hour = temp_df[nuclear_gen_col]
            gen_min_year = temp_df['nuclear_gen_min_year']
            gen_max_year = temp_df['nuclear_gen_max_year']
            
            gen_range = gen_max_year - gen_min_year
            gen_range = gen_range.replace(0, np.nan).fillna(1.0)
            
            gen_factor = ((gen_hour - gen_min_year) / gen_range).clip(0, 1)
            
            # Dynamic efficiency
            dynamic_eff = eff_max - gen_factor * (eff_max - eff_min)
            dynamic_srmc = (uranium_cost / dynamic_eff) + var_om
            
            # Initialize and update
            srmc['srmc_nuclear'] = (uranium_cost / eff_avg) + var_om
            srmc.loc[active_mask, 'srmc_nuclear'] = dynamic_srmc.loc[active_mask]
            
            print(f"  ✓ {'nuclear':10s}: Dynamic SRMC for {active_mask.sum():,} hours ({active_mask.sum()/len(master_df)*100:.1f}% coverage)")
        else:
            srmc['srmc_nuclear'] = (uranium_cost / eff_avg) + var_om
    else:
        uranium_cost = THERMAL_PARAMS['nuclear']['uranium_cost_eur_per_mwh_th']
        var_om = THERMAL_PARAMS['nuclear']['var_om']
        eff_avg = (THERMAL_PARAMS['nuclear']['eff_min'] + THERMAL_PARAMS['nuclear']['eff_max']) / 2
        srmc['srmc_nuclear'] = (uranium_cost / eff_avg) + var_om
    

    # ========== STEP 4: VECTORIZED HYDRO RESERVOIR SRMC (WITH PROXY SUPPORT) ==========
    if zone in ZONES_EXCLUDE_HYDRO_RESERVOIR:
        #  ZONE EXCLUDED: No reservoir capacity
        srmc['srmc_hydro_reservoir_min'] = 999.0
        srmc['srmc_hydro_reservoir_max'] = 999.0
        srmc['srmc_hydro_reservoir_avg'] = 999.0
        print(f"   {'hydro_reservoir':10s}: Zone {zone} EXCLUDED (negligible capacity)")

    else:
        # ZONE ELIGIBLE: Check if zone has hydro reservoir GENERATION
        hydro_res_gen_col = None
        for col in master_df.columns:
            if 'hydro_reservoir' in col.lower() or 'hydro_water_reservoir' in col.lower():
                if 'gen' in col.lower():
                    hydro_res_gen_col = col
                    break
        
        has_hydro_generation = False
        if hydro_res_gen_col and hydro_res_gen_col in master_df.columns:
            total_hydro_gen = master_df[hydro_res_gen_col].sum()
            has_hydro_generation = total_hydro_gen > 10
        
        #  VECTORIZED CALCULATION
        if has_hydro_generation:
            # ========== GET FILLING PERCENTAGE (WITH PROXY SUPPORT) ==========
            filling = None
            filling_source = None
            
            #  Check if zone needs proxy
            effective_zone = zone
            if zone in RESERVOIR_FILLING_PROXIES:
                effective_zone = RESERVOIR_FILLING_PROXIES[zone]
                print(f"   {zone}: Using {effective_zone} as reservoir proxy")
            
            #  PRIORITY 1: Use Cell 9 reservoir data (direct or proxy)
            if effective_zone in reservoir_data:
                try:
                    # Get reservoir data (from proxy or direct)
                    reservoir_df = reservoir_data[effective_zone]
                    
                    #  CRITICAL: Align timezone
                    if master_df.index.tz is None:
                        master_index_utc = master_df.index.tz_localize('UTC')
                    else:
                        master_index_utc = master_df.index.tz_convert('UTC')
                    
                    # Align reservoir data to master_df index
                    filling_series = reservoir_df['filling_pct'].reindex(
                        master_index_utc,
                        method='ffill',
                        tolerance=pd.Timedelta('7 days')
                    )
                    
                    # Convert back to master_df's original timezone
                    filling_series.index = master_df.index
                    
                    # Check coverage
                    valid_count = filling_series.notna().sum()
                    coverage = valid_count / len(master_df) * 100
                    
                    if coverage > 50:
                        filling = filling_series
                        if zone == effective_zone:
                            filling_source = f'ENTSO-E ({coverage:.0f}% coverage)'
                        else:
                            filling_source = f'{effective_zone} proxy ({coverage:.0f}% coverage)'
                        
                        # Fill NaN with forward/backward fill
                        filling = filling.fillna(method='ffill').fillna(method='bfill')
                        
                        # If still NaN after ffill/bfill, use median of available data
                        if filling.isna().any():
                            median_fill = filling.median()
                            if pd.notna(median_fill):
                                filling = filling.fillna(median_fill)
                            else:
                                # Last resort: use 50% if no data at all
                                print(f"   {zone}: No valid filling data, using 50% default")
                                filling = filling.fillna(50.0)
                    else:
                        print(f"   {zone}: {effective_zone} proxy coverage only {coverage:.0f}% (too low, < 50%)")
                        filling = None
                
                except Exception as e:
                    print(f"   {zone}: Error loading {effective_zone} proxy data: {e}")
                    filling = None
            else:
                print(f"   {zone}: No proxy data available (effective_zone={effective_zone} not in reservoir_data)")
                filling = None
            
            #  PRIORITY 2: Check if already in master_df (from previous merge)
            if filling is None and 'reservoir_filling' in master_df.columns:
                filling_series = master_df['reservoir_filling']
                if filling_series.notna().sum() > len(master_df) * 0.5:
                    filling = filling_series
                    filling_source = 'master_df column'
                    print(f"   {zone}: Using reservoir_filling column from master_df")
                else:
                    print(f"   {zone}: reservoir_filling column has insufficient data")
            
            #  IF NO DATA: Skip hydro_reservoir calculation entirely
            if filling is None:
                print(f"   {'hydro_reservoir':10s}: Zone {zone} - NO DATA AVAILABLE (no ENTSO-E, no proxy, no master_df column)")
                srmc['srmc_hydro_reservoir_min'] = np.nan
                srmc['srmc_hydro_reservoir_max'] = np.nan
                srmc['srmc_hydro_reservoir_avg'] = np.nan
            
            else:
                # ========== PROCEED WITH SRMC CALCULATION ==========
                
                # Diagnostic output
                mean_fill = filling.mean()
                min_fill = filling.min()
                max_fill = filling.max()
                std_fill = filling.std()
                
                # ========== CALCULATE RESIDUAL DEMAND ==========
                non_dispatchable_cols = [
                    c for c in master_df.columns 
                    if any(tech in c.lower() for tech in ['wind_onshore', 'wind_offshore', 'solar', 'hydro_ror', 'run-of-river'])
                ]
                
                if non_dispatchable_cols:
                    non_dispatchable_gen = master_df[non_dispatchable_cols].sum(axis=1)
                else:
                    non_dispatchable_gen = 0
                
                residual_demand = master_df.get('load', 0) - non_dispatchable_gen
                residual_demand = residual_demand.clip(lower=0)
                residual_demand_max = residual_demand.max() if len(residual_demand) > 0 else 1.0
                
                # ========== GET GAS SRMC FOR CEILING ==========
                if 'srmc_gas' in srmc.columns:
                    gas_srmc_avg = srmc['srmc_gas']
                else:
                    gas_price = master_df['gas_price']    # fail loudly if commodities missing
                    co2_price = master_df['co2_price']
                    eff_map = get_efficiency(zone, 'gas')
                    gas_srmc_avg = (gas_price / eff_map['avg']) + (co2_price * THERMAL_PARAMS['gas']['co2_factor'] / eff_map['avg']) + THERMAL_PARAMS['gas']['var_om']
                
                # ========== VECTORIZED SRMC USING CELL 10a PARAMETERS ==========
                params = HYDRO_RESERVOIR_PARAMS
                
                gas_ceiling = gas_srmc_avg * params['gas_ceiling_factor']
                reservoir_scarcity = 1.0 - (filling / 100.0)
                bid_factor = params['min_factor'] + (reservoir_scarcity * (params['max_factor'] - params['min_factor']))
                srmc_base = gas_ceiling * bid_factor
                
                demand_factor = (residual_demand / residual_demand_max).clip(0, 1)
                demand_push = demand_factor * (gas_ceiling - srmc_base) * params['demand_weight']
                srmc_base = (srmc_base + demand_push).clip(upper=gas_ceiling)
                
                buffer = params['flexibility_buffer']
                srmc['srmc_hydro_reservoir_min'] = (srmc_base - buffer).clip(lower=params['base_srmc_low'])
                srmc['srmc_hydro_reservoir_max'] = (srmc_base + buffer).clip(upper=params['base_srmc_high'])
                srmc['srmc_hydro_reservoir_max'] = srmc['srmc_hydro_reservoir_max'].clip(upper=gas_ceiling)
                
                # Enforce minimum spread
                spread = srmc['srmc_hydro_reservoir_max'] - srmc['srmc_hydro_reservoir_min']
                too_narrow = spread < 8.0
                
                if too_narrow.any():
                    mid = (srmc['srmc_hydro_reservoir_min'] + srmc['srmc_hydro_reservoir_max']) / 2
                    srmc.loc[too_narrow, 'srmc_hydro_reservoir_min'] = (mid - 4.0).clip(lower=params['base_srmc_low'])
                    srmc.loc[too_narrow, 'srmc_hydro_reservoir_max'] = (mid + 4.0).clip(upper=params['base_srmc_high'])
                
                srmc['srmc_hydro_reservoir_avg'] = (srmc['srmc_hydro_reservoir_min'] + srmc['srmc_hydro_reservoir_max']) / 2
                
                # PRINT DIAGNOSTIC
                print(f"  ✓ {'hydro_reservoir':10s}: Source={filling_source} | "
                    f"Fill: {min_fill:.1f}%-{mean_fill:.1f}%-{max_fill:.1f}% (σ={std_fill:.1f}) | "
                    f"SRMC: {srmc['srmc_hydro_reservoir_avg'].mean():.1f} EUR avg")
        
        elif zone in ZONES_WITH_HYDRO_RESERVOIR_GEN and not has_hydro_generation:
            # Zone should have hydro but no generation column found
            print(f"   {'hydro_reservoir':10s}: Zone {zone} expected to have generation but none found")
            srmc['srmc_hydro_reservoir_min'] = np.nan
            srmc['srmc_hydro_reservoir_max'] = np.nan
            srmc['srmc_hydro_reservoir_avg'] = np.nan
        
        else:
            # Zone has no hydro reservoir generation
            print(f"   {'hydro_reservoir':10s}: Zone {zone} has no reservoir generation")
            srmc['srmc_hydro_reservoir_min'] = np.nan
            srmc['srmc_hydro_reservoir_max'] = np.nan
            srmc['srmc_hydro_reservoir_avg'] = np.nan

    # ========== STEP 5: VECTORIZED PUMPED STORAGE SRMC ( USES CELL 10a PARAMS) ==========
    
    #  Import parameters from Cell 10a (single source of truth!)
    params = PUMPED_STORAGE_PARAMS
    
    if 'price' in master_df.columns:
        #  VECTORIZED: Extract hour of day
        hour = master_df.index.hour
        is_night = hour.isin(NIGHT_HOURS)  # ← Use global NIGHT_HOURS from Cell 10a
        
        #  VECTORIZED: daily benchmark = mean GAS SRMC over the calendar day
        if 'srmc_gas' in srmc.columns:
            daily_gas_srmc = srmc['srmc_gas'].groupby(srmc.index.date).transform('mean')

            night_factor = params['night_factor']
            day_factor = params['day_factor']
            buffer = params['flexibility_buffer']

            srmc_base = np.where(is_night, daily_gas_srmc * night_factor,
                                           daily_gas_srmc * day_factor)
            srmc_base = np.clip(srmc_base, params['base_srmc_low'], params['base_srmc_high'])

            srmc['srmc_pumped_storage_min'] = np.clip(srmc_base - buffer,
                                                      params['base_srmc_low'], None)
            srmc['srmc_pumped_storage_max'] = np.clip(srmc_base + buffer,
                                                      None, params['base_srmc_high'])
            

            spread = srmc['srmc_pumped_storage_max'] - srmc['srmc_pumped_storage_min']
            too_narrow = spread < 8.0
            if too_narrow.any():
                mid = (srmc['srmc_pumped_storage_min'] + srmc['srmc_pumped_storage_max']) / 2
                srmc.loc[too_narrow, 'srmc_pumped_storage_min'] = (mid - 4.0).clip(lower=params['base_srmc_low'])
                srmc.loc[too_narrow, 'srmc_pumped_storage_max'] = (mid + 4.0).clip(upper=params['base_srmc_high'])
            
            # REQUIRED for price-setting: TECH_SRMC_MAP points hydro_pumped at this
            # column; without it every matching step skips PHS silently. (Re-added
            # 20260929 - dropped inadvertently in the 20260920 PHS rework.)
            srmc['srmc_pumped_storage_avg'] = (srmc['srmc_pumped_storage_min'] +
                                               srmc['srmc_pumped_storage_max']) / 2


            print(f"  ✓ {'pumped_storage':10s}: Vectorized, SRMC-anchored "
                  f"(night={night_factor}, day={day_factor})")
        else:
            raise RuntimeError(f"{zone}: srmc_gas missing - cannot anchor pumped storage")
        
    
        
    # ========== STEP 6: STATIC SRMC FOR RENEWABLES ==========
    # These technologies have minimal variable costs (O&M only)

    # Hydro Run-of-River
    srmc['srmc_hydro_ror_min'] = SRMC_STATIC['hydro_ror'][0]
    srmc['srmc_hydro_ror_max'] = SRMC_STATIC['hydro_ror'][1]
    srmc['srmc_hydro_ror'] = np.mean(SRMC_STATIC['hydro_ror'])

    # Wind Onshore
    srmc['srmc_wind_onshore_min'] = SRMC_STATIC['wind_onshore'][0]
    srmc['srmc_wind_onshore_max'] = SRMC_STATIC['wind_onshore'][1]
    srmc['srmc_wind_onshore'] = np.mean(SRMC_STATIC['wind_onshore'])

    # Wind Offshore
    srmc['srmc_wind_offshore_min'] = SRMC_STATIC['wind_offshore'][0]
    srmc['srmc_wind_offshore_max'] = SRMC_STATIC['wind_offshore'][1]
    srmc['srmc_wind_offshore'] = np.mean(SRMC_STATIC['wind_offshore'])

    # Solar
    srmc['srmc_solar_min'] = SRMC_STATIC['solar'][0]
    srmc['srmc_solar_max'] = SRMC_STATIC['solar'][1]
    srmc['srmc_solar'] = np.mean(SRMC_STATIC['solar'])

    # Waste
    srmc['srmc_waste_min'] = SRMC_STATIC['waste'][0]
    srmc['srmc_waste_max'] = SRMC_STATIC['waste'][1]
    srmc['srmc_waste'] = np.mean(SRMC_STATIC['waste'])

    # Other Renewables
    srmc['srmc_other_renewable_min'] = SRMC_STATIC['other_renewable'][0]
    srmc['srmc_other_renewable_max'] = SRMC_STATIC['other_renewable'][1]
    srmc['srmc_other_renewable'] = np.mean(SRMC_STATIC['other_renewable'])

    print(f"  ✓ {'Renewables':10s}: Static O&M costs applied (hydro RoR, wind, solar, waste)")


    # Clean up year column if it exists
    master_df.drop(columns=['year'], inplace=True, errors='ignore')
    
    # ==========  CRITICAL: RETURN THE SRMC DATAFRAME ==========
    return srmc

# ========== TEST CALCULATION ==========
print("=" * 80)
print("TESTING DYNAMIC SRMC CALCULATION (2015-2026)")
print("=" * 80)

if complete_zones and 'master_test' in locals() and not master_test.empty:
    # Get test zone name
    if 'zone' in master_test.columns:
        test_zone = master_test['zone'].iloc[0]
    else:
        test_zone = complete_zones[0]
    
    print(f"\n{'='*60}")
    print(f"CALCULATING CORRECTED DYNAMIC SRMC FOR {test_zone}")
    print(f"{'='*60}")
    print(f"Date range: {master_test.index.min().date()} to {master_test.index.max().date()}")
    print(f"Years: {master_test.index.year.min()} - {master_test.index.year.max()}")
    
    srmc_test = calculate_srmc_all_techs(master_test, test_zone)
    
    print(f"\nSRMC DataFrame shape: {srmc_test.shape}")
    print(f"Date range: {srmc_test.index.min()} to {srmc_test.index.max()}")
    
    # ========== VALIDATION: SHOW DYNAMIC BEHAVIOR ==========
    print(f"\n{'='*60}")
    print("DYNAMIC SRMC VALIDATION (Gas Example)")
    print(f"{'='*60}")
    
    # Find generation column for gas
    gas_col = None
    for col in master_test.columns:
        if 'gen_gas' in col.lower() or 'fossil_gas' in col.lower():
            gas_col = col
            break
    
    if gas_col:
        # Show correlation between generation and SRMC
        gas_gen = master_test[gas_col]
        gas_srmc = srmc_test['srmc_gas']
        
        # Find hours with min/median/max generation
        active_hours = gas_gen[gas_gen > 0.01]
        if len(active_hours) > 0:
            min_gen_hour = active_hours.idxmin()
            max_gen_hour = active_hours.idxmax()
            median_gen_hour = active_hours.sort_values().index[len(active_hours)//2]
            
            print(f"\nGas Generation vs SRMC (CORRECTED LOGIC):")
            print(f"  MIN generation hour:")
            print(f"    Generation: {gas_gen[min_gen_hour]:.0f} MW")
            print(f"    SRMC: {gas_srmc[min_gen_hour]:.2f} EUR/MWh (should be LOW - best plants)")
            
            print(f"\n  MEDIAN generation hour:")
            print(f"    Generation: {gas_gen[median_gen_hour]:.0f} MW")
            print(f"    SRMC: {gas_srmc[median_gen_hour]:.2f} EUR/MWh (should be MIDDLE)")
            
            print(f"\n  MAX generation hour:")
            print(f"    Generation: {gas_gen[max_gen_hour]:.0f} MW")
            print(f"    SRMC: {gas_srmc[max_gen_hour]:.2f} EUR/MWh (should be HIGH - worst plants) ")
    
    print(f"\n{'='*60}")
    print(" DYNAMIC SRMC CALCULATION COMPLETE")
    print(f"{'='*60}")
    print("\nKey features:")
    print("  - Yearly normalization (generation min/max per year)")
    print("  - Dynamic efficiency: Low gen → high eff, High gen → low eff")
    print("  - Hydro reservoir: Shadow price from filling % (with proxy support)")
    print("  - Pumped storage: Round-trip arbitrage cost")
    print("  - Nuclear: Uranium cost + dynamic efficiency")
    print("  - RES: Static O&M costs")
    print("  - Works across 2015-2026 period")
    print("  - Handles zone transitions (DE_AT_LU, DE_LU, AT)")

else:
    print("\n No master_test available - skipping test calculation")
    print("   Run Cell 10 first to create master_test")


# In[ ]:


#  Cell 12 - SIMPLIFIED 3-STEP MARGINAL TECH IDENTIFICATION (2015-2026)

# STEP 1-4: Direct match (±10 EUR tolerance) - domestic + neighbors
# STEP 4.5: Scarcity pricing check (price > 1.5× gas SRMC)
# STEP 5: Closest BELOW price (no limit) - domestic + neighbors
# STEP 6: Lowest SRMC (last resort) - domestic + neighbors
#  ALL FEATURES: Generation checks, price coupling, neighbor SRMC, pumped hydro logic, scarcity detection
# Updated for 2015-2026 period with zone transitions
# NOW EXTENDS TO 2026

import numpy as np
import pandas as pd

storage_state = {}
print("=" * 80)
print("SIMPLIFIED 3-STEP MARGINAL TECHNOLOGY IDENTIFICATION + SCARCITY CHECK (2015-2026)")
print("=" * 80)

# ========================================
# CONFIGURATION
# ========================================
UNIFORM_TOLERANCE = CFG.UNIFORM_TOLERANCE   # <- control file
PRICE_TOLERANCE = CFG.PRICE_TOLERANCE   # <- control file
FLOW_THRESHOLD = CFG.FLOW_THRESHOLD   # <- control file



def get_generation_threshold(tech_key, zone=None):
    """Get generation threshold with zone-specific overrides."""
    if zone and zone in ZONE_THRESHOLDS:
        if tech_key in ZONE_THRESHOLDS[zone]:
            return ZONE_THRESHOLDS[zone][tech_key]
    return GEN_MIN_THRESHOLDS.get(tech_key, GEN_MIN_THRESHOLDS['default'])

# ========================================
# TECHNOLOGY MAPPING
# ========================================
TECH_SRMC_MAP = {
    'gas': 'srmc_gas', 'coal': 'srmc_coal', 'lignite': 'srmc_lignite',
    'oil': 'srmc_oil', 'oil_shale': 'srmc_oil_shale',  # ADDED oil_shale
    'nuclear': 'srmc_nuclear', 'biomass': 'srmc_biomass',
    'chp': 'srmc_chp', 'hydro_reservoir': 'srmc_hydro_reservoir_avg',
    'hydro_pumped': 'srmc_pumped_storage_avg', 'hydro_ror': 'srmc_hydro_ror',
    'wind_onshore': 'srmc_wind_onshore', 'wind_offshore': 'srmc_wind_offshore',
    'solar': 'srmc_solar', 'waste': 'srmc_waste',
    'other_renewable': 'srmc_other_renewable',
}

# ========================================
# GENERATION CACHE
# ========================================
def build_generation_cache(master_df):
    """Build cache mapping tech → generation column for O(1) lookup."""
    cache = {}
    
    gen_patterns = {
        'coal_derived_gas': ['gen_coal_derived_gas', 'coal_derived_gas', 'coal-derived'],
        'lignite': ['gen_lignite', 'gen_fossil_brown_coal', 'fossil_brown_coal', 'brown_coal'],
        'coal': ['gen_coal', 'gen_fossil_hard_coal', 'fossil_hard_coal'],
        'gas': ['gen_gas', 'gen_fossil_gas'],
        'oil_shale': ['gen_oil_shale', 'oil_shale', 'oilshale'],  # NEW - check BEFORE oil
        'oil': ['gen_oil', 'gen_fossil_oil', 'fossil_oil'],
        'nuclear': ['gen_nuclear', 'nuclear'],
        'wind_offshore': ['gen_wind_offshore', 'wind_offshore'],
        'wind_onshore': ['gen_wind_onshore', 'wind_onshore'],
        'solar': ['gen_solar', 'solar'],
        'biomass': ['gen_biomass', 'biomass'],
        'waste': ['gen_waste', 'waste'],
        'other_renewable': ['gen_other_renewable', 'other_renewable'],
        'hydro_pumped': ['gen_hydro_pumped', 'hydro_pumped_storage', 'pumped_storage'],
        'hydro_reservoir': ['gen_hydro_reservoir', 'hydro_water_reservoir', 'water_reservoir'],
        'hydro_ror': ['gen_hydro_ror', 'hydro_run-of-river', 'run_of_river'],
        'chp': ['gen_chp', 'chp'],
        'other': ['gen_other'],
    }
    
    assigned_columns = set()
    
    for tech, patterns in gen_patterns.items():
        for col in master_df.columns:
            if col in assigned_columns:
                continue
            col_lower = col.lower()
            for pattern in patterns:
                pattern_lower = pattern.lower()
                if pattern_lower in col_lower:
                    if tech == 'gas' and ('coal_derived' in col_lower or 'coal-derived' in col_lower):
                        continue
                    if tech == 'coal' and ('lignite' in col_lower or 'brown' in col_lower or 'coal_derived' in col_lower):
                        continue
                    if tech == 'wind_onshore' and 'offshore' in col_lower:
                        continue
                    if tech == 'oil' and ('oil_shale' in col_lower or 'oilshale' in col_lower):  # NEW
                        continue
                    if tech == 'hydro_reservoir' and 'pumped' in col_lower:
                        continue
                    if tech == 'other' and any(x in col_lower for x in ['renewable', 'chp', 'biomass', 'waste']):
                        continue
                    cache[tech] = col
                    assigned_columns.add(col)
                    break
            if tech in cache:
                break
    
    return cache

def get_generation_cached(row, tech_key, cache):
    """O(1) lookup using pre-built cache with STRICT validation."""
    col = cache.get(tech_key)
    if not col or col not in row.index:
        return 0.0
    val = row.get(col)
    if val is None or pd.isna(val):
        return 0.0
    try:
        float_val = float(val)
    except (ValueError, TypeError):
        return 0.0
    if float_val < 0 or np.isinf(float_val):
        return 0.0
    return float_val

# ========================================
# NEIGHBOR CACHE
# ========================================
def build_neighbour_srmc_cache(zone, neighbours, df_prices, df_gen, df_load, df_flows,
                                gas_prices_df, co2_prices_df, commodity_prices_df, reservoir_data):
    """Build SRMC cache AND master_df cache for all neighbor zones."""
    neighbour_srmc_cache = {}
    neighbour_master_cache = {}
    neighbour_gen_caches = {}
    
    print(f"    Processing {len(neighbours)} neighbors: {neighbours}")
    
    for i, neighbour in enumerate(neighbours, 1):
        print(f"      [{i}/{len(neighbours)}] Processing {neighbour}...")
        try:
            neighbour_master = build_zone_master(
                neighbour, df_prices, df_gen, df_load, df_flows,
                gas_prices_df, co2_prices_df, commodity_prices_df, reservoir_data
            )
            if neighbour_master.empty:
                print(f"           Empty master_df, skipping")
                continue
            
            neighbour_srmc = calculate_srmc_all_techs(neighbour_master, neighbour)
            neighbour_gen_cache = build_generation_cache(neighbour_master)
            
            neighbour_srmc_cache[neighbour] = neighbour_srmc
            neighbour_master_cache[neighbour] = neighbour_master
            neighbour_gen_caches[neighbour] = neighbour_gen_cache
            
            print(f"          ✓ {neighbour} complete!")
        except Exception as e:
            print(f"           ERROR: {type(e).__name__}: {e}")
            import traceback
            print("          Traceback:")
            traceback.print_exc()
            print(f"          Skipping {neighbour} and continuing...")
            continue
    
    print(f"    ✓ Neighbor cache complete: {len(neighbour_srmc_cache)} zones loaded")
    return neighbour_srmc_cache, neighbour_master_cache, neighbour_gen_caches

# Helper to deintify more in detail the decisions made by this algorithm
# Added on 25/06/2026
def scan_closest_below_above(row, srmc_df, zone, gen_cache, coupled_zones,
                             neighbour_srmc_cache, neighbour_master_cache, neighbour_gen_caches,
                             price_home, hour, pumped_storage_allowed, NIGHT_HOURS):
    """
    DIAGNOSTIC ONLY — does NOT affect identification.
    Scan all eligible (generating) technologies, domestic + price-coupled neighbours,
    using the SAME eligibility rules and SRMC convention as STEP 5, and return the
    closest technology strictly BELOW the price and the closest strictly ABOVE it.
    """
    below_tech, below_dist = None, np.inf
    above_tech, above_dist = None, np.inf

    def _consider(label, srmc_compare):
        nonlocal below_tech, below_dist, above_tech, above_dist
        if pd.isna(srmc_compare):
            return
        if srmc_compare < price_home:
            d = price_home - srmc_compare
            if d < below_dist:
                below_dist, below_tech = d, label
        elif srmc_compare > price_home:
            d = srmc_compare - price_home
            if d < above_dist:
                above_dist, above_tech = d, label
        # exactly equal -> a direct hit, neither strictly above nor below; ignore

    # ----- DOMESTIC -----
    for tech_key, srmc_col in TECH_SRMC_MAP.items():
        if srmc_col not in srmc_df.columns:
            continue
        if tech_key == 'hydro_pumped' and not pumped_storage_allowed:
            continue
        gen = get_generation_cached(row, tech_key, gen_cache)
        threshold = get_generation_threshold(tech_key, zone)
        if pd.isna(gen) or gen < threshold:
            continue
        if tech_key == 'solar' and hour in NIGHT_HOURS:
            continue
        srmc_min_col = f'{srmc_col}_min'
        srmc_max_col = f'{srmc_col}_max'
        if srmc_min_col in srmc_df.columns and srmc_max_col in srmc_df.columns:
            srmc_min = srmc_df.loc[row.name, srmc_min_col]
            srmc_max = srmc_df.loc[row.name, srmc_max_col]
            if pd.isna(srmc_min) or pd.isna(srmc_max):
                continue
            srmc_compare = (srmc_min + srmc_max) / 2
        else:
            srmc_compare = srmc_df.loc[row.name, srmc_col]
        _consider(tech_key, srmc_compare)

    # ----- NEIGHBOURS (price-coupled only, mirrors STEP 5) -----
    if len(coupled_zones) > 1:
        for neighbour in coupled_zones:
            if neighbour == zone or neighbour not in neighbour_srmc_cache:
                continue
            neighbour_srmc_df = neighbour_srmc_cache[neighbour]
            neighbour_master = neighbour_master_cache[neighbour]
            neighbour_gen_cache = neighbour_gen_caches[neighbour]
            try:
                neighbour_row = neighbour_master.loc[row.name]
            except KeyError:
                continue
            for tech_key, srmc_col in TECH_SRMC_MAP.items():
                if srmc_col not in neighbour_srmc_df.columns:
                    continue
                neighbour_gen = get_generation_cached(neighbour_row, tech_key, neighbour_gen_cache)
                threshold = get_generation_threshold(tech_key, neighbour)
                if tech_key == 'hydro_pumped':
                    npc = neighbour_row.get('gen_hydro_pumped_consumption', 0)
                    if pd.isna(npc):
                        npc = 0
                    neighbour_pumped_net = neighbour_gen - npc
                    if hour in NIGHT_HOURS or neighbour_pumped_net < threshold:
                        continue
                    neighbour_gen = neighbour_pumped_net
                if pd.isna(neighbour_gen) or neighbour_gen < threshold:
                    continue
                if tech_key == 'solar' and hour in NIGHT_HOURS:
                    continue
                srmc_min_col = f'{srmc_col}_min'
                srmc_max_col = f'{srmc_col}_max'
                if srmc_min_col in neighbour_srmc_df.columns and srmc_max_col in neighbour_srmc_df.columns:
                    srmc_min = neighbour_srmc_df.loc[row.name, srmc_min_col]
                    srmc_max = neighbour_srmc_df.loc[row.name, srmc_max_col]
                    if pd.isna(srmc_min) or pd.isna(srmc_max):
                        continue
                    srmc_compare = (srmc_min + srmc_max) / 2
                else:
                    srmc_compare = neighbour_srmc_df.loc[row.name, srmc_col]
                _consider(f"{tech_key}_{neighbour}", srmc_compare)

    return {
        'diag_closest_below_tech': below_tech,
        'diag_closest_below_dist': (below_dist if below_tech else np.nan),
        'diag_closest_above_tech': above_tech,
        'diag_closest_above_dist': (above_dist if above_tech else np.nan),
    }

# ========================================
# 3-STEP MARGINAL TECH IDENTIFICATION
# ========================================
def identify_marginal_tech_7step_bulletproof_v3(row, master_df, srmc_df, zone, gen_cache, 
                                                neighbours, df_prices, neighbour_srmc_cache,
                                                neighbour_master_cache, neighbour_gen_caches):
    """
    SIMPLIFIED 3-STEP MARGINAL TECH IDENTIFICATION + SCARCITY CHECK
    
    STEP 1-4: Direct match (±10 EUR) - domestic + neighbors
    STEP 4.5: Scarcity pricing check (price > 1.5× gas SRMC)
    STEP 5: Closest BELOW price (no limit) - domestic + neighbors
    STEP 6: Lowest SRMC (last resort) - domestic + neighbors
    
    ALL FEATURES KEPT:
    - Generation checks (must be generating)
    - Price coupling detection
    - Neighbor SRMC checking
    - Pumped hydro logic (net generation)
    - Solar night-time exclusion
    - Dynamic + static SRMC matching
    - Scarcity pricing detection
    """
    
    # ========== STEP 0: VALIDATE INPUT ==========
    price_home = row.get('price', np.nan)
    if pd.isna(price_home):
        return {
            'marginal_techs': [],
            'fallback_mode': 'no_price',
            'high_price_flag': np.nan,
            'price_setting_zone': zone,
            'flow_convergence': zone
        }
    
    # ========== STEP 1: DETECT PRICE COUPLING ==========
    coupled_zones = [zone]
    imports = row.get('imports', 0)
    exports = row.get('exports', 0)
    total_flow = abs(imports) + abs(exports)
    
    if total_flow > FLOW_THRESHOLD:
        for neighbour in neighbours:
            neighbour_price_col = f'price_{neighbour}'
            if neighbour_price_col in master_df.columns:
                neighbour_price = row.get(neighbour_price_col, np.nan)
            elif neighbour_price_col in df_prices.columns:
                neighbour_price = df_prices.loc[row.name, neighbour_price_col]
            else:
                continue
            if pd.isna(neighbour_price):
                continue
            if abs(price_home - neighbour_price) <= PRICE_TOLERANCE:
                coupled_zones.append(neighbour)
    
    flow_convergence = '+'.join(sorted(coupled_zones)) if len(coupled_zones) > 1 else zone
    
    # ========== PRE-FILTER: EXTRACT HOUR ==========
    hour = row.name.hour if hasattr(row.name, 'hour') else 12
    NIGHT_HOURS = [0, 1, 2, 3, 4, 22, 23]
    
    # PUMPED HYDRO: Check NET generation (generation - consumption)
    pumped_gen = get_generation_cached(row, 'hydro_pumped', gen_cache)
    pumped_threshold = get_generation_threshold('hydro_pumped', zone)
    pumped_consumption = row.get('gen_hydro_pumped_consumption', 0)
    if pd.isna(pumped_consumption):
        pumped_consumption = 0
    if not pd.isna(pumped_gen):
        pumped_net = pumped_gen - pumped_consumption
    else:
        pumped_net = -pumped_consumption
    pumped_storage_allowed = (hour not in NIGHT_HOURS and pumped_net >= pumped_threshold)
    
    marginal = []
    fallback_mode = None
    
    # ========== STEP 2-4: DIRECT MATCH (±10 EUR) ==========
    candidates = []
    
    # CHECK DOMESTIC TECHNOLOGIES
    for tech_key, srmc_col in TECH_SRMC_MAP.items():
        if srmc_col not in srmc_df.columns:
            continue
        if tech_key == 'hydro_pumped' and not pumped_storage_allowed:
            continue
        
        # CHECK GENERATION FIRST
        gen = get_generation_cached(row, tech_key, gen_cache)
        threshold = get_generation_threshold(tech_key, zone)
        if pd.isna(gen) or gen < threshold:
            continue
        
        # SOLAR: Exclude at night
        if tech_key == 'solar' and hour in NIGHT_HOURS:
            continue
        
        # TRY DYNAMIC SRMC FIRST
        srmc_val = srmc_df.loc[row.name, srmc_col]
        if pd.notna(srmc_val):
            distance = abs(price_home - srmc_val)
            if distance <= UNIFORM_TOLERANCE:
                unit_surplus = price_home - srmc_val
                candidates.append((tech_key, unit_surplus, distance, srmc_val))
                continue
        
        # TRY STATIC SRMC RANGE
        srmc_min_col = f'{srmc_col}_min'
        srmc_max_col = f'{srmc_col}_max'
        if srmc_min_col in srmc_df.columns and srmc_max_col in srmc_df.columns:
            srmc_min = srmc_df.loc[row.name, srmc_min_col]
            srmc_max = srmc_df.loc[row.name, srmc_max_col]
            if pd.notna(srmc_min) and pd.notna(srmc_max):
                if srmc_min <= price_home <= srmc_max:
                    srmc_mid = (srmc_min + srmc_max) / 2
                    distance = 0.0
                    unit_surplus = price_home - srmc_mid
                    candidates.append((tech_key, unit_surplus, distance, srmc_mid))
    
    # CHECK NEIGHBOR TECHNOLOGIES
    if len(coupled_zones) > 1:
        for neighbour in coupled_zones:
            if neighbour == zone or neighbour not in neighbour_srmc_cache:
                continue
            
            neighbour_srmc_df = neighbour_srmc_cache[neighbour]
            neighbour_master = neighbour_master_cache[neighbour]
            neighbour_gen_cache = neighbour_gen_caches[neighbour]
            
            try:
                neighbour_row = neighbour_master.loc[row.name]
            except KeyError:
                continue
            
            for tech_key, srmc_col in TECH_SRMC_MAP.items():
                if srmc_col not in neighbour_srmc_df.columns:
                    continue
                
                # CHECK GENERATION FIRST
                neighbour_gen = get_generation_cached(neighbour_row, tech_key, neighbour_gen_cache)
                threshold = get_generation_threshold(tech_key, neighbour)
                
                # PUMPED HYDRO: Check NET generation
                if tech_key == 'hydro_pumped':
                    neighbour_pumped_consumption = neighbour_row.get('gen_hydro_pumped_consumption', 0)
                    if pd.isna(neighbour_pumped_consumption):
                        neighbour_pumped_consumption = 0
                    neighbour_pumped_net = neighbour_gen - neighbour_pumped_consumption
                    if hour in NIGHT_HOURS or neighbour_pumped_net < threshold:
                        continue
                    neighbour_gen = neighbour_pumped_net
                
                # FIX: Check generation for ALL neighbor techs
                if pd.isna(neighbour_gen) or neighbour_gen < threshold:
                    continue
                
                
                # SOLAR: Exclude at night
                if tech_key == 'solar' and hour in NIGHT_HOURS:
                    continue
                
                # TRY DYNAMIC SRMC
                neighbour_srmc_val = neighbour_srmc_df.loc[row.name, srmc_col]
                if pd.notna(neighbour_srmc_val):
                    distance = abs(price_home - neighbour_srmc_val)
                    if distance <= UNIFORM_TOLERANCE:
                        unit_surplus = price_home - neighbour_srmc_val
                        tech_label = f"{tech_key}_{neighbour}"
                        candidates.append((tech_label, unit_surplus, distance, neighbour_srmc_val))
                        continue
                
                # TRY STATIC SRMC RANGE
                srmc_min_col = f'{srmc_col}_min'
                srmc_max_col = f'{srmc_col}_max'
                if srmc_min_col in neighbour_srmc_df.columns and srmc_max_col in neighbour_srmc_df.columns:
                    srmc_min = neighbour_srmc_df.loc[row.name, srmc_min_col]
                    srmc_max = neighbour_srmc_df.loc[row.name, srmc_max_col]
                    if pd.notna(srmc_min) and pd.notna(srmc_max):
                        if srmc_min <= price_home <= srmc_max:
                            srmc_mid = (srmc_min + srmc_max) / 2
                            distance = 0.0
                            unit_surplus = price_home - srmc_mid
                            tech_label = f"{tech_key}_{neighbour}"
                            candidates.append((tech_label, unit_surplus, distance, srmc_mid))
    
    # SELECT WINNER (closest match)
    if candidates:
        min_distance = min(d for _, _, d, _ in candidates)
        tied_techs = [t for t, _, d, _ in candidates if d == min_distance]
        marginal = sorted(tied_techs)
        fallback_mode = None
    
    # ========== STEP 4.5: SCARCITY PRICING CHECK ==========
    if not marginal:
        gas_srmc_col = TECH_SRMC_MAP.get('gas')
        
        if gas_srmc_col and gas_srmc_col in srmc_df.columns:
            gas_min_col = f'{gas_srmc_col}_min'
            gas_max_col = f'{gas_srmc_col}_max'
            
            # Get gas SRMC (use average of min/max if available)
            if gas_min_col in srmc_df.columns and gas_max_col in srmc_df.columns:
                gas_srmc = (srmc_df.loc[row.name, gas_min_col] + 
                        srmc_df.loc[row.name, gas_max_col]) / 2
            else:
                gas_srmc = srmc_df.loc[row.name, gas_srmc_col]
            
            if pd.notna(gas_srmc):
                scarcity_threshold = 1.5 * gas_srmc
                
                if price_home > scarcity_threshold:
                    # FIRST: Check if gas is generating in HOME zone
                    gas_gen = get_generation_cached(row, 'gas', gen_cache)
                    gas_threshold = get_generation_threshold('gas', zone)
                    
                    if pd.notna(gas_gen) and gas_gen >= gas_threshold:
                        # Gas is generating in home zone → it's marginal
                        marginal = ['gas']
                        fallback_mode = 'scarcity_pricing'
                    else:
                        # Gas not generating in home zone
                        # Check if price converged with neighbors (import pricing)
                        price_setting_zones = flow_convergence
                        
                        if pd.notna(price_setting_zones) and isinstance(price_setting_zones, str):
                            # Extract neighbor zones (exclude home zone)
                            zones_list = [z.strip() for z in price_setting_zones.split('+')]
                            neighbor_zones = [z for z in zones_list if z != zone]
                            
                            if neighbor_zones:
                                # Price set by gas in neighbor zone (import pricing)
                                neighbor_zone = neighbor_zones[0]
                                marginal = [f'gas_{neighbor_zone}']
                                fallback_mode = 'scarcity_pricing'
                            else:
                                # No neighbors - true domestic scarcity (rare)
                                marginal = ['scarcity']
                                fallback_mode = 'scarcity_pricing'
                        else:
                            # No convergence data - default to scarcity
                            marginal = ['scarcity']
                            fallback_mode = 'scarcity_pricing'
    
    # ========== STEP 5: CLOSEST BELOW PRICE (NO LIMIT) ==========
    if not marginal:
        best_tech, min_distance_below = None, np.inf
        
        # CHECK DOMESTIC TECHNOLOGIES
        for tech_key, srmc_col in TECH_SRMC_MAP.items():
            if srmc_col not in srmc_df.columns:
                continue
            if tech_key == 'hydro_pumped' and not pumped_storage_allowed:
                continue
            
            # CHECK GENERATION FIRST
            gen = get_generation_cached(row, tech_key, gen_cache)
            threshold = get_generation_threshold(tech_key, zone)
            if pd.isna(gen) or gen < threshold:
                continue
            if tech_key == 'solar' and hour in NIGHT_HOURS:
                continue
            
            # GET SRMC
            srmc_min_col = f'{srmc_col}_min'
            srmc_max_col = f'{srmc_col}_max'
            if srmc_min_col in srmc_df.columns and srmc_max_col in srmc_df.columns:
                srmc_min = srmc_df.loc[row.name, srmc_min_col]
                srmc_max = srmc_df.loc[row.name, srmc_max_col]
                if pd.isna(srmc_min) or pd.isna(srmc_max):
                    continue
                srmc_compare = (srmc_min + srmc_max) / 2
            else:
                srmc_compare = srmc_df.loc[row.name, srmc_col]
                if pd.isna(srmc_compare):
                    continue
            
            # ONLY TECHS BELOW PRICE
            if srmc_compare >= price_home:
                continue
            
            distance_below = price_home - srmc_compare
            if distance_below < min_distance_below:
                min_distance_below = distance_below
                best_tech = tech_key
        
        # CHECK NEIGHBOR TECHNOLOGIES
        if len(coupled_zones) > 1:
            for neighbour in coupled_zones:
                if neighbour == zone or neighbour not in neighbour_srmc_cache:
                    continue
                
                neighbour_srmc_df = neighbour_srmc_cache[neighbour]
                neighbour_master = neighbour_master_cache[neighbour]
                neighbour_gen_cache = neighbour_gen_caches[neighbour]
                
                try:
                    neighbour_row = neighbour_master.loc[row.name]
                except KeyError:
                    continue
                
                for tech_key, srmc_col in TECH_SRMC_MAP.items():
                    if srmc_col not in neighbour_srmc_df.columns:
                        continue
                    
                    # CHECK GENERATION FIRST
                    neighbour_gen = get_generation_cached(neighbour_row, tech_key, neighbour_gen_cache)
                    threshold = get_generation_threshold(tech_key, neighbour)
                    
                    # PUMPED HYDRO: Check NET generation
                    if tech_key == 'hydro_pumped':
                        neighbour_pumped_consumption = neighbour_row.get('gen_hydro_pumped_consumption', 0)
                        if pd.isna(neighbour_pumped_consumption):
                            neighbour_pumped_consumption = 0
                        neighbour_pumped_net = neighbour_gen - neighbour_pumped_consumption
                        if hour in NIGHT_HOURS or neighbour_pumped_net < threshold:
                            continue
                        neighbour_gen = neighbour_pumped_net
                    
                    # FIX: Check generation for ALL neighbor techs
                    if pd.isna(neighbour_gen) or neighbour_gen < threshold:
                        continue
                    if tech_key == 'solar' and hour in NIGHT_HOURS:
                        continue
                    
                    # GET SRMC
                    srmc_min_col = f'{srmc_col}_min'
                    srmc_max_col = f'{srmc_col}_max'
                    if srmc_min_col in neighbour_srmc_df.columns and srmc_max_col in neighbour_srmc_df.columns:
                        srmc_min = neighbour_srmc_df.loc[row.name, srmc_min_col]
                        srmc_max = neighbour_srmc_df.loc[row.name, srmc_max_col]
                        if pd.isna(srmc_min) or pd.isna(srmc_max):
                            continue
                        srmc_compare = (srmc_min + srmc_max) / 2
                    else:
                        srmc_compare = neighbour_srmc_df.loc[row.name, srmc_col]
                        if pd.isna(srmc_compare):
                            continue
                    
                    # ONLY TECHS BELOW PRICE
                    if srmc_compare >= price_home:
                        continue
                    
                    distance_below = price_home - srmc_compare
                    if distance_below < min_distance_below:
                        min_distance_below = distance_below
                        best_tech = f"{tech_key}_{neighbour}"
        
        if best_tech:
            marginal = [best_tech]
            fallback_mode = 'closest_below'
    
    # ========== STEP 6: LOWEST SRMC (LAST RESORT) ==========
    if not marginal:
        best_tech, min_srmc = None, np.inf
        
        # CHECK DOMESTIC TECHNOLOGIES
        for tech_key, srmc_col in TECH_SRMC_MAP.items():
            if srmc_col not in srmc_df.columns:
                continue
            if tech_key == 'hydro_pumped' and not pumped_storage_allowed:
                continue
            
            # CHECK GENERATION FIRST
            gen = get_generation_cached(row, tech_key, gen_cache)
            threshold = get_generation_threshold(tech_key, zone)
            if pd.isna(gen) or gen < threshold:
                continue
            if tech_key == 'solar' and hour in NIGHT_HOURS:
                continue
            
            # USE LOWER EDGE
            srmc_min_col = f'{srmc_col}_min'
            if srmc_min_col in srmc_df.columns:
                srmc_lower = srmc_df.loc[row.name, srmc_min_col]
            else:
                srmc_lower = srmc_df.loc[row.name, srmc_col]
            
            if pd.isna(srmc_lower):
                continue
            
            if srmc_lower < min_srmc:
                min_srmc = srmc_lower
                best_tech = tech_key
        
        # CHECK NEIGHBOR TECHNOLOGIES
        if len(coupled_zones) > 1:
            for neighbour in coupled_zones:
                if neighbour == zone or neighbour not in neighbour_srmc_cache:
                    continue
                
                neighbour_srmc_df = neighbour_srmc_cache[neighbour]
                neighbour_master = neighbour_master_cache[neighbour]
                neighbour_gen_cache = neighbour_gen_caches[neighbour]
                
                try:
                    neighbour_row = neighbour_master.loc[row.name]
                except KeyError:
                    continue
                
                for tech_key, srmc_col in TECH_SRMC_MAP.items():
                    if srmc_col not in neighbour_srmc_df.columns:
                        continue
                    
                    # CHECK GENERATION FIRST
                    neighbour_gen = get_generation_cached(neighbour_row, tech_key, neighbour_gen_cache)
                    threshold = get_generation_threshold(tech_key, neighbour)
                    
                    # PUMPED HYDRO: Check NET generation
                    if tech_key == 'hydro_pumped':
                        neighbour_pumped_consumption = neighbour_row.get('gen_hydro_pumped_consumption', 0)
                        if pd.isna(neighbour_pumped_consumption):
                            neighbour_pumped_consumption = 0
                        neighbour_pumped_net = neighbour_gen - neighbour_pumped_consumption
                        if hour in NIGHT_HOURS or neighbour_pumped_net < threshold:
                            continue
                        neighbour_gen = neighbour_pumped_net
                    
                    # FIX: Check generation for ALL neighbor techs
                    if pd.isna(neighbour_gen) or neighbour_gen < threshold:
                        continue
                    if tech_key == 'solar' and hour in NIGHT_HOURS:
                        continue
                    
                    # USE LOWER EDGE
                    srmc_min_col = f'{srmc_col}_min'
                    if srmc_min_col in neighbour_srmc_df.columns:
                        srmc_lower = neighbour_srmc_df.loc[row.name, srmc_min_col]
                    else:
                        srmc_lower = neighbour_srmc_df.loc[row.name, srmc_col]
                    
                    if pd.isna(srmc_lower):
                        continue
                    
                    if srmc_lower < min_srmc:
                        min_srmc = srmc_lower
                        best_tech = f"{tech_key}_{neighbour}"
        
        if best_tech:
            marginal = [best_tech]
            fallback_mode = 'lowest_srmc'
    
    # ========== CALCULATE METADATA ==========
    high_price_flag = np.nan
    if 'srmc_gas' in srmc_df.columns:
        gas_srmc = srmc_df.loc[row.name, 'srmc_gas']
        if pd.notna(gas_srmc) and gas_srmc > 0:
            high_price_flag = price_home / gas_srmc
    
    # ========== CALCULATE PRICE SETTING ZONE (ROBUST FIX) ==========
    # UPDATED FOR 2015-2026
    KNOWN_ZONES = {
        'AT', 'BE', 'BG', 'CH', 'CY', 'CZ', 'EE', 'ES', 'FI', 'FR', 'GB', 'GR', 'HR',
        'HU', 'LT', 'LV', 'NL', 'PL', 'PT', 'RO', 'RS', 'SI', 'SK', 'AL', 'BA', 'ME', 'MK', 'MT', 'UA',
        'DE_AT_LU', 'DE_LU', 'IE_SEM',
        'DK_1', 'DK_2', 'NO_1', 'NO_2', 'NO_3', 'NO_4', 'NO_5', 'SE_1', 'SE_2', 'SE_3', 'SE_4',
        'IT_NORD', 'IT_CNOR', 'IT_CSUD', 'IT_SUD', 'IT_SICI', 'IT_SARD',
        'IT_BRNN', 'IT_FOGN', 'IT_ROSN', 'IT_CALA',
    }

    price_setting_zones = set()

    for tech in marginal:
        zone_found = False
        
        # Try to extract zone from right side
        if '_' in tech:
            # Try 2-part zone first (e.g., DK_1, DE_LU, IT_NORD)
            parts = tech.split('_')
            
            # Check last 2 parts for compound zone
            if len(parts) >= 2:
                potential_zone_2part = f"{parts[-2]}_{parts[-1]}"
                if potential_zone_2part in KNOWN_ZONES:
                    price_setting_zones.add(potential_zone_2part)
                    zone_found = True
            
            # Check last part for simple zone
            if not zone_found and len(parts) >= 2:
                potential_zone_1part = parts[-1]
                if potential_zone_1part in KNOWN_ZONES:
                    price_setting_zones.add(potential_zone_1part)
                    zone_found = True
        
        # If no neighbor zone found, it's a domestic tech
        if not zone_found:
            price_setting_zones.add(zone)

    # Always create the combined zone string
    if price_setting_zones:
        price_setting_region = '+'.join(sorted(price_setting_zones))
    else:
        price_setting_region = zone

#  Second part of the helper to better identify the selection process of the algorthm
# Added on 25/06/2026
# ========== DIAGNOSTIC: CLOSEST-ABOVE vs PICKED (does NOT affect identification) ==========
    diag = {
        'diag_closest_below_tech': None, 'diag_closest_below_dist': np.nan,
        'diag_closest_above_tech': None, 'diag_closest_above_dist': np.nan,
        'diag_picked_label': None, 'diag_swap_diff_fuel': np.nan, 'diag_swap_diff_zone': np.nan,
    }
    if fallback_mode in (None, 'closest_below') and len(marginal) >= 1:
        diag.update(scan_closest_below_above(
            row, srmc_df, zone, gen_cache, coupled_zones,
            neighbour_srmc_cache, neighbour_master_cache, neighbour_gen_caches,
            price_home, hour, pumped_storage_allowed, NIGHT_HOURS
        ))
        diag['diag_picked_label'] = ' + '.join(marginal)
        picked = marginal[0] if len(marginal) == 1 else None
        above = diag['diag_closest_above_tech']
        if picked is not None and above is not None:
            def _split(lbl):
                if '_' in lbl:
                    parts = lbl.split('_')
                    if len(parts) >= 2 and f"{parts[-2]}_{parts[-1]}" in KNOWN_ZONES:
                        return '_'.join(parts[:-2]), f"{parts[-2]}_{parts[-1]}"
                    if parts[-1] in KNOWN_ZONES:
                        return '_'.join(parts[:-1]), parts[-1]
                return lbl, zone
            pf, pz = _split(picked)
            af, az = _split(above)
            diag['diag_swap_diff_fuel'] = (pf != af)
            diag['diag_swap_diff_zone'] = (pz != az)

    return {
        'marginal_techs': marginal,
        'fallback_mode': fallback_mode,
        'high_price_flag': high_price_flag,
        'price_setting_zone': price_setting_region,
        'flow_convergence': flow_convergence,
        **diag,
    }
# ========================================
# APPLY TO ALL HOURS
# ========================================
def apply_7step_analysis_bulletproof_v3(master_df, srmc_df, zone, neighbours, 
                                        df_prices, df_gen, df_load, df_flows,
                                        gas_prices_df, co2_prices_df, commodity_prices_df, 
                                        reservoir_data):
    """Apply simplified 3-step analysis (+ scarcity check) to all hours."""
    
    print(f"  Building generation cache for {zone}...")
    gen_cache = build_generation_cache(master_df)
    print(f"  Cache built: {len(gen_cache)} technologies mapped")
    
    print(f"  Building neighbor SRMC + master_df cache...")
    neighbour_srmc_cache, neighbour_master_cache, neighbour_gen_caches = build_neighbour_srmc_cache(
        zone, neighbours, df_prices, df_gen, df_load, df_flows,
        gas_prices_df, co2_prices_df, commodity_prices_df, reservoir_data
    )
    print(f"  Neighbor cache built: {len(neighbour_srmc_cache)} zones")
    
    results = []
    for idx in master_df.index:
        row = master_df.loc[idx]
        result = identify_marginal_tech_7step_bulletproof_v3(
            row, master_df, srmc_df, zone, gen_cache, neighbours, 
            df_prices, neighbour_srmc_cache, neighbour_master_cache, neighbour_gen_caches
        )
        results.append(result)
    
    results_df = pd.DataFrame(results, index=master_df.index)
    results_df['marginal_tech'] = results_df['marginal_techs'].apply(
        lambda x: ' + '.join(x) if isinstance(x, list) and len(x) > 0 else 'unknown'
    )
    
    return results_df, neighbour_master_cache, neighbour_gen_caches, neighbour_srmc_cache

# ========================================
# EXPAND CO-MARGINAL FOR REPORTING
# ========================================
def expand_comarginal_for_reporting(results_df):
    """Convert co-marginal labels like 'biomass + gas' into separate rows."""
    expanded_rows = []
    for idx, row in results_df.iterrows():
        marginal_tech = row['marginal_tech']
        if ' + ' in marginal_tech:
            techs = [t.strip() for t in marginal_tech.split('+')]
            for tech in techs:
                new_row = row.copy()
                new_row['marginal_tech'] = tech
                new_row['is_comarginal'] = True
                new_row['comarginal_partner'] = marginal_tech
                expanded_rows.append(new_row)
        else:
            new_row = row.copy()
            new_row['is_comarginal'] = False
            new_row['comarginal_partner'] = ''
            expanded_rows.append(new_row)
    return pd.DataFrame(expanded_rows)


# ========================================
# TEST ON SAMPLE ZONE
# ========================================
if 'master_test' in locals() and not master_test.empty and 'srmc_test' in locals() and not srmc_test.empty:
    # ---- SAFEGUARD: every mapped tech must have a usable SRMC column ----20260929
    _missing = [f"{t} -> {c}" for t, c in TECH_SRMC_MAP.items()
                if c not in srmc_test.columns
                and not (f'{c}_min' in srmc_test.columns and f'{c}_max' in srmc_test.columns)]
    if _missing:
        print(f" TECH_SRMC_MAP columns missing from SRMC frame - "
              f"these techs can NEVER set the price: {_missing}")
    # Determine test zone
    if 'zone' in master_test.columns:
        test_zone = master_test['zone'].iloc[0]
    else:
        test_zone = complete_zones[0] if complete_zones else 'UNKNOWN'
    
    print(f"\n{'='*60}")
    print(f"TESTING SIMPLIFIED 3-STEP + SCARCITY ON {test_zone}")
    print(f"{'='*60}")
    print(f"Date range: {master_test.index.min().date()} to {master_test.index.max().date()}")
    print(f"Years: {master_test.index.year.min()} - {master_test.index.year.max()}")
    
    test_neighbours = NEIGHBOURS.get(test_zone, [])
    
    results_7step, neighbour_master_cache, neighbour_gen_caches, neighbour_srmc_cache = apply_7step_analysis_bulletproof_v3(
        master_test, srmc_test, test_zone, test_neighbours, 
        df_prices, df_gen, df_load, df_flows,
        gas_prices_df, co2_prices_df, commodity_prices_df, reservoir_data
    )
    
    results_7step['year'] = results_7step.index.year
    
    identified = (results_7step['marginal_tech'] != 'unknown').sum()
    total = len(results_7step)
    pct = identified / total * 100
    
    print(f"\n{'='*60}")
    print("IDENTIFICATION RESULTS (3-STEP + SCARCITY)")
    print(f"{'='*60}")
    print(f"Total hours: {total:,}")
    print(f"Identified: {identified:,} ({pct:.1f}%)")
    print(f"Unknown: {total - identified:,} ({100-pct:.1f}%)")
    
    print(f"\n{'='*60}")
    print("FALLBACK MODE ANALYSIS")
    print(f"{'='*60}")
    fallback_counts = results_7step['fallback_mode'].value_counts()
    for mode, count in fallback_counts.items():
        mode_label = mode if pd.notna(mode) else 'Direct match (±10 EUR)'
        pct_mode = count / total * 100
        print(f"{mode_label:30s}: {count:>7,} ({pct_mode:>5.1f}%)")
    
    results_7step_expanded = expand_comarginal_for_reporting(results_7step)
    
    print(f"\n{'='*60}")
    print("MARGINAL TECHNOLOGY SHARE")
    print(f"{'='*60}")
    marginal_counts = results_7step_expanded['marginal_tech'].value_counts()
    for tech, count in marginal_counts.head(15).items():
        pct_tech = count / len(results_7step) * 100
        comarginal_count = results_7step_expanded[
            (results_7step_expanded['marginal_tech'] == tech) & 
            (results_7step_expanded['is_comarginal'] == True)
        ].shape[0]
        if comarginal_count > 0:
            print(f"  {tech:30s}: {count:>7,} ({pct_tech:>5.1f}%) [{comarginal_count:,} co-marginal]")
        else:
            print(f"  {tech:30s}: {count:>7,} ({pct_tech:>5.1f}%)")
    
    # NEW: Year-by-year breakdown
    print(f"\n{'='*60}")
    print("YEAR-BY-YEAR IDENTIFICATION RATES")
    print(f"{'='*60}")
    
    for year in sorted(results_7step['year'].unique()):
        year_data = results_7step[results_7step['year'] == year]
        year_identified = (year_data['marginal_tech'] != 'unknown').sum()
        year_total = len(year_data)
        year_pct = (year_identified / year_total * 100) if year_total > 0 else 0
        
        print(f"{year}: {year_identified:>6,} / {year_total:>6,} ({year_pct:>5.1f}%)")
    
    print(f"\n{'='*60}")
    print("SIMPLIFIED 3-STEP + SCARCITY COMPLETE (2015-2026)")
    print(f"{'='*60}")
    print("\nAlgorithm:")
    print("  STEP 1-4: Direct match (±10 EUR)")
    print("  STEP 4.5: Scarcity pricing check (price > 1.5× gas SRMC)")
    print("  STEP 5: Closest BELOW price (no limit)")
    print("  STEP 6: Lowest SRMC (last resort, handles negative prices)")
    print("\nFeatures:")
    print("  Generation checks (must be generating)")
    print("  Price coupling detection")
    print("  Neighbor SRMC checking (domestic + neighbors)")
    print("  Pumped hydro logic (net generation)")
    print("  Solar night-time exclusion")
    print("  Dynamic + static SRMC matching")
    print("  Scarcity pricing detection (extreme price spikes)")
    print("  Handles 2015-2026 period with zone transitions")
    print(f"{'='*60}")

else:
    print("\n master_test or srmc_test not available")
    print("   Run Cells 10 and 11 first")


# In[32]:


cb = results_7step[results_7step['fallback_mode']=='closest_below']
print("closest_below hours:", len(cb))
print("with a closest-above:", cb['diag_closest_above_tech'].notna().sum())
print(cb[['marginal_tech','diag_closest_below_dist','diag_closest_above_tech',
          'diag_closest_above_dist','diag_swap_diff_fuel','diag_swap_diff_zone']].head(10))


# In[ ]:


# Cell 12e - Multi-Zone Processing with Bulletproof Error Handling (2015-2026)
# Updated to handle zone transitions and full date range including 2026
# NOW EXTENDS TO 2026
 
import pandas as pd
import numpy as np
from pathlib import Path
import time
 
print("=" * 80)
print("MULTI-ZONE MARGINAL TECHNOLOGY IDENTIFICATION (2015-2026)")
print("=" * 80)
 
# ========================================
# CONFIGURATION
# ========================================
 
# PHASE 1: Initial zones with transition zones
ZONES_PHASE1 = [
    'DE_AT_LU',  # NEW: 2015-2018 period
    'DE_LU',     # 2019+ period
    'AT',        # 2019+ period (separate from DE_LU)
    'ES', 'FR', 'NL', 'IT_NORD', 'SE_4', 'DK_2', 'PL', 'HU', 'GR'
]
 
ZONES_TO_PROCESS = ZONES_PHASE1
 
print(f"\n📊 Processing {len(ZONES_TO_PROCESS)} zones (2015-2026):")
print(f"   {', '.join(ZONES_TO_PROCESS)}")
 
print(f"\n🔄 Zone transitions:")
print(f"   • DE_AT_LU: 2015-2018 data")
print(f"   • DE_LU: 2019-2026 data")
print(f"   • AT: 2019-2026 data (independent)")
 
# ========================================
# ERROR-SAFE PROCESSING FUNCTION
# ========================================
 
def process_single_zone_safe(zone, df_prices, df_gen, df_load, df_flows,
                              gas_prices_df, co2_prices_df, commodity_prices_df,
                              reservoir_data):
    """
    ERROR-SAFE: Returns None on failure, doesn't crash
    Handles 2015-2026 date range
    Handles zone transitions
    """
    
    print(f"\n{'='*60}")
    print(f"PROCESSING ZONE: {zone}")
    print(f"{'='*60}")
    
    try:
        # ========== STEP 1: BUILD MASTER DATAFRAME ==========
        print(f"  [1/5] Building master DataFrame...")
        
        master = build_zone_master(
            zone, df_prices, df_gen, df_load, df_flows,
            gas_prices_df, co2_prices_df, commodity_prices_df,
            reservoir_data
        )
        
        if master.empty or 'price' not in master.columns:
            print(f"   No price data available for {zone} - SKIPPING")
            return None, None, None, None, None
        
        # Show date range
        date_range = f"{master.index.min().date()} to {master.index.max().date()}"
        years = f"{master.index.year.min()} - {master.index.year.max()}"
        print(f"     ✓ Master DataFrame: {master.shape}")
        print(f"     ✓ Date range: {date_range} ({years})")
        
        # Zone transition validation
        if zone == 'DE_AT_LU':
            max_year = master.index.year.max()
            if max_year > 2018:
                print(f"      WARNING: DE_AT_LU has data after 2018 (max year: {max_year})")
        
        elif zone in ['DE_LU', 'AT']:
            min_year = master.index.year.min()
            if min_year < 2019:
                print(f"      WARNING: {zone} has data before 2019 (min year: {min_year})")
        
        # ========== STEP 2: CALCULATE SRMC ==========
        print(f"  [2/5] Calculating SRMC...")
        
        srmc = calculate_srmc_all_techs(master, zone)
        
        if srmc.empty:
            print(f"   SRMC calculation failed for {zone} - SKIPPING")
            return None, None, None, None, None
        
        print(f"     ✓ SRMC DataFrame: {srmc.shape}")
        
        # ========== STEP 3: RUN 7-STEP ANALYSIS ==========
        print(f"  [3/5] Running 7-step analysis...")
        
        neighbours = NEIGHBOURS.get(zone, [])
        print(f"     Neighbors: {neighbours if neighbours else 'None'}")
        
        results, neighbour_master_cache, neighbour_gen_caches, neighbour_srmc_cache = apply_7step_analysis_bulletproof_v3(
            master, srmc, zone, neighbours,
            df_prices, df_gen, df_load, df_flows,
            gas_prices_df, co2_prices_df, commodity_prices_df,
            reservoir_data
        )
        
        if results.empty:
            print(f"   7-step analysis failed for {zone} - SKIPPING")
            return None, None, None, None, None
        
        # ========== STEP 4: VALIDATE ==========
        identified = (results['marginal_tech'] != 'unknown').sum()
        total = len(results)
        pct = identified / total * 100
        
        print(f"  [4/5] ✓ Identified: {identified:,}/{total:,} hours ({pct:.1f}%)")
        
        # Year-by-year breakdown
        results['year'] = results.index.year
        print(f"\n     Year-by-year identification:")
        for year in sorted(results['year'].unique()):
            year_results = results[results['year'] == year]
            year_identified = (year_results['marginal_tech'] != 'unknown').sum()
            year_total = len(year_results)
            year_pct = (year_identified / year_total * 100) if year_total > 0 else 0
            print(f"       {year}: {year_identified:>6,}/{year_total:>6,} ({year_pct:>5.1f}%)")
        
        # ========== STEP 5: SAVE ==========
        output_file = OUTPUT_DIR / f'results_7step_{zone}_2015_2026.csv'
        results.to_csv(output_file)
        print(f"  [5/5] Saved: {output_file}")
        
        return results, master, srmc, neighbour_master_cache, neighbour_srmc_cache
        
    except Exception as e:
        # CATCH ERROR, LOG IT, CONTINUE TO NEXT ZONE
        print(f"   ERROR in {zone}: {str(e)[:100]}...")
        print(f"     (Full traceback saved to log)")
        
        # Save error log
        error_log = OUTPUT_DIR / f'ERROR_LOG_{zone}.txt'
        with open(error_log, 'w') as f:
            import traceback
            f.write(f"Zone: {zone}\n")
            f.write(f"Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Error: {str(e)}\n\n")
            f.write("Full traceback:\n")
            f.write(traceback.format_exc())
        
        print(f"     Error log saved: {error_log}")
        
        return None, None, None, None, None
 
 
# ========================================
# MAIN PROCESSING LOOP (ERROR-SAFE)
# ========================================
 
all_zone_results = {}
all_zone_masters = {}
all_zone_srmcs = {}
 
processing_summary = {
    'success': [],
    'failed': [],
    'no_data': [],
    'errors': {},
    'timing': {}
}
 
start_time = time.time()
 
print(f"\n{'='*80}")
print("STARTING ERROR-SAFE MULTI-ZONE PROCESSING (2015-2026)")
print(f"{'='*80}")
 
for i, zone in enumerate(ZONES_TO_PROCESS, 1):
    zone_start = time.time()
    
    print(f"\n[{i}/{len(ZONES_TO_PROCESS)}] Processing {zone}...")
    
    # TRY-EXCEPT WRAPPER (EXTRA SAFETY)
    try:
        results, master, srmc, neighbour_master, neighbour_srmc = process_single_zone_safe(
            zone, df_prices, df_gen, df_load, df_flows,
            gas_prices_df, co2_prices_df, commodity_prices_df,
            reservoir_data
        )
        
        zone_elapsed = time.time() - zone_start
        processing_summary['timing'][zone] = zone_elapsed
        
        if results is not None:
            all_zone_results[zone] = results
            all_zone_masters[zone] = master
            all_zone_srmcs[zone] = srmc
            processing_summary['success'].append(zone)
            print(f"  ⏱️ Zone processing time: {zone_elapsed:.1f}s")
        else:
            processing_summary['no_data'].append(zone)
    
    except Exception as e:
        # CATCH-ALL: If even the safe function fails
        print(f"   CRITICAL ERROR in {zone}: {str(e)[:100]}")
        processing_summary['failed'].append(zone)
        processing_summary['errors'][zone] = str(e)
        
        zone_elapsed = time.time() - zone_start
        processing_summary['timing'][zone] = zone_elapsed
 
elapsed = time.time() - start_time
 
# ========================================
# FINAL SUMMARY
# ========================================
 
print(f"\n{'='*80}")
print("MULTI-ZONE PROCESSING COMPLETE (2015-2026)")
print(f"{'='*80}")
 
print(f"\n⏱️ Total processing time: {elapsed:.1f} seconds ({elapsed/60:.1f} minutes)")
print(f"   Average per zone: {elapsed/len(ZONES_TO_PROCESS):.1f} seconds")
 
print(f"\nSuccessfully processed: {len(processing_summary['success'])} zones")
if processing_summary['success']:
    print(f"   {', '.join(processing_summary['success'])}")
    
    # Show timing for successful zones
    print(f"\n   Processing times:")
    for zone in processing_summary['success']:
        timing = processing_summary['timing'].get(zone, 0)
        print(f"     {zone:12s}: {timing:>6.1f}s")
 
if processing_summary['no_data']:
    print(f"\n No data: {len(processing_summary['no_data'])} zones")
    print(f"   {', '.join(processing_summary['no_data'])}")
 
if processing_summary['failed']:
    print(f"\n Failed: {len(processing_summary['failed'])} zones")
    print(f"   {', '.join(processing_summary['failed'])}")
    for zone, error in processing_summary['errors'].items():
        print(f"      - {zone}: {error[:80]}...")
 
# ========================================
# COMBINE SUCCESSFUL RESULTS
# ========================================
 
if all_zone_results:
    print(f"\n{'='*60}")
    print("COMBINING SUCCESSFUL RESULTS")
    print(f"{'='*60}")
    
    # Add 'zone' column to each DataFrame BEFORE concatenating
    results_with_zones = []
    
    for zone, df in all_zone_results.items():
        df_copy = df.copy()
        
        # Add zone column if it doesn't exist
        if 'zone' not in df_copy.columns:
            df_copy['zone'] = zone
        
        results_with_zones.append(df_copy)
    
    # Now concatenate (keep datetime index)
    combined_results = pd.concat(results_with_zones, ignore_index=False)
    
    combined_file = OUTPUT_DIR / 'results_7step_ALL_ZONES_2015_2026.csv'
    combined_results.to_csv(combined_file)
    
    print(f"\nCombined results: {combined_file}")
    print(f"   Total rows: {len(combined_results):,}")
    print(f"   Zones: {combined_results['zone'].nunique()}")
    print(f"   Date range: {combined_results.index.min()} to {combined_results.index.max()}")
    print(f"   Years: {combined_results.index.year.min()} - {combined_results.index.year.max()}")
    
    # Zone breakdown
    print(f"\n   Rows per zone:")
    zone_counts = combined_results['zone'].value_counts().sort_index()
    for zone, count in zone_counts.items():
        pct = count / len(combined_results) * 100
        print(f"     {zone:12s}: {count:>8,} rows ({pct:>5.1f}%)")
    
    # Year breakdown
    print(f"\n   Rows per year:")
    year_counts = combined_results['year'].value_counts().sort_index()
    for year, count in year_counts.items():
        pct = count / len(combined_results) * 100
        zones_in_year = combined_results[combined_results['year'] == year]['zone'].nunique()
        print(f"     {year}: {count:>8,} rows ({pct:>5.1f}%) - {zones_in_year} zones")
 
# ========================================
# SAVE PROCESSING SUMMARY
# ========================================
 
summary_file = OUTPUT_DIR / 'processing_summary_2015_2026.txt'
with open(summary_file, 'w') as f:
    f.write("MULTI-ZONE PROCESSING SUMMARY (2015-2026)\n")
    f.write("=" * 80 + "\n\n")
    f.write(f"Total zones: {len(ZONES_TO_PROCESS)}\n")
    f.write(f"Successful: {len(processing_summary['success'])}\n")
    f.write(f"No data: {len(processing_summary['no_data'])}\n")
    f.write(f"Failed: {len(processing_summary['failed'])}\n\n")
    
    f.write("Successful zones:\n")
    for zone in processing_summary['success']:
        timing = processing_summary['timing'].get(zone, 0)
        f.write(f"  {zone}: {timing:.1f}s\n")
    
    if processing_summary['no_data']:
        f.write("\nNo data zones:\n")
        for zone in processing_summary['no_data']:
            f.write(f"  {zone}\n")
    
    if processing_summary['failed']:
        f.write("\nFailed zones:\n")
        for zone in processing_summary['failed']:
            error = processing_summary['errors'].get(zone, 'Unknown error')
            f.write(f"  {zone}: {error}\n")
 
print(f"\n💾 Processing summary saved: {summary_file}")
 
print("\n" + "=" * 80)
 
 


# In[ ]:


# Cell 12e_EU - Multi-Zone Processing: ALL EU ZONES (Skip Already Processed) (2015-2026)
# Updated to handle zone transitions and continue from Cell 12e
# NOW EXTENDS TO 2026
#  Run Cell 12e FIRST, then run this cell to process remaining zones
 
import pandas as pd
import numpy as np
from pathlib import Path
import time
 
print("=" * 80)
print("MULTI-ZONE MARGINAL TECH IDENTIFICATION - ALL EU ZONES (2015-2026)")
print("=" * 80)
 
# ========================================
# CONFIGURATION - ALL EU ZONES
# ========================================
 
# Check what's already been processed
if 'all_zone_results' in locals() and all_zone_results:
    ZONES_ALREADY_PROCESSED = list(all_zone_results.keys())
    print(f"\nFound {len(ZONES_ALREADY_PROCESSED)} zones already in memory from Cell 12e:")
    print(f"   {', '.join(sorted(ZONES_ALREADY_PROCESSED))}")
else:
    ZONES_ALREADY_PROCESSED = []
    all_zone_results = {}
    all_zone_masters = {}
    all_zone_srmcs = {}
    print(f"\n No zones in memory - starting fresh")
    print(f"   (Tip: Run Cell 12e first to process Phase 1 zones)")
 
# [CFG-OWNED] processing universe = zones sheet (Cell 0: ZONES_EU_ALL = ALL_ZONES)
ZONES_EU_ALL = ALL_ZONES
 

# EU bidding zones ONLY - used for EU weighted averages
# (CH, GB, IE_SEM and NO_1-NO_5 are processed but excluded from EU aggregates)
# [CFG-OWNED, def removed] ZONES_EU_ONLY comes from Cell 0 (zones sheet, eu_scope == TRUE)

# ========================================
# FILTER TO NEW ZONES ONLY
# ========================================
 
def get_available_zones(df_prices, df_gen):
    """
    Identify which zones have data available.
    """
    available_zones = set()
    
    # Check price columns
    for col in df_prices.columns:
        if col.startswith('price_'):
            zone = col.replace('price_', '')
            available_zones.add(zone)
    
    # Check generation columns (look for zone suffixes)
    for col in df_gen.columns:
        for zone in ZONES_EU_ALL:
            if zone in col:
                available_zones.add(zone)
                break
    
    return sorted(list(available_zones))
 
# Get available zones
AVAILABLE_ZONES = get_available_zones(df_prices, df_gen)
 
# FILTER: Only zones that (1) have data AND (2) NOT already processed
ZONES_NEW = [z for z in ZONES_EU_ALL if z in AVAILABLE_ZONES and z not in ZONES_ALREADY_PROCESSED]
 
print(f"\n📊 Total EU zones in list: {len(ZONES_EU_ALL)}")
print(f"📊 Zones with data available: {len(AVAILABLE_ZONES)}")
print(f"📊 Already processed (will skip): {len(ZONES_ALREADY_PROCESSED)}")
 
if ZONES_ALREADY_PROCESSED:
    print(f"\n   Already processed zones:")
    for zone in sorted(ZONES_ALREADY_PROCESSED):
        print(f"     ✓ {zone}")
 
print(f"\nNEW zones to process: {len(ZONES_NEW)}")
 
if ZONES_NEW:
    print(f"\n   NEW zones to process:")
    for i, zone in enumerate(ZONES_NEW, 1):
        # Indicate expected period for transition zones
        if zone == 'DE_AT_LU':
            period = "(2015-2018)"
        elif zone in ['DE_LU', 'AT']:
            period = "(2019-2026)"  # Updated to 2026
        else:
            period = "(2015-2026)"  # Updated to 2026
        print(f"     {i:2d}. {zone:12s} {period}")
else:
    print(f"\n No new zones to process!")
    print(f"   All zones already processed or no data available")
 
ZONES_TO_PROCESS = ZONES_NEW
 
# ========================================
# REUSE PROCESSING FUNCTION FROM CELL 12e
# ========================================
# (Same function as Cell 12e, no need to duplicate)
 
# ========================================
# MAIN PROCESSING LOOP
# ========================================
 
if len(ZONES_TO_PROCESS) == 0:
    print(f"\n{'='*80}")
    print("NO NEW ZONES TO PROCESS")
    print(f"{'='*80}")
    print(f"\nAll {len(all_zone_results)} zones already processed.")
    print(f"\nTo process specific zones, add them to ZONES_TO_PROCESS manually.")
 
else:
    # Track new processing
    processing_summary = {
        'success': [],
        'failed': [],
        'no_data': [],
        'errors': {},
        'timing': {},
        'skipped': ZONES_ALREADY_PROCESSED.copy()
    }
    
    start_time = time.time()
    
    print(f"\n{'='*80}")
    print(f"STARTING MULTI-ZONE PROCESSING ({len(ZONES_TO_PROCESS)} NEW ZONES)")
    print(f"{'='*80}")
    
    for i, zone in enumerate(ZONES_TO_PROCESS, 1):
        zone_start = time.time()
        
        print(f"\n[{i}/{len(ZONES_TO_PROCESS)}] Processing {zone}...")
        
        try:
            results, master, srmc, neighbour_master, neighbour_srmc = process_single_zone_safe(
                zone, df_prices, df_gen, df_load, df_flows,
                gas_prices_df, co2_prices_df, commodity_prices_df,
                reservoir_data
            )
            
            zone_elapsed = time.time() - zone_start
            processing_summary['timing'][zone] = zone_elapsed
            
            if results is not None:
                all_zone_results[zone] = results
                all_zone_masters[zone] = master
                all_zone_srmcs[zone] = srmc
                processing_summary['success'].append(zone)
                print(f"  ⏱️ Zone processing time: {zone_elapsed:.1f}s")
            else:
                processing_summary['no_data'].append(zone)
        
        except Exception as e:
            print(f"   CRITICAL ERROR in {zone}: {str(e)[:100]}")
            processing_summary['failed'].append(zone)
            processing_summary['errors'][zone] = str(e)
            
            zone_elapsed = time.time() - zone_start
            processing_summary['timing'][zone] = zone_elapsed
    
    elapsed = time.time() - start_time
    
    # ========================================
    # SUMMARY
    # ========================================
    
    print(f"\n{'='*80}")
    print("NEW ZONES PROCESSING COMPLETE")
    print(f"{'='*80}")
    
    print(f"\n⏱️ Processing time: {elapsed:.1f} seconds ({elapsed/60:.1f} minutes)")
    print(f"   Average per zone: {elapsed/len(ZONES_TO_PROCESS):.1f} seconds")
    
    print(f"\nNewly processed: {len(processing_summary['success'])} zones")
    if processing_summary['success']:
        print(f"   {', '.join(processing_summary['success'])}")
    
    if processing_summary['no_data']:
        print(f"\n No data: {len(processing_summary['no_data'])} zones")
        print(f"   {', '.join(processing_summary['no_data'])}")
    
    if processing_summary['failed']:
        print(f"\n Failed: {len(processing_summary['failed'])} zones")
        print(f"   {', '.join(processing_summary['failed'])}")
 
# ========================================
# COMBINE ALL RESULTS (OLD + NEW)
# ========================================
 
print(f"\n{'='*80}")
print("TOTAL ZONES IN MEMORY")
print(f"{'='*80}")
 
print(f"\n📊 Total zones processed: {len(all_zone_results)}")
if 'processing_summary' in locals():
    print(f"   Previously in memory: {len(processing_summary['skipped'])}")
    print(f"   Newly processed: {len(processing_summary['success'])}")
 
print(f"\nAll processed zones:")
for zone in sorted(all_zone_results.keys()):
    results_count = len(all_zone_results[zone])
    years = f"{all_zone_results[zone].index.year.min()}-{all_zone_results[zone].index.year.max()}"
    print(f"   {zone:12s}: {results_count:>8,} hours ({years})")
 
if all_zone_results:
    print(f"\n{'='*60}")
    print("COMBINING ALL RESULTS")
    print(f"{'='*60}")
    
    results_with_zones = []
    
    for zone, df in all_zone_results.items():
        df_copy = df.copy()
        if 'zone' not in df_copy.columns:
            df_copy['zone'] = zone
        results_with_zones.append(df_copy)
    
    combined_results = pd.concat(results_with_zones, ignore_index=False)
    
    combined_file = OUTPUT_DIR / 'results_7step_ALL_EU_ZONES_2015_2026.csv'
    combined_results.to_csv(combined_file)
    
    print(f"\nCombined results: {combined_file}")
    print(f"   Total rows: {len(combined_results):,}")
    print(f"   Zones: {combined_results['zone'].nunique()}")
    print(f"   Date range: {combined_results.index.min()} to {combined_results.index.max()}")
    print(f"   Years: {combined_results.index.year.min()}-{combined_results.index.year.max()}")
    
    # Regional summary
    print(f"\n{'='*60}")
    print("REGIONAL SUMMARY")
    print(f"{'='*60}")
    
    regions = {
        'Central Europe': ['AT', 'CH', 'CZ', 'SK', 'PL', 'HU'],
        'Western Europe': ['DE_AT_LU', 'DE_LU', 'FR', 'BE', 'NL'],
        'Iberian Peninsula': ['ES', 'PT'],
        'Italy': ['IT_NORD', 'IT_CNOR', 'IT_CSUD', 'IT_SUD', 'IT_SICI', 'IT_SARD', 'IT_CALA', 'IT_BRNN'],
        'Nordic': ['SE_1', 'SE_2', 'SE_3', 'SE_4', 'NO_1', 'NO_2', 'NO_3', 'NO_4', 'NO_5', 'FI', 'DK_1', 'DK_2'],
        'Baltic': ['EE', 'LV', 'LT'],
        'South-East Europe': ['SI', 'HR', 'RO', 'BG', 'GR'],
        'South-East Europe (non-EU)': ['RS', 'AL', 'ME', 'MK'],
        'Other': ['GB', 'IE_SEM'],
    }
    
    for region, zones in regions.items():
        region_zones = [z for z in zones if z in all_zone_results]
        if region_zones:
            total_hours = sum(len(all_zone_results[z]) for z in region_zones)
            print(f"\n{region}: {len(region_zones)} zones, {total_hours:,} hours")
            print(f"   Zones: {', '.join(region_zones)}")
 
print("\n" + "=" * 80)
print("ALL EU ZONES PROCESSING COMPLETE (2015-2026)")
print("=" * 80)
 
 
# In[300]:
 
 
# SCRATCH v3 - CONTESTED-HOURS SHARE (from all_zone_srmcs)
import re
import numpy as np
import pandas as pd
 
TOL = UNIFORM_TOLERANCE
counts = {}
 
for zone in ZONES_EU_ONLY:
    m = all_zone_masters.get(zone)
    S = all_zone_srmcs.get(zone)
    if m is None or S is None or 'price' not in m.columns:
        continue
    cand = {}
    for c in S.columns:
        mt = re.match(r'srmc_([a-z_]+?)(_avg)?$', c)
        if mt and not c.endswith(('_min', '_max')):
            tech = mt.group(1)
            if mt.group(2) or tech not in cand:
                cand[tech] = c
    if len(cand) < 2:
        continue
    Sv = S[list(cand.values())].reindex(m.index)
    close = Sv.sub(m['price'], axis=0).abs().le(
        np.maximum(TOL, 0.15 * m['price'].abs()), axis=0)
    n_close = close.sum(axis=1)
    valid = m['price'].notna()
    yr = m.index.year
    for year in np.unique(yr[valid]):
        sel = valid & (yr == year)
        d = counts.setdefault(int(year), [0, 0])
        d[0] += int((n_close[sel] >= 2).sum())
        d[1] += int(sel.sum())
 
tab = pd.DataFrame([{'year': y, 'contested_%': round(a / b * 100, 1),
                     'zone_hours': b} for y, (a, b) in sorted(counts.items())])
print(tab.to_string(index=False))
 


# In[ ]:


# Cell 12k_COUNTRY_AGGREGATION - Load-Weighted Price Setting for Multi-Zone Countries (2019-2026)
# Method 2: Annual Load-Weighted (market size weighting) - SAME AS EU-WIDE CELL
# Aggregates countries with multiple bidding zones (DK, SE, IT, NO)
# NOW EXTENDS TO 2026 (but uses 2019 vs 2025 for comparison as 2026 is partial year)
 
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
 
print("=" * 80)
print("MULTI-ZONE COUNTRY AGGREGATION: LOAD-WEIGHTED PRICE SETTING (2019-2026)")
print("Method: Annual Load-Weighted (Market Size) - Same as EU-Wide Analysis")
print("=" * 80)
 
# ========================================
# DEFINE ALL KNOWN ZONE CODES
# ========================================
KNOWN_ZONES = {
    'DE_LU', 'IT_NORD', 'IT_CNOR', 'IT_CSUD', 'IT_SUD', 'IT_SICI', 'IT_SARD',
    'IT_CALA', 'IT_BRNN', 'SE_1', 'SE_2', 'SE_3', 'SE_4', 'DK_1', 'DK_2',
    'NO_1', 'NO_2', 'NO_3', 'NO_4', 'NO_5',
    'FI', 'FR', 'ES', 'PT', 'NL', 'BE', 'AT', 'CH', 'CZ', 'PL', 'HU', 'RO',
    'BG', 'GR', 'HR', 'SI', 'SK', 'EE', 'LV', 'LT', 'GB', 'IE', 'AL', 'MK',
    'RS', 'BA', 'ME', 'UA', 'BY', 'TR'
}
 
# ========================================
# HELPER: EXTRACT BASE TECHNOLOGY
# ========================================
def get_base_tech(tech_string):
    """
    Extract base technology from marginal tech string.
    
    Examples:
    - "gas_DE_LU" → "gas"
    - "coal_IT_NORD" → "coal"
    - "wind_onshore" → "wind_onshore"
    - "gas + coal" → ["gas", "coal"]
    - "gas_DE_LU + coal_NL" → ["gas", "coal"]
    """
    if pd.isna(tech_string) or tech_string == 'unknown' or tech_string == 'scarcity':
        return [tech_string]
    
    # Split by " + " for co-marginal techs
    techs = [t.strip() for t in tech_string.split('+')]
    
    base_techs = []
    for tech in techs:
        # Check if tech ends with any known zone (longest first)
        base_tech = tech
        
        for zone in sorted(KNOWN_ZONES, key=len, reverse=True):
            if tech.endswith(f'_{zone}'):
                # Remove the zone suffix
                base_tech = tech[:-len(zone)-1]  # -1 for the underscore
                break
        
        base_techs.append(base_tech)
    
    return base_techs
 
# ========================================
# MAIN FUNCTION: CALCULATE LOAD-WEIGHTED FUEL MIX
# ========================================
def calculate_country_fuel_mix_aggregated(all_zone_results, country, zones, year=2025):
    """
    Calculate load-weighted fuel mix for multi-zone country.
    AGGREGATES BY BASE TECHNOLOGY (strips zone suffixes).
    
    ANSWERS THE QUESTION:
    "Of all the electricity consumed in [COUNTRY] in [YEAR], 
     what percentage was priced based on gas (or wind, or coal)?"
    
    Method: Annual Load-Weighted (same as EU-wide analysis)
    """
    
    print(f"\n{'='*60}")
    print(f"CALCULATING {country} FUEL MIX (LOAD-WEIGHTED, AGGREGATED)")
    print(f"{'='*60}")
    print(f"Question: Of all electricity consumed in {country} in {year},")
    print(f"          what percentage was priced based on each technology?")
    print(f"\nZones: {zones}")
    print(f"Year: {year}")
    
    # Initialize counters
    tech_load_hours = {}
    total_load_hours = 0
    zone_contributions = {zone: {} for zone in zones}
    
    print(f"\n🔍 Starting zone processing...")
    
    # ========================================
    # LOOP THROUGH EACH ZONE
    # ========================================
    for zone in zones:
        print(f"\n  Processing {zone}...")
        
        if zone not in all_zone_results:
            print(f"     Warning: {zone} not found in all_zone_results")
            continue
        
        df_zone = all_zone_results[zone]
        print(f"    ✓ Got zone results: {len(df_zone)} rows")
        
        df_year = df_zone[df_zone.index.year == year].copy()
        print(f"    ✓ Filtered to year {year}: {len(df_year)} rows")
        
        if df_year.empty:
            print(f"     Warning: No data for {zone} in {year}, skipping")
            continue
        
        if 'marginal_tech' not in df_year.columns:
            print(f"     Warning: marginal_tech column not found for {zone}")
            continue
        
        if zone not in all_zone_masters:
            print(f"     Warning: {zone} not found in all_zone_masters")
            continue
        
        master_zone = all_zone_masters[zone]
        print(f"    ✓ Got master data: {len(master_zone)} rows")
        
        if 'load' not in master_zone.columns:
            print(f"     Warning: load column not found for {zone}")
            continue
        
        master_year = master_zone[master_zone.index.year == year].copy()
        print(f"    ✓ Filtered master to year {year}: {len(master_year)} rows")
        
        df_merged = df_year[['marginal_tech']].join(master_year[['load']], how='inner')
        print(f"    ✓ Merged data: {len(df_merged)} rows")
        
        if df_merged.empty:
            print(f"     Warning: No merged data for {zone} in {year}, skipping")
            continue
        
        # ========================================
        # AGGREGATE BY BASE TECHNOLOGY
        # ========================================
        row_count = 0
        
        for idx, row in df_merged.iterrows():
            marginal_tech_raw = row['marginal_tech']
            load = row['load']
            
            # Skip nan load
            if pd.isna(load):
                continue
            
            # Extract base technologies (handles co-marginal)
            base_techs = get_base_tech(marginal_tech_raw)
            
            # If co-marginal, split load equally among techs
            load_per_tech = load / len(base_techs) if len(base_techs) > 0 else 0
            
            if pd.isna(load_per_tech):
                continue
            
            for base_tech in base_techs:
                if base_tech == 'unknown':
                    continue
                
                # Add to country total
                if base_tech not in tech_load_hours:
                    tech_load_hours[base_tech] = 0.0
                
                tech_load_hours[base_tech] += load_per_tech
                
                # Track zone contribution
                if base_tech not in zone_contributions[zone]:
                    zone_contributions[zone][base_tech] = 0.0
                zone_contributions[zone][base_tech] += load_per_tech
            
            row_count += 1
            if row_count % 1000 == 0:
                print(f"      ... processed {row_count} rows")
        
        # Total load for this zone
        zone_total_load = df_merged['load'].sum()
        total_load_hours += zone_total_load
        
        print(f"    ✓ {zone} complete: {zone_total_load:,.0f} MW-hours")
        print(f"    ✓ Technologies found: {len(zone_contributions[zone])}")
    
    # ========================================
    # DEBUG: Check accumulated data
    # ========================================
    print(f"\n🔍 Aggregation complete")
    print(f"  Total load-hours: {total_load_hours:,.0f}")
    print(f"  Technologies found: {len(tech_load_hours)}")
    
    print(f"\n  All technologies (sorted by MW-hours):")
    for tech, mwh in sorted(tech_load_hours.items(), key=lambda x: -x[1])[:20]:
        pct = (mwh / total_load_hours * 100) if total_load_hours > 0 else 0
        print(f"    {tech:30s}: {mwh:15,.0f} MW-hours ({pct:6.2f}%)")
    
    # ========================================
    # CALCULATE PERCENTAGES
    # ========================================
    if total_load_hours == 0:
        print(f"\n   ERROR: No load data found for {country} in {year}")
        return {}
    
    fuel_mix = {}
    nan_techs = []
    
    for tech, load_hours in tech_load_hours.items():
        if pd.isna(load_hours):
            nan_techs.append(tech)
            print(f"   WARNING: {tech} has NaN MW-hours (skipping)")
        elif load_hours > 0:
            fuel_mix[tech] = (load_hours / total_load_hours) * 100
    
    if nan_techs:
        print(f"\n   TOTAL NaN TECHNOLOGIES: {len(nan_techs)}")
    
    # ========================================
    # SUMMARY OUTPUT
    # ========================================
    print(f"\n{'='*60}")
    print(f"{country} FUEL MIX RESULTS ({year}) - AGGREGATED BY TECHNOLOGY")
    print(f"{'='*60}")
    print(f"Total load-hours: {total_load_hours:,.0f} MW-hours")
    print(f"\nFuel mix (load-weighted, aggregated):")
    print(f"Answer: Of all electricity consumed in {country}, X% was priced by:")
    
    for tech, pct in sorted(fuel_mix.items(), key=lambda x: -x[1]):
        print(f"  {tech:30s}: {pct:6.2f}%")
    
    # Validation
    total_pct = sum(fuel_mix.values())
    print(f"\n  {'Total':30s}: {total_pct:6.2f}% {' if abs(total_pct - 100) < 0.01 else ' ERROR!'}")
    
    # ========================================
    # ZONE BREAKDOWN
    # ========================================
    print(f"\n{'='*60}")
    print(f"ZONE CONTRIBUTIONS (AGGREGATED BY TECHNOLOGY)")
    print(f"{'='*60}")
    
    for zone in zones:
        if zone not in zone_contributions:
            continue
        
        zone_techs = {tech: load for tech, load in zone_contributions[zone].items() if load > 0}
        
        if not zone_techs:
            continue
        
        zone_total = sum(zone_techs.values())
        
        if zone_total == 0:
            continue
        
        zone_pct = (zone_total / total_load_hours * 100) if total_load_hours > 0 else 0
        print(f"\n{zone} ({zone_pct:.1f}% of total load):")
        
        for tech, load_hours in sorted(zone_techs.items(), key=lambda x: -x[1])[:10]:
            tech_pct_of_zone = (load_hours / zone_total * 100) if zone_total > 0 else 0
            tech_pct_of_country = (load_hours / total_load_hours * 100) if total_load_hours > 0 else 0
            print(f"  {tech:25s}: {tech_pct_of_zone:5.1f}% (of zone) | {tech_pct_of_country:5.1f}% (of country)")
    
    print(f"\nReturning fuel_mix with {len(fuel_mix)} technologies")
    return fuel_mix
 
# ========================================
# APPLY TO MULTI-ZONE COUNTRIES (2019-2026)
# ========================================
 
MULTI_ZONE_COUNTRIES = {
    'DK': ['DK_1', 'DK_2'],
    'SE': ['SE_1', 'SE_2', 'SE_3', 'SE_4'],
    'NO': ['NO_1', 'NO_2', 'NO_3', 'NO_4', 'NO_5'],  # Added Norway
    'IT': ['IT_NORD', 'IT_CNOR', 'IT_CSUD', 'IT_SUD', 'IT_SICI', 'IT_SARD']
}
 
# Use 2015 and 2025 for comparison (2026 is partial year)
YEARS = [2015, YEAR_H]
 
country_fuel_mixes_aggregated = {year: {} for year in YEARS}
 
print("\n" + "="*80)
print("MULTI-ZONE COUNTRY AGGREGATION (2019 & 2025) - AGGREGATED BY TECHNOLOGY")
print("="*80)
print("\nCentral Question:")
print("'Of all the electricity consumed in [COUNTRY] in [YEAR],")
print(" what percentage was priced based on gas (or wind, or coal)?'")
print("\nMethod: Annual Load-Weighted (same as EU-wide analysis)")
print(f"Note: Using 2025 for comparison (2026 data available but partial year)")
print("="*80)
 
for year in YEARS:
    print(f"\n\n{'#'*80}")
    print(f"{'#'*80}")
    print(f"PROCESSING YEAR: {year}")
    print(f"{'#'*80}")
    print(f"{'#'*80}")
    
    for country, zones in MULTI_ZONE_COUNTRIES.items():
        try:
            print(f"\n📍 Starting {country}...")
            fuel_mix = calculate_country_fuel_mix_aggregated(
                all_zone_results,
                country, 
                zones, 
                year=year
            )
            
            if fuel_mix:
                country_fuel_mixes_aggregated[year][country] = fuel_mix
                print(f"\n{country} complete: {len(fuel_mix)} technologies")
            else:
                print(f"\n {country} returned empty fuel_mix")
                
        except Exception as e:
            print(f"\n ERROR processing {country} in {year}:")
            print(f"   {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            continue
 
print(f"\n{'='*80}")
print("ALL COUNTRY FUEL MIXES CALCULATED (2019 & 2025) - AGGREGATED")
print(f"{'='*80}")
print(f"   Data available through 2026, but using 2025 for comparison")
 
# Check what we got
print(f"\n🔍 Final results:")
for year in YEARS:
    print(f"\n{year}:")
    for country, fuel_mix in country_fuel_mixes_aggregated[year].items():
        print(f"  {country}: {len(fuel_mix)} technologies")
        top3 = sorted(fuel_mix.items(), key=lambda x: -x[1])[:3]
        print(f"    Top 3: {', '.join([f'{tech} ({pct:.1f}%)' for tech, pct in top3])}")
 
# ========================================
# OPTIONAL: EXPORT TO CSV
# ========================================
print(f"\n{'='*60}")
print("EXPORTING COUNTRY FUEL MIXES TO CSV")
print(f"{'='*60}")
 
export_data = []
for year in YEARS:
    for country, fuel_mix in country_fuel_mixes_aggregated[year].items():
        for tech, pct in fuel_mix.items():
            export_data.append({
                'year': year,
                'country': country,
                'technology': tech,
                'price_setting_share_%': pct
            })
 
if export_data:
    df_export = pd.DataFrame(export_data)
    df_export = df_export.sort_values(['country', 'year', 'price_setting_share_%'], ascending=[True, True, False])
    
    export_file = OUTPUT_DIR / 'multi_zone_countries_fuel_mix_2019_2025.csv'
    df_export.to_csv(export_file, index=False)
    print(f"  Saved: {export_file}")
else:
    print(f"   No data to export")
 
print("\n" + "=" * 80)
print("MULTI-ZONE COUNTRY AGGREGATION COMPLETE (2019-2026 data)")
print("=" * 80)
 


# In[ ]:


# Cell 12k_multi_year - EU-WIDE EVOLUTION: Price-Setting vs Generation Mix (2015-2026)
# Method 2: Annual Load-Weighted (market size weighting)
# All years 2015-2026            <-- was 2019-2026
# 9 REPORTING CATEGORIES         <-- was 6
#
# UPDATED (this version):
#   * YEARS = 2015..2026 (was 2019..2026). All figures / CSV / tables follow.
#   * TECH_CATEGORIES replaced by the 9 categories used across the analysis
#     (Cell 10.2 / SA5 / SA6 / SA7):
#         Coal | Lignite | Gas | Other Fossil Fuels | Hydro Reservoir |
#         Hydro RoR | Nuclear | Biomass | VRE
#     ORDER MATTERS: patterns are matched with startswith, first hit wins, so
#     'coal_derived_gas' must precede 'coal' and 'other_renewable' precede
#     'other' - otherwise coal_derived_gas -> Coal and other_renewable -> the
#     'other' catch-all (both wrong).
#   * get_generation_by_category rewritten to the same 9 categories, driven by
#     one BASE_TO_CAT map so price-setting and generation use IDENTICAL grouping.
#     The old version also silently dropped coal_derived_gas / other_renewable.
#


import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from pathlib import Path

print("=" * 80)
print("EU-WIDE PRICE-SETTING VS GENERATION MIX EVOLUTION (2015-2026)")
print("Method: Annual Load-Weighted (Market Size)")
print("=" * 80)

# ========================================
# CONFIGURATION
# ========================================

from collections import OrderedDict

# ---- 9 reporting categories -------------------------------------------------
# pattern lists are used with startswith(); longest prefixes FIRST.
TECH_CATEGORIES = OrderedDict([
    ('Coal',               ['coal_derived_gas_DUMMY', 'coal', 'coal_']),   # see note below
    ('Lignite',            ['lignite', 'lignite_']),
    ('Gas',                ['gas', 'gas_']),
    ('Other Fossil Fuels', ['coal_derived_gas', 'oil_shale', 'oil', 'chp', 'other_renewable_DUMMY',
                            'other', 'oil_', 'oil_shale_', 'chp_']),
    ('Hydro Reservoir',    ['hydro_reservoir', 'hydro_pumped',
                            'hydro_reservoir_', 'hydro_pumped_']),
    ('Hydro RoR',          ['hydro_ror', 'hydro_ror_']),
    ('Nuclear',            ['nuclear', 'nuclear_']),
    ('Biomass',            ['biomass', 'biomass_']),
    ('VRE',                ['wind_onshore', 'wind_offshore', 'solar', 'waste',
                            'other_renewable', 'wind_', 'solar_', 'waste_']),
])

# base techs, longest first, for stripping a zone suffix off a raw label
_BASE_TECHS_ORDERED = [
    'coal_derived_gas', 'oil_shale', 'other_renewable',
    'hydro_reservoir', 'hydro_ror', 'hydro_pumped',
    'wind_onshore', 'wind_offshore',
    'gas', 'coal', 'lignite', 'nuclear', 'oil',
    'solar', 'biomass', 'waste', 'chp', 'other',
]

COLORS = CATEGORY_COLORS   # palette from control-file 'style' tab (Cell 0)

categories = ['Coal', 'Lignite', 'Gas', 'Other Fossil Fuels',
              'Hydro Reservoir', 'Hydro RoR', 'Nuclear', 'Biomass', 'VRE']

# ALL YEARS 2015-2026
YEARS = list(range(2015, 2027))

# ========================================
# HELPER FUNCTIONS
# ========================================

def _base_tech(label):
    """Strip zone suffix -> base tech (longest prefix wins)."""
    t = str(label).strip().lower()
    for b in _BASE_TECHS_ORDERED:
        if t == b or t.startswith(b):
            return b
    return t

# Modified 20260909 to include fractionality split of co-marginal techs
# Hardened 20260911: helper functions bind their mappings AT DEFINITION TIME
#    (default arguments), so 6-category cells redefining classify_marginal_tech,
#    BASE_TO_CAT or _base_tech later in the session can no longer contaminate them.
def classify_marginal_tech(tech_str, _bt=_base_tech, _map=dict(BASE_TO_CAT)):
    """Classify a single marginal_tech label into one of the 9 categories.
    Co-marginal strings are handled by classify_marginal_tech_frac below."""
    if pd.isna(tech_str) or tech_str == 'unknown' or tech_str == '':
        return 'Unknown'
    return _map.get(_bt(tech_str), 'Unknown')

def classify_marginal_tech_frac(tech_str, _clf=None):
    """Fractional co-marginal classification: returns [(category, weight)], weights summing to 1.
    'wind_onshore + solar + biomass' -> [('VRE', 1/3), ('VRE', 1/3), ('Biomass', 1/3)]."""
    if _clf is None:
        _clf = _CLASSIFY_9_PINNED
    if pd.isna(tech_str) or tech_str == 'unknown' or tech_str == '':
        return [('Unknown', 1.0)]
    parts = [p.strip() for p in str(tech_str).split(' + ') if p.strip()]
    if not parts:
        return [('Unknown', 1.0)]
    w = 1.0 / len(parts)
    return [(_clf(p), w) for p in parts]

_CLASSIFY_9_PINNED = classify_marginal_tech   # module-level pin used by the frac helper

def get_generation_by_category(master_df, _map=dict(BASE_TO_CAT), _cats=tuple(categories)):
    """Total generation by the SAME 9 categories, from master_df columns."""
    category_totals = {cat: 0.0 for cat in _cats}

    # base tech -> candidate column-name patterns
    gen_col_mapping = {
        'coal_derived_gas': ['gen_coal_derived_gas', 'coal_derived_gas'],
        'lignite':          ['gen_lignite', 'gen_fossil_brown_coal'],
        'coal':             ['gen_coal', 'gen_fossil_hard_coal'],
        'gas':              ['gen_gas', 'gen_fossil_gas'],
        'oil_shale':        ['gen_oil_shale'],
        'oil':              ['gen_oil', 'gen_fossil_oil'],
        'chp':              ['gen_chp'],
        'other':            ['gen_other'],
        'nuclear':          ['gen_nuclear'],
        'hydro_reservoir':  ['gen_hydro_reservoir', 'gen_hydro_water_reservoir'],
        'hydro_pumped':     ['gen_hydro_pumped', 'gen_hydro_pumped_storage'],
        'hydro_ror':        ['gen_hydro_ror', 'gen_hydro_run-of-river'],
        'wind_onshore':     ['gen_wind_onshore'],
        'wind_offshore':    ['gen_wind_offshore'],
        'solar':            ['gen_solar'],
        'biomass':          ['gen_biomass'],
        'waste':            ['gen_waste'],
        'other_renewable':  ['gen_other_renewable'],
    }

    used_cols = set()
    for tech, col_patterns in gen_col_mapping.items():
        cat = BASE_TO_CAT.get(tech)
        if cat is None:
            continue
        matched = None
        for pattern in col_patterns:
            for col in master_df.columns:
                if col in used_cols:
                    continue
                cl, pl = col.lower(), pattern.lower()
                if pl not in cl:
                    continue
                # guard against prefix collisions between similar column names
                if tech == 'gas' and ('coal_derived' in cl or 'coal-derived' in cl):
                    continue
                if tech == 'coal' and ('lignite' in cl or 'brown' in cl or 'derived' in cl):
                    continue
                if tech == 'oil' and ('shale' in cl):
                    continue
                if tech == 'other' and any(x in cl for x in
                                           ['renewable', 'chp', 'biomass', 'waste']):
                    continue
                if tech == 'hydro_reservoir' and 'pumped' in cl:
                    continue
                if tech == 'wind_onshore' and 'offshore' in cl:
                    continue
                matched = col
                break
            if matched:
                break
        if matched:
            used_cols.add(matched)
            category_totals[cat] += master_df[matched].fillna(0).sum()

    return category_totals

# ========================================
# MAIN ANALYSIS FUNCTION - ALL YEARS
# ========================================

def create_eu_wide_evolution():
    """EU-wide evolution analysis (2015-2026), annual load-weighted."""

    if not all_zone_results:
        print(" No zone results available. Run data loading first!")
        return None

    available_zones = [z for z in ZONES_EU_ONLY if z in all_zone_results and z in all_zone_masters]

    if not available_zones:
        print(f" No EU zones available")
        return None

    print(f"\n📊 Analyzing {len(available_zones)} EU zones: {', '.join(available_zones)}")
    print(f"📅 Years: {YEARS[0]}-{YEARS[-1]}")

    eu_results = {}

    for year in YEARS:
        print(f"\n{'='*60}")
        print(f"PROCESSING YEAR {year}")
        print(f"{'='*60}")

        zone_annual_loads = {}
        zone_ps_shares = {}
        eu_generation_totals = {cat: 0.0 for cat in categories}

        for zone in available_zones:
            results_df = all_zone_results[zone]
            master_df = all_zone_masters[zone]

            mask = (results_df.index.year == year)
            results_period = results_df[mask]

            mask_master = (master_df.index.year == year)
            master_period = master_df[mask_master]

            if len(results_period) == 0 or len(master_period) == 0:
                continue

            if 'load' not in master_period.columns:
                continue

            zone_annual_load = master_period['load'].sum()
            if pd.isna(zone_annual_load) or zone_annual_load <= 0:
                continue
            zone_annual_loads[zone] = zone_annual_load

            total_hours = len(results_period)
            ps_counts = {cat: 0 for cat in categories}

            # # Modified 20260909 to include fractionality split of co-marginal techs
            for idx, row in results_period.iterrows():
                for category, w in classify_marginal_tech_frac(row['marginal_tech']):
                    if category in ps_counts:
                        ps_counts[category] += w

            zone_ps_shares[zone] = {cat: (count / total_hours * 100)
                                    for cat, count in ps_counts.items()}

            gen_by_category = get_generation_by_category(master_period)
            for cat, gen_mwh in gen_by_category.items():
                eu_generation_totals[cat] += gen_mwh

        total_eu_annual_load = sum(zone_annual_loads.values())

        if total_eu_annual_load == 0:
            print(f"   No load data for {year} - year skipped")
            continue

        eu_ps_annual = {cat: 0.0 for cat in categories}
        for zone, annual_load in zone_annual_loads.items():
            zone_weight = annual_load / total_eu_annual_load
            for cat in categories:
                eu_ps_annual[cat] += zone_ps_shares[zone].get(cat, 0.0) * zone_weight

        eu_ps_annual = {cat: round(val, 1) for cat, val in eu_ps_annual.items()}

        total_gen = sum(eu_generation_totals.values())
        if total_gen > 0:
            eu_gen_shares = {cat: round(val / total_gen * 100, 1)
                             for cat, val in eu_generation_totals.items()}
        else:
            eu_gen_shares = {cat: 0.0 for cat in categories}

        eu_results[year] = {
            'ps_annual': eu_ps_annual,
            'generation': eu_gen_shares,
            'total_annual_load': total_eu_annual_load,
            'total_generation_MWh': total_gen,
            'zones_count': len(zone_annual_loads)
        }

        print(f"  {year} complete:")
        print(f"     Zones: {len(zone_annual_loads)}")
        print(f"     Total annual load: {total_eu_annual_load:,.0f} MWh")
        print(f"     Total generation: {total_gen:,.0f} MWh")
        print(f"     PS sum: {sum(eu_ps_annual.values()):.1f}% "
              f"(remainder = unknown/scarcity)")

    if not eu_results:
        print(" No data collected")
        return None

    years_sorted = sorted(eu_results.keys())
    y0, y1 = min(years_sorted), max(years_sorted)

    # ========================================
    # EXPORT CSV (ALL YEARS)
    # ========================================

    print(f"\n{'='*60}")
    print("EXPORTING CSV TABLE")
    print(f"{'='*60}")

    export_data = []
    for year in years_sorted:
        for cat in categories:
            export_data.append({
                'year': year,
                'category': cat,
                'ps_annual_weighted_%': eu_results[year]['ps_annual'].get(cat, 0.0),
                'generation_share_%': eu_results[year]['generation'].get(cat, 0.0),
                'delta_ps_gen_%': round(eu_results[year]['ps_annual'].get(cat, 0.0) -
                                        eu_results[year]['generation'].get(cat, 0.0), 1),
                'zones': eu_results[year]['zones_count'],
            })

    df_export = pd.DataFrame(export_data)
    df_export['category'] = pd.Categorical(df_export['category'],
                                           categories=categories, ordered=True)
    df_export = df_export.sort_values(['year', 'category']).reset_index(drop=True)

    export_file = OUTPUT_DIR / f'EU_price_setting_vs_generation_{y0}_{y1}_evolution.csv'
    df_export.to_csv(export_file, index=False)
    print(f"  Saved: {export_file}")

    # ========================================
    # EVOLUTION PLOT (one column per year, 2 rows)
    # ========================================

    print(f"\n{'='*60}")
    print("CREATING EU-WIDE EVOLUTION PLOT")
    print(f"{'='*60}")

    n_years = len(years_sorted)

    fig, axes = plt.subplots(2, n_years, figsize=(3.0 * n_years, 10))
    axes = np.atleast_2d(axes)

    fig.suptitle(f'EU-Wide Evolution: Price-Setting vs Generation Mix ({y0}-{y1})\n'
                 f'Method: Annual Load-Weighted | {len(available_zones)} EU zones | 9 categories',
                 fontsize=18, fontweight='bold', y=0.98)

    for col_idx, year in enumerate(years_sorted):
        # ---------- ROW 1: PRICE-SETTING ----------
        ax_ps = axes[0, col_idx]
        bottom = 0
        for cat in categories:
            val = eu_results[year]['ps_annual'].get(cat, 0.0)
            ax_ps.bar(0, val, bottom=bottom, color=COLORS[cat], width=0.6,
                      edgecolor='white', linewidth=1.5)
            if val > 4:
                ax_ps.text(0, bottom + val/2, f'{val:.0f}',
                           ha='center', va='center', fontweight='bold',
                           color='white', fontsize=10)
            bottom += val

        ax_ps.set_ylim(0, 100); ax_ps.set_xlim(-0.5, 0.5); ax_ps.set_xticks([])
        ax_ps.set_title(f'{year}', fontsize=13, fontweight='bold', pad=8)
        ax_ps.set_ylabel('Price-Setting\nShare (%)' if col_idx == 0 else '',
                         fontsize=13, fontweight='bold')
        for s in ['top', 'right', 'bottom']:
            ax_ps.spines[s].set_visible(False)
        ax_ps.grid(axis='y', alpha=0.3, linestyle='--')

        # ---------- ROW 2: GENERATION ----------
        ax_gen = axes[1, col_idx]
        bottom = 0
        for cat in categories:
            val = eu_results[year]['generation'].get(cat, 0.0)
            ax_gen.bar(0, val, bottom=bottom, color=COLORS[cat], width=0.6,
                       edgecolor='white', linewidth=1.5)
            if val > 4:
                ax_gen.text(0, bottom + val/2, f'{val:.0f}',
                            ha='center', va='center', fontweight='bold',
                            color='white', fontsize=10)
            bottom += val

        ax_gen.set_ylim(0, 100); ax_gen.set_xlim(-0.5, 0.5); ax_gen.set_xticks([])
        ax_gen.set_ylabel('Generation\nShare (%)' if col_idx == 0 else '',
                          fontsize=13, fontweight='bold')
        for s in ['top', 'right', 'bottom']:
            ax_gen.spines[s].set_visible(False)
        ax_gen.grid(axis='y', alpha=0.3, linestyle='--')

    handles = [plt.Rectangle((0, 0), 1, 1, color=COLORS[cat], label=cat) for cat in categories]
    fig.legend(handles=handles, loc='lower center', bbox_to_anchor=(0.5, 0.01),
               ncol=min(len(categories), 5), fontsize=12, frameon=False)

    plt.tight_layout(rect=[0, 0.07, 1, 0.96])
    fig_file = OUTPUT_DIR / f'EU_price_setting_vs_generation_{y0}_{y1}_evolution.png'
    plt.savefig(fig_file, dpi=300, bbox_inches='tight')
    print(f"  Saved: {fig_file}")
    plt.show()

    # ========================================
    # TREND LINE CHART
    # ========================================

    print(f"\n{'='*60}")
    print("CREATING TREND LINE CHART")
    print(f"{'='*60}")

    fig2, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10))
    fig2.suptitle(f'EU-Wide Trends: Price-Setting vs Generation Mix ({y0}-{y1})\n'
                  'Method: Annual Load-Weighted | 9 categories',
                  fontsize=18, fontweight='bold')

    for cat in categories:
        ps_values = [eu_results[y]['ps_annual'].get(cat, 0.0) for y in years_sorted]
        ax1.plot(years_sorted, ps_values, marker='o', linewidth=2.5,
                 markersize=7, color=COLORS[cat], label=cat)
    ax1.set_xlabel('Year', fontsize=13, fontweight='bold')
    ax1.set_ylabel('Price-Setting Share (%)', fontsize=13, fontweight='bold')
    ax1.set_title('Price-Setting Technology Trends', fontsize=15, fontweight='bold', pad=15)
    ax1.legend(loc='best', fontsize=10, frameon=True, ncol=2)
    ax1.grid(True, alpha=0.3, linestyle='--')

    for cat in categories:
        gen_values = [eu_results[y]['generation'].get(cat, 0.0) for y in years_sorted]
        ax2.plot(years_sorted, gen_values, marker='s', linewidth=2.5,
                 markersize=7, color=COLORS[cat], label=cat)
    ax2.set_xlabel('Year', fontsize=13, fontweight='bold')
    ax2.set_ylabel('Generation Share (%)', fontsize=13, fontweight='bold')
    ax2.set_title('Generation Mix Trends', fontsize=15, fontweight='bold', pad=15)
    ax2.legend(loc='best', fontsize=10, frameon=True, ncol=2)
    ax2.grid(True, alpha=0.3, linestyle='--')

    plt.tight_layout()
    fig2_file = OUTPUT_DIR / f'EU_price_setting_vs_generation_{y0}_{y1}_trends.png'
    plt.savefig(fig2_file, dpi=300, bbox_inches='tight')
    print(f"  Saved: {fig2_file}")
    plt.show()

    # ========================================
    # COMPARISON TABLE (first vs last year)
    # ========================================

    print(f"\n{'='*104}")
    print(f"EU-WIDE EVOLUTION: ANNUAL LOAD-WEIGHTED ({y0}-{y1})")
    print(f"{'='*104}")
    print(f"  {'Category':<20s} | {str(y0)+' PS':>10s} | {str(y1)+' PS':>10s} | {'Change':>10s} | "
          f"{str(y0)+' Gen':>10s} | {str(y1)+' Gen':>10s} | {'Change':>10s}")
    print(f"  {'-'*104}")

    for cat in categories:
        ps_a = eu_results[y0]['ps_annual'].get(cat, 0.0)
        ps_b = eu_results[y1]['ps_annual'].get(cat, 0.0)
        gen_a = eu_results[y0]['generation'].get(cat, 0.0)
        gen_b = eu_results[y1]['generation'].get(cat, 0.0)
        print(f"  {cat:<20s} | {ps_a:>9.1f}% | {ps_b:>9.1f}% | {ps_b-ps_a:>+9.1f} | "
              f"{gen_a:>9.1f}% | {gen_b:>9.1f}% | {gen_b-gen_a:>+9.1f}")

    # zone-count per year (2015-2018 uses DE_AT_LU; 2019+ uses DE_LU + AT)
    print(f"\n  Contributing zones per year:")
    print("   " + "  ".join(f"{y}:{eu_results[y]['zones_count']}" for y in years_sorted))

    print(f"\n{'='*104}")
    print("EU-WIDE EVOLUTION ANALYSIS COMPLETE")
    print(f"{'='*104}")
    print(f"\n Files saved to: {OUTPUT_DIR}")
    print(f"   - CSV:             {export_file.name}")
    print(f"   - Plot (Stacked):  {fig_file.name}")
    print(f"   - Plot (Trends):   {fig2_file.name}")

    return eu_results

# Modified 20260909 to include fractionality split of co-marginal techs
# ========================================
# RUN ANALYSIS
# ========================================

eu_results = create_eu_wide_evolution()

# Pinned 9-category definitions for downstream cells (export cell, 12Z, etc.)
#    Safe against later 6-category cells overwriting `categories` / `classify_marginal_tech`
CATEGORIES_9    = list(categories)
BASE_TO_CAT_9   = dict(BASE_TO_CAT)
classify_9      = classify_marginal_tech
gen_by_cat_9    = get_generation_by_category
classify_frac_9 = classify_marginal_tech_frac


# In[21]:


# Cell 12k_area - EU STACKED-AREA PAIR: Power Mix vs Price-Setting Mix (9 categories)
# Place AFTER 12k_multi_year (reads its evolution CSV from OUTPUT_DIR).
import matplotlib.pyplot as plt
import pandas as pd
 
df = pd.read_csv(OUTPUT_DIR / 'EU_price_setting_vs_generation_2015_2026_evolution.csv')
 
CAT_ORDER = ['Coal', 'Lignite', 'Other Fossil Fuels', 'Gas', 'Nuclear',
             'Hydro Reservoir', 'Hydro RoR', 'Biomass', 'VRE']
COLORS = CATEGORY_COLORS   # palette from control-file 'style' tab (Cell 0)
 
 
LAST_YEAR = YEAR_H            # set 2026 to include the H1-2026 column
 
d = df[df.year <= LAST_YEAR]
gen = d.pivot_table(index='year', columns='category',
                    values='generation_share_%')[CAT_ORDER]
ps = d.pivot_table(index='year', columns='category',
                   values='ps_annual_weighted_%')[CAT_ORDER]
# display normalization to 100% (price-setting rows exclude unknown/scarcity)
gen_n = gen.div(gen.sum(axis=1), axis=0) * 100
ps_n = ps.div(ps.sum(axis=1), axis=0) * 100
 
fig, axes = plt.subplots(1, 2, figsize=(15, 5.5))
for ax, mat, title, ylab in [
        (axes[0], gen_n, 'POWER MIX', 'Generation (%)'),
        (axes[1], ps_n, 'PRICE SETTING MIX', 'Price-setting (%)')]:
    ax.stackplot(mat.index, [mat[c].values for c in CAT_ORDER],
                 colors=[COLORS[c] for c in CAT_ORDER],
                 labels=CAT_ORDER, edgecolor='white', linewidth=0.4)
    ax.set_title(title, fontsize=15, fontweight='bold', fontstyle='italic',
                 color='#003399', pad=12)
    ax.set_ylabel(ylab, fontsize=11)
    ax.set_xlim(mat.index.min(), mat.index.max())
    ax.set_ylim(0, 100)
    ax.set_xticks(mat.index)
    ax.set_yticks(range(0, 101, 10))
    ax.tick_params(labelsize=10)
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
 
handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, loc='lower center', ncol=5, frameon=False,
           fontsize=10, bbox_to_anchor=(0.5, -0.05))
plt.tight_layout(rect=[0, 0.05, 1, 1])
f = OUTPUT_DIR / f'EU_area_powermix_vs_pricesetting_2015_{LAST_YEAR}.png'
plt.savefig(f, dpi=300, bbox_inches='tight')
plt.show()
print(f'Saved: {f}')
 
 
# In[323]:
 
 
# classify_marginal_tech = lambda s: 'CONTAMINATED'          # simulate a 6-cat cell
# print(classify_frac_9('hydro_ror + solar'))                 # must print RoR + VRE, NOT CONTAMINATED


# In[ ]:


# Cell 12k_export - ZONE-LEVEL & COUNTRY-LEVEL EXPORT: Price-Setting vs Generation (2015-2026, 9 categories)
# Reuses the EXACT definitions of Cell 12k_multi_year (categories, BASE_TO_CAT, classify_marginal_tech,
#    get_generation_by_category) so every number here is consistent with the headline EU evolution CSV.
# Produces two long-format CSVs + two wide pivots, each including the annual-load-weighted EU rows:
#      1) ps_gen_ZONE_level_2015_2026_9cat.csv     - one row per EU bidding zone x year x category
#      2) ps_gen_COUNTRY_level_2015_2026_9cat.csv  - one row per Member State x year x category
#    EU rows are identical in both files (load-weighted mean is associative).
# RUN AFTER Cell 12e_EU and Cell 12k_multi_year (needs their globals in memory).

import pandas as pd
import numpy as np

print("=" * 80)
print("ZONE & COUNTRY LEVEL EXPORT: PRICE-SETTING VS GENERATION (2015-2026, 9 CATEGORIES)")
print("=" * 80)

# ========================================
# PRE-FLIGHT: required globals from earlier cells
# ========================================
# Modified 20260909 to include fractionality split of co-marginal techs
_required = ['all_zone_results', 'all_zone_masters', 'ZONES_EU_ONLY',
             'CATEGORIES_9', 'BASE_TO_CAT_9', 'classify_9', 'gen_by_cat_9', 'classify_frac_9']
_missing = [g for g in _required if g not in globals()]
if _missing:
    raise RuntimeError(f"Missing globals {_missing} - run Cell 12e_EU and Cell 12k_multi_year first.")

# use the pinned 9-category definitions (immune to 6-category cells run later)
categories = CATEGORIES_9

if len(categories) != 9:
    raise RuntimeError(f"Expected the 9-category scheme from Cell 12k_multi_year, found {len(categories)}: {categories}. "
                       "Re-run Cell 12k_multi_year so its definitions are the ones in memory.")

YEARS_EXPORT = list(range(2015, 2027))

# ========================================
# ZONE -> MEMBER STATE MAPPING (EU bidding zones only)
# ========================================

ZONE_TO_MS = {
    'AT': 'Austria', 'BE': 'Belgium', 'BG': 'Bulgaria', 'HR': 'Croatia', 'CZ': 'Czechia',
    'DK_1': 'Denmark', 'DK_2': 'Denmark',
    'EE': 'Estonia', 'FI': 'Finland', 'FR': 'France',
    'DE_LU': 'Germany', 'DE_AT_LU': 'Germany',   # pre-Oct-2018 zone incl. AT & LU -> footnote
    'GR': 'Greece', 'HU': 'Hungary',
    'IT_NORD': 'Italy', 'IT_CNOR': 'Italy', 'IT_CSUD': 'Italy', 'IT_SUD': 'Italy',
    'IT_SICI': 'Italy', 'IT_SARD': 'Italy', 'IT_CALA': 'Italy', 'IT_BRNN': 'Italy',
    'LV': 'Latvia', 'LT': 'Lithuania', 'NL': 'Netherlands', 'PL': 'Poland', 'PT': 'Portugal',
    'RO': 'Romania', 'SK': 'Slovakia', 'SI': 'Slovenia', 'ES': 'Spain',
    'SE_1': 'Sweden', 'SE_2': 'Sweden', 'SE_3': 'Sweden', 'SE_4': 'Sweden',
}

zones_in_scope = [z for z in ZONES_EU_ONLY if z in all_zone_results and z in all_zone_masters]
_unmapped = [z for z in zones_in_scope if z not in ZONE_TO_MS]
if _unmapped:
    print(f" Zones in ZONES_EU_ONLY without a Member State mapping (will be skipped at country level): {_unmapped}")

print(f"\n📊 Zones in scope: {len(zones_in_scope)}")
print(f"   Years: {YEARS_EXPORT[0]}-{YEARS_EXPORT[-1]}")
print(f"   Categories: {categories}")

# ========================================
# STEP 1: ZONE x YEAR RAW STATISTICS
#   ps_counts[cat]  : hours where cat set the price
#   gen_mwh[cat]    : MWh generated by cat
#   hours           : hours in results for that zone-year (denominator of PS shares, incl. Unknown)
#   load_mwh        : annual load (weight for EU / country aggregation)
# ========================================

zone_year = {}   # {(zone, year): dict}

for zone in zones_in_scope:
    results_df = all_zone_results[zone]
    master_df  = all_zone_masters[zone]
    if 'load' not in master_df.columns:
        print(f"   {zone}: no 'load' column - skipped")
        continue

    # Modified 20260909 to include fractionality split of co-marginal techs
    # 20260909: fractional co-marginal classification, exploded to (timestamp, category, weight)
    pairs = results_df['marginal_tech'].map(classify_frac_9)
    frac = pd.DataFrame(
        [(ts, c, w) for ts, lst in pairs.items() for (c, w) in lst],
        columns=['ts', 'cat', 'w']).set_index('ts')

    for year in YEARS_EXPORT:
        m_r = results_df.index.year == year
        m_m = master_df.index.year == year
        if not m_r.any() or not m_m.any():
            continue

        master_period = master_df[m_m]
        load_mwh = master_period['load'].sum()
        if pd.isna(load_mwh) or load_mwh <= 0:
            continue

        # Modified 20260909 to include fractionality split of co-marginal techs
        hours = int(m_r.sum())                       # per-hour count -> denominator unchanged
        frac_year = frac[frac.index.year == year]
        counts = frac_year.groupby('cat')['w'].sum()
        ps_counts = {cat: float(counts.get(cat, 0.0)) for cat in categories}
        unknown_hours = float(counts.get('Unknown', 0.0))

        # Modified 20260909 to include fractionality split of co-marginal techs
        gen_mwh = gen_by_cat_9(master_period)

        zone_year[(zone, year)] = {
            'ps_counts': ps_counts,
            'unknown_hours': unknown_hours,
            'hours': hours,
            'load_mwh': float(load_mwh),
            'gen_mwh': {cat: float(gen_mwh.get(cat, 0.0)) for cat in categories},
        }

print(f"\n✓ Zone-year combinations with data: {len(zone_year)}")

# ========================================
# HELPERS
# ========================================

def _shares_from_counts(ps_counts, hours):
    return {cat: (ps_counts[cat] / hours * 100.0) if hours > 0 else 0.0 for cat in categories}

def _gen_shares(gen_mwh):
    tot = sum(gen_mwh.values())
    return {cat: (gen_mwh[cat] / tot * 100.0) if tot > 0 else 0.0 for cat in categories}

def _aggregate(members):
    """
    Load-weighted aggregation of a list of zone-year stat dicts.
      PS share  = sum_z (load_z * ps_share_z) / sum_z load_z      (annual-load-weighted, = 12k Method 2)
      Gen share = sum_z gen_mwh_z[cat] / sum_z total_gen_z         (MWh-weighted by construction)
    """
    total_load = sum(m['load_mwh'] for m in members)
    ps_w = {cat: 0.0 for cat in categories}
    unk_w = 0.0
    gen_tot = {cat: 0.0 for cat in categories}
    hours_tot = 0
    for m in members:
        w = m['load_mwh'] / total_load if total_load > 0 else 0.0
        sh = _shares_from_counts(m['ps_counts'], m['hours'])
        for cat in categories:
            ps_w[cat] += w * sh[cat]
            gen_tot[cat] += m['gen_mwh'][cat]
        unk_w += w * ((m['unknown_hours'] / m['hours'] * 100.0) if m['hours'] > 0 else 0.0)
        hours_tot += m['hours']
    return {
        'ps_share': ps_w,
        'gen_share': _gen_shares(gen_tot),
        'unknown_share': unk_w,
        'load_mwh': total_load,
        'gen_mwh_total': sum(gen_tot.values()),
        'hours': hours_tot,
        'n_zones': len(members),
    }

def _rows(geography, level, year, agg, zones_label):
    out = []
    for cat in categories:
        ps = agg['ps_share'][cat]
        gs = agg['gen_share'][cat]
        out.append({
            'geography': geography,
            'level': level,
            'year': year,
            'category': cat,
            'ps_share_%': round(ps, 2),
            'gen_share_%': round(gs, 2),
            'delta_ps_gen_%': round(ps - gs, 2),
            'unknown_share_%': round(agg['unknown_share'], 2),
            'hours': agg['hours'],
            'n_zones': agg['n_zones'],
            'zones': zones_label,
            'annual_load_MWh': round(agg['load_mwh'], 0),
            'total_generation_MWh': round(agg['gen_mwh_total'], 0),
        })
    return out

# ========================================
# STEP 2: EU (load-weighted) rows - identical for both files
# ========================================

eu_rows = []
for year in YEARS_EXPORT:
    members = [v for (z, y), v in zone_year.items() if y == year]
    if not members:
        continue
    agg = _aggregate(members)
    zlabel = ', '.join(sorted(z for (z, y) in zone_year if y == year))
    eu_rows += _rows('EU (load-weighted)', 'EU', year, agg, zlabel)

# ========================================
# STEP 3: ZONE-LEVEL FILE
# ========================================

zone_rows = []
for (zone, year), v in sorted(zone_year.items()):
    agg = _aggregate([v])
    zone_rows += _rows(zone, 'zone', year, agg, zone)

df_zone = pd.DataFrame(eu_rows + zone_rows)
df_zone['category'] = pd.Categorical(df_zone['category'], categories=categories, ordered=True)
df_zone = df_zone.sort_values(['level', 'geography', 'year', 'category']).reset_index(drop=True)

# ========================================
# STEP 4: COUNTRY-LEVEL FILE
# ========================================

country_rows = []
for year in YEARS_EXPORT:
    by_ms = {}
    for (zone, y), v in zone_year.items():
        if y != year or zone not in ZONE_TO_MS:
            continue
        by_ms.setdefault(ZONE_TO_MS[zone], []).append((zone, v))
    for ms, items in sorted(by_ms.items()):
        agg = _aggregate([v for _, v in items])
        zlabel = ', '.join(sorted(z for z, _ in items))
        country_rows += _rows(ms, 'country', year, agg, zlabel)

df_country = pd.DataFrame(eu_rows + country_rows)
df_country['category'] = pd.Categorical(df_country['category'], categories=categories, ordered=True)
df_country = df_country.sort_values(['level', 'geography', 'year', 'category']).reset_index(drop=True)

# ========================================
# STEP 5: EXPORT (long + wide)
# ========================================

def _wide(df, value_col):
    w = df.pivot_table(index=['level', 'geography', 'year'], columns='category',
                       values=value_col, observed=True)
    w = w[categories]                      # keep category order
    w.columns = [f'{c} ({value_col})' for c in w.columns]
    meta = df.groupby(['level', 'geography', 'year'], observed=True)[
        ['unknown_share_%', 'hours', 'n_zones', 'zones', 'annual_load_MWh', 'total_generation_MWh']].first()
    return w.join(meta).reset_index()

f_zone_long = OUTPUT_DIR / 'ps_gen_ZONE_level_2015_2026_9cat.csv'
f_ctry_long = OUTPUT_DIR / 'ps_gen_COUNTRY_level_2015_2026_9cat.csv'
f_zone_wide = OUTPUT_DIR / 'ps_gen_ZONE_level_2015_2026_9cat_WIDE.csv'
f_ctry_wide = OUTPUT_DIR / 'ps_gen_COUNTRY_level_2015_2026_9cat_WIDE.csv'

df_zone.to_csv(f_zone_long, index=False)
df_country.to_csv(f_ctry_long, index=False)

# wide: price-setting and generation side by side per geography-year
wz = _wide(df_zone, 'ps_share_%').merge(
    _wide(df_zone, 'gen_share_%')[['level', 'geography', 'year'] +
                                  [f'{c} (gen_share_%)' for c in categories]],
    on=['level', 'geography', 'year'])
wc = _wide(df_country, 'ps_share_%').merge(
    _wide(df_country, 'gen_share_%')[['level', 'geography', 'year'] +
                                     [f'{c} (gen_share_%)' for c in categories]],
    on=['level', 'geography', 'year'])
wz.to_csv(f_zone_wide, index=False)
wc.to_csv(f_ctry_wide, index=False)



# Added on 20260909 for weekly newsletter export
# ========================================
# STEP 5b: "FORMAT WEEKLY" EXPORT (BIDDING-ZONE LEVEL)
#   rows = bidding zone x 9 technologies (short names), cols = years, values = ps_share_%
#   includes the EU (load-weighted) block; missing zone-years filled with 0.00
# ========================================

CAT_TO_SHORT = {
    'Coal': 'coal',
    'Lignite': 'lignite',
    'Gas': 'gas',
    'Hydro Reservoir': 'hydro_res',
    'Hydro RoR': 'hydro_ror',
    'Nuclear': 'nuclear',
    'Other Fossil Fuels': 'other_fossil',
    'Biomass': 'biomass',
    'VRE': 'renewables',
}
SHORT_ORDER = ['coal', 'lignite', 'gas', 'hydro_res', 'hydro_ror',
               'nuclear', 'other_fossil', 'biomass', 'renewables']

fw = df_zone.copy()
fw['Technology'] = fw['category'].map(CAT_TO_SHORT)
fw_piv = fw.pivot_table(index=['geography', 'Technology'], columns='year',
                        values='ps_share_%', observed=True)
fw_piv = fw_piv.reindex(columns=YEARS_EXPORT).fillna(0.0).round(2)

# order: EU block first, then zones alphabetically; techs in fixed order
geos = ['EU (load-weighted)'] + sorted(g for g in fw['geography'].unique()
                                       if g != 'EU (load-weighted)')
fw_piv = fw_piv.reindex(pd.MultiIndex.from_product([geos, SHORT_ORDER],
                                                   names=['Zone', 'Technology']))
fw_out = fw_piv.reset_index()

f_weekly = OUTPUT_DIR / f'ps_ZONE_format_weekly_{YEARS_EXPORT[0]}_{YEARS_EXPORT[-1]}.csv'
fw_out.to_csv(f_weekly, index=False)
print(f"  Format weekly (zones): {f_weekly.name}  ({len(fw_out)} rows = "
      f"{len(geos)} zones x {len(SHORT_ORDER)} technologies)")




print(f"\n💾 Saved:")
for f, d in [(f_zone_long, df_zone), (f_ctry_long, df_country)]:
    print(f"   {f.name:50s} {len(d):>7,} rows  "
          f"({d[d['level']!='EU']['geography'].nunique()} geographies + EU, "
          f"{d['year'].nunique()} years)")
print(f"   {f_zone_wide.name:50s} {len(wz):>7,} rows")
print(f"   {f_ctry_wide.name:50s} {len(wc):>7,} rows")

# ========================================
# STEP 6: CONSISTENCY CHECKS
# ========================================

print(f"\n{'='*80}")
print("CONSISTENCY CHECKS")
print(f"{'='*80}")

# (a) EU rows identical in both files
eu_z = df_zone[df_zone['level'] == 'EU'].set_index(['year', 'category'])['ps_share_%']
eu_c = df_country[df_country['level'] == 'EU'].set_index(['year', 'category'])['ps_share_%']
print(f"  EU rows identical in zone & country files: {eu_z.equals(eu_c)}")

# (b) EU rows vs headline 12k_multi_year CSV (if present)
evo_file = OUTPUT_DIR / 'EU_price_setting_vs_generation_2015_2026_evolution.csv'
if evo_file.exists():
    evo = pd.read_csv(evo_file)
    if {'year', 'category', 'ps_annual_weighted_%'} <= set(evo.columns):
        evo_s = evo.set_index(['year', 'category'])['ps_annual_weighted_%']
        cmp = pd.concat([eu_z.rename('export'), evo_s.rename('12k_evolution')], axis=1).dropna()
        cmp['diff'] = (cmp['export'] - cmp['12k_evolution']).abs()
        print(f"  vs 12k_multi_year evolution CSV: max |diff| = {cmp['diff'].max():.2f} pp "
              f"({'OK - same method' if cmp['diff'].max() <= 0.15 else 'CHECK - larger than rounding'})")
    else:
        print("  12k evolution CSV found but columns differ - skipped comparison")
else:
    print("  12k evolution CSV not found - skipped comparison")

# (c) coverage warnings the reader must footnote
print(f"\n  Zone-years with < 90% of hours (partial data):")
part = df_zone[(df_zone['level'] == 'zone') & (df_zone['category'] == categories[0])]
for _, r in part.iterrows():
    full = 4344 if r['year'] == 2026 else (8784 if r['year'] in (2016, 2020, 2024) else 8760)
    if r['hours'] < 0.9 * full:
        print(f"    {r['geography']:10s} {int(r['year'])}: {int(r['hours']):>5,} h ({r['hours']/full*100:.0f}%)")

print(f"\n  Zones per year (EU rows): " +
      ", ".join(f"{int(y)}:{int(n)}" for y, n in
                df_zone[(df_zone['level']=='EU') & (df_zone['category']==categories[0])][['year','n_zones']].values))

print(f"\n📌 Notes for the paper:")
print(f"   • Germany 2015-2018 = DE_AT_LU (incl. Austria & Luxembourg); Austria appears from 2019")
print(f"   • 2026 = H1 only (and fewer zones if 2026 inputs are incomplete) - footnote")
print(f"   • PS shares use ALL hours as denominator; 'unknown_share_%' shows the unclassified remainder")
print("\n" + "=" * 80)


# In[325]:


# In[23]:


# Cell 12k_fossil_lines - FOSSIL SHARE OF PRICE-SETTING: ES, GR, IT_NORD (2015-2025)
# Place AFTER Cell 12k_export (reads ps_gen_ZONE_level_2015_2026_9cat.csv).
import matplotlib.pyplot as plt
import pandas as pd

zdf = pd.read_csv(OUTPUT_DIR / 'ps_gen_ZONE_level_2015_2026_9cat.csv')

FOSSIL_CATS = ['Coal', 'Lignite', 'Gas', 'Other Fossil Fuels']
ZONES_PLOT = ['IT_NORD', 'GR', 'ES']
LAST_YEAR = YEAR_H

STYLE = {   # color, linestyle, marker - mirroring the attached EC look
    'IT_NORD': dict(color='#808080', linestyle='-', marker='s'),
    'GR':      dict(color='#1f6079', linestyle='--', marker='o'),
    'ES':      dict(color='#ffc000', linestyle='--', marker='^'),
}

sub = zdf[(zdf.geography.isin(ZONES_PLOT)) & (zdf.year <= LAST_YEAR)]
# fossil share within ATTRIBUTED hours (renormalized over the 9 categories)
tot = (sub.groupby(['geography', 'year'])['ps_share_%'].sum()
          .rename('total').reset_index())
fos = (sub[sub.category.isin(FOSSIL_CATS)]
       .groupby(['geography', 'year'])['ps_share_%'].sum()
       .rename('fossil').reset_index())
t = fos.merge(tot, on=['geography', 'year'])
t['fossil_share_%'] = t['fossil'] / t['total'] * 100

fig, ax = plt.subplots(figsize=(8.2, 5.6))
for z in ZONES_PLOT:
    s = t[t.geography == z].sort_values('year')
    ax.plot(s.year, s['fossil_share_%'], label=z, linewidth=2.6,
            markersize=7, **STYLE[z])

ax.set_title('Spain, Greece, and Italy', fontsize=16, loc='left',
             color='#404040', pad=14)
ax.set_ylabel('Fossil fuels in price-setting (%)', fontsize=12)
ax.set_ylim(0, 100)
ax.set_yticks(range(0, 101, 10))
ax.set_xticks(sorted(t.year.unique()))
ax.tick_params(labelsize=11)
ax.grid(False)
for s_ in ('top', 'right'):
    ax.spines[s_].set_visible(False)
ax.legend(loc='lower center', bbox_to_anchor=(0.5, -0.22), ncol=3,
          frameon=False, fontsize=11)
plt.tight_layout()
f = OUTPUT_DIR / f'fossil_share_ps_ES_GR_ITNORD_2015_{LAST_YEAR}.png'
plt.savefig(f, dpi=300, bbox_inches='tight')
plt.show()

print(f'Saved: {f}')
print(t.pivot_table(index='geography', columns='year',
                    values='fossil_share_%').round(1).to_string())


# In[ ]:


# Cell 10.2: Analysis - Domestic vs Neighbour Price Setting (Country-Level Perspective)
# UPDATED: Main graph shows neighbour imports broken down by technology categories
 
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np
from collections import OrderedDict
 
print("=" * 80)
print("DOMESTIC VS NEIGHBOUR PRICE SETTING ANALYSIS (2025) - COUNTRY PERSPECTIVE")
print("=" * 80)
 
# ========================================
# DEFINE TECHNOLOGY GROUPS
# ========================================
 
TECH_CATEGORIES = OrderedDict([
    ('Coal', ['coal', 'coal_']),
    ('Lignite', ['lignite', 'lignite_']),
    ('Gas', ['gas', 'gas_']),
    ('Other Fossil Fuels', ['oil', 'oil_shale', 'chp', 'oil_', 'oil_shale_', 'chp_']),
    ('Hydro Reservoir', ['hydro_reservoir', 'hydro_pumped', 'hydro_reservoir_', 'hydro_pumped_']),
    ('Hydro RoR', ['hydro_ror', 'hydro_ror_']),
    ('Nuclear', ['nuclear', 'nuclear_']),
    ('Biomass', ['biomass', 'biomass_', 'waste', 'waste_']),
    ('VRE', ['wind_onshore', 'wind_offshore', 'solar', 'other_renewable',
             'wind_', 'solar_']),
])
 
COLORS = CATEGORY_COLORS   # palette from control-file 'style' tab (Cell 0)
 
categories = list(TECH_CATEGORIES.keys())
 
def classify_technology(tech_str):
    """Classify technology into category."""
    if pd.isna(tech_str) or tech_str == 'unknown' or tech_str == '':
        return 'Unknown'
    
    if ' + ' in tech_str:
        tech_str = tech_str.split(' + ')[0].strip()
    
    tech_lower = tech_str.lower()
    
    for category, patterns in TECH_CATEGORIES.items():
        for pattern in patterns:
            if tech_lower.startswith(pattern):
                return category
    
    return 'Unknown'
 
print(f"\n📊 Technology Categories:")
for _c in list(TECH_CATEGORIES.keys()):
    print(f"   • {_c}")
 
# ========================================
# DEFINE COUNTRY GROUPINGS
# ========================================
 
COUNTRY_ZONES = {
    'AT': ['AT'],
    'BE': ['BE'],
    'BG': ['BG'],
    'CZ': ['CZ'],
    'DE': ['DE_LU'],
    'DK': ['DK_1', 'DK_2'],
    'EE': ['EE'],
    'ES': ['ES'],
    'FI': ['FI'],
    'FR': ['FR'],
    'GR': ['GR'],
    'HR': ['HR'],
    'HU': ['HU'],
    'IT': ['IT_NORD', 'IT_CNOR', 'IT_CSUD', 'IT_SUD', 'IT_SICI', 'IT_SARD', 'IT_CALA', 'IT_BRNN', 'IT_FOGN', 'IT_ROSN'],
    'LT': ['LT'],
    'LV': ['LV'],
    'NL': ['NL'],
    'PL': ['PL'],
    'PT': ['PT'],
    'RO': ['RO'],
    'SE': ['SE_1', 'SE_2', 'SE_3', 'SE_4'],
    'SI': ['SI'],
    'SK': ['SK'],
}
 
# Create reverse mapping: zone → country
ZONE_TO_COUNTRY = {}
for country, zones in COUNTRY_ZONES.items():
    for zone in zones:
        ZONE_TO_COUNTRY[zone] = country
 
print(f"\nCountry groupings defined for {len(COUNTRY_ZONES)} countries/markets")
 
# ========================================
# HELPER FUNCTION: EXTRACT COUNTRIES FROM PRICE_SETTING_ZONE
# ========================================
 
def extract_countries_from_price_setting(price_setting_zone_str):
    """Extract unique countries from price_setting_zone string."""
    if pd.isna(price_setting_zone_str) or price_setting_zone_str == '':
        return []
    
    # Split by '+'
    zones = [z.strip() for z in price_setting_zone_str.split('+')]
    
    # Map to countries
    countries = set()
    for zone in zones:
        if zone in ZONE_TO_COUNTRY:
            countries.add(ZONE_TO_COUNTRY[zone])
        else:
            # Unknown zone (e.g., GB, CH, NO_X) - treat as own country
            if '_' in zone:
                country_code = zone.split('_')[0]
            else:
                country_code = zone
            countries.add(country_code)
    
    return sorted(list(countries))
 
# ========================================
# CALCULATE SHARES FOR EACH COUNTRY (2025 ONLY)
# NEW: Track neighbour hours by tech category
# ========================================
 
results = []
neighbour_tech_details = []
 
# Get all EU zones
NON_EU_ZONES = ['GB', 'CH', 'NO_1', 'NO_2', 'NO_3', 'NO_4', 'NO_5', 'DE_AT_LU', 'IE_SEM', 'UA', 'ME', 'AL', 'BA', 'MK', 'RS']
EU_ZONES = [z for z in complete_zones if z not in NON_EU_ZONES]
 
# Process by country
for country, country_zones in COUNTRY_ZONES.items():
    # Skip historical zone
    if country == 'DE_AT_LU':
        continue
    
    # Skip if no zones in EU_ZONES
    zones_to_analyze = [z for z in country_zones if z in EU_ZONES]
    if not zones_to_analyze:
        continue
    
    print(f"\n📍 Processing {country} ({len(zones_to_analyze)} zones: {', '.join(zones_to_analyze)})")
    
    # ========================================
    # COLLECT DATA FROM ALL ZONES
    # ========================================
    
    zone_dataframes = {}
    
    for zone in zones_to_analyze:
        if zone not in all_zone_results:
            continue
        
        results_df = all_zone_results[zone]
        
        if 'year' not in results_df.columns:
            continue
        
        if 2025 not in results_df['year'].values:
            continue
        
        df_2025 = results_df[results_df['year'] == YEAR_H].copy()
        
        # Add country classification for this zone
        df_2025['setting_countries'] = df_2025['price_setting_zone'].apply(extract_countries_from_price_setting)
        df_2025['zone'] = zone
        
        zone_dataframes[zone] = df_2025
    
    if not zone_dataframes:
        print(f"   No 2025 data available")
        continue
    
    # ========================================
    # SYNCHRONIZED HOUR ANALYSIS WITH TECH CATEGORY TRACKING
    # ========================================
    
    # Get common timestamps across all zones
    all_timestamps = None
    for zone, df in zone_dataframes.items():
        if all_timestamps is None:
            all_timestamps = set(df.index)
        else:
            all_timestamps = all_timestamps.intersection(set(df.index))
    
    all_timestamps = sorted(list(all_timestamps))
    
    print(f"  📅 Analyzing {len(all_timestamps)} common timestamps")
    
    domestic_hours = 0
    neighbour_hours_by_cat = {cat: 0 for cat in categories}
    neighbour_hours_by_cat['Other'] = 0
    
    for timestamp in all_timestamps:
        # Check ALL zones for this timestamp
        has_neighbour_influence = False
        timestamp_tech_categories = set()
        
        for zone, df in zone_dataframes.items():
            if timestamp not in df.index:
                continue
            
            setting_countries = df.loc[timestamp, 'setting_countries']
            
            # Check if this zone has neighbour influence
            if setting_countries != [country]:
                has_neighbour_influence = True
                
                # Get the marginal tech and classify it
                marginal_tech = df.loc[timestamp, 'marginal_tech']
                tech_category = classify_technology(marginal_tech)
                timestamp_tech_categories.add(tech_category)
        
        # CLASSIFY THE ENTIRE HOUR
        if has_neighbour_influence:
            # Determine primary tech category for this hour
            # Priority: if multiple categories, take the first non-Unknown
            primary_cat = 'Other'
            for cat in categories:
                if cat in timestamp_tech_categories:
                    primary_cat = cat
                    break
            
            if primary_cat == 'Other' and 'Unknown' in timestamp_tech_categories:
                primary_cat = 'Other'
            
            neighbour_hours_by_cat[primary_cat] += 1
            
            # TRACK NEIGHBOUR TECHNOLOGIES WITH CLASSIFICATION
            for zone, df in zone_dataframes.items():
                if timestamp not in df.index:
                    continue
                
                row = df.loc[timestamp]
                marginal_tech = row.get('marginal_tech', '')
                
                if pd.isna(marginal_tech) or marginal_tech == 'unknown':
                    continue
                
                # Handle co-marginal
                if ' + ' in marginal_tech:
                    techs = marginal_tech.split(' + ')
                else:
                    techs = [marginal_tech]
                
                for tech in techs:
                    tech = tech.strip()
                    
                    # Classify technology
                    tech_category = classify_technology(tech)
                    
                    # Check if tech has zone suffix (neighbour tech)
                    if '_' in tech:
                        parts = tech.split('_')
                        
                        # Check for 2-part zones
                        if len(parts) >= 2:
                            potential_zone = f"{parts[-2]}_{parts[-1]}"
                            if potential_zone in ALL_ZONES:
                                neighbour_country = ZONE_TO_COUNTRY.get(potential_zone, potential_zone)
                                if neighbour_country != country:
                                    tech_key = '_'.join(parts[:-2]) if len(parts) > 2 else parts[0]
                                    neighbour_tech_details.append({
                                        'country': country,
                                        'home_zone': zone,
                                        'timestamp': timestamp,
                                        'neighbour_country': neighbour_country,
                                        'neighbour_zone': potential_zone,
                                        'technology': tech_key,
                                        'tech_category': tech_category,
                                        'full_tech': tech
                                    })
                                continue
                        
                        # Check for 1-part zones
                        if parts[-1] in ALL_ZONES:
                            neighbour_country = ZONE_TO_COUNTRY.get(parts[-1], parts[-1])
                            if neighbour_country != country:
                                tech_key = '_'.join(parts[:-1])
                                neighbour_tech_details.append({
                                    'country': country,
                                    'home_zone': zone,
                                    'timestamp': timestamp,
                                    'neighbour_country': neighbour_country,
                                    'neighbour_zone': parts[-1],
                                    'technology': tech_key,
                                    'tech_category': tech_category,
                                    'full_tech': tech
                                })
        else:
            domestic_hours += 1
    
    total_hours = len(all_timestamps)
    total_neighbour_hours = sum(neighbour_hours_by_cat.values())
    
    domestic_share = (domestic_hours / total_hours * 100) if total_hours > 0 else 0
    neighbour_shares_by_cat = {cat: (hours / total_hours * 100) if total_hours > 0 else 0 
                               for cat, hours in neighbour_hours_by_cat.items()}
    
    print(f"  Domestic: {domestic_share:.1f}% ({domestic_hours:,} hours)")
    print(f"  Neighbour Total: {(total_neighbour_hours/total_hours*100):.1f}% ({total_neighbour_hours:,} hours)")
    for cat in categories:
        if neighbour_shares_by_cat[cat] > 0:
            print(f"     • {cat}: {neighbour_shares_by_cat[cat]:.1f}%")
    
    result_row = {
        'country': country,
        'zones': ', '.join(zones_to_analyze),
        'num_zones': len(zones_to_analyze),
        'domestic_2025': domestic_share,
        'total_hours_2025': int(total_hours),
        'domestic_hours_2025': int(domestic_hours),
    }
    
    # Add neighbour shares by category
    for cat in categories:
        result_row[f'neighbour_{cat.lower().replace(" ", "_")}_2025'] = neighbour_shares_by_cat[cat]
        result_row[f'neighbour_{cat.lower().replace(" ", "_")}_hours'] = int(neighbour_hours_by_cat[cat])
    
    # Total neighbour
    result_row['neighbour_total_2025'] = sum(neighbour_shares_by_cat[cat] for cat in categories)
    result_row['neighbour_total_hours'] = int(total_neighbour_hours)
    
    results.append(result_row)
 
# Create DataFrames
df_country_analysis = pd.DataFrame(results)
df_neighbour_techs = pd.DataFrame(neighbour_tech_details)
 
if len(df_country_analysis) == 0:
    print("\n ERROR: No countries processed.")
else:
    # Sort by domestic share (descending)
    df_country_analysis = df_country_analysis.sort_values('domestic_2025', ascending=False).reset_index(drop=True)
 
    # ---- CSV EXPORT of Visualization 1 (domestic + neighbour-by-tech) ----
    from pathlib import Path as _Path
    _out = _Path(globals().get('PLOTS_DIR') or globals().get('OUTPUT_DIR') or '.')
    _cols = ['country', 'domestic_2025', 'domestic_hours_2025']
    for _cat in categories:
        _key = _cat.lower().replace(" ", "_")
        _cols += [f'neighbour_{_key}_2025', f'neighbour_{_key}_hours']
    _cols += ['neighbour_total_2025', 'neighbour_total_hours']
    _cols = [c for c in _cols if c in df_country_analysis.columns]
    _csv = _out / 'domestic_vs_neighbour_by_tech_2025.csv'
    df_country_analysis[_cols].to_csv(_csv, index=False)
    print(f"📄 Saved graph-1 data: {_csv.name}  ({len(df_country_analysis)} countries)")
    
    print(f"\n{'='*80}")
    print(f"Analysis complete for {len(df_country_analysis)} EU countries")
    print(f"{'='*80}")
    
    # ========================================
    # MAIN VISUALIZATION: DOMESTIC VS NEIGHBOUR BY TECH CATEGORY (2025)
    # NEW: Stacked bar with domestic + neighbour breakdown by category
    # ========================================
    
    print(f"\n{'='*80}")
    print("CREATING VISUALIZATION 1: DOMESTIC VS NEIGHBOUR BY TECH CATEGORY (2025)")
    print(f"{'='*80}")
    
    fig1 = go.Figure()
    
    countries = df_country_analysis['country'].values
    
    # Add Domestic bar (bottom layer)
    fig1.add_trace(
        go.Bar(
            name='Domestic',
            x=countries,
            y=df_country_analysis['domestic_2025'],
            marker_color='#1f4e78',
            hovertemplate=(
                '<b>%{x}</b><br>' +
                'Domestic: %{y:.1f}%<br>' +
                'Hours: %{customdata}<br>' +
                '<extra></extra>'
            ),
            customdata=df_country_analysis['domestic_hours_2025'],
            text=df_country_analysis['domestic_2025'].apply(lambda x: f"{x:.0f}%" if x >= 3 else ""),
            textposition='inside',
            textfont=dict(color='white', size=10, family='Arial Black')
        )
    )
    
    # Add Neighbour bars by category (stacked)
    for cat in categories:
        col_name = f'neighbour_{cat.lower().replace(" ", "_")}_2025'
        col_hours = f'neighbour_{cat.lower().replace(" ", "_")}_hours'
        
        fig1.add_trace(
            go.Bar(
                name=f'Neighbour: {cat}',
                x=countries,
                y=df_country_analysis[col_name],
                marker_color=COLORS[cat],
                hovertemplate=(
                    '<b>%{x}</b><br>' +
                    f'{cat}: %{{y:.1f}}%<br>' +
                    'Hours: %{customdata}<br>' +
                    '<extra></extra>'
                ),
                customdata=df_country_analysis[col_hours],
                text=df_country_analysis[col_name].apply(lambda x: f"{x:.0f}%" if x >= 2 else ""),
                textposition='inside',
                textfont=dict(color='white', size=9, family='Arial Black')
            )
        )
    
    # Update layout
    fig1.update_layout(
        title={
            'text': '<b>Share of Price-Setting Hours: Domestic vs Neighbour by Technology (2025)</b><br><sub>Country-Level Perspective - Neighbour imports split by technology category</sub>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#333', 'family': 'Arial'}
        },
        barmode='stack',
        height=700,
        width=1500,
        hovermode='x unified',
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family="Arial", size=11, color='#333'),
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02,
            font=dict(size=11)
        ),
        xaxis=dict(
            title=dict(text='<b>Country</b>', font=dict(size=13, color='#333')),
            tickfont=dict(size=11),
            showgrid=False,
            showline=True,
            linewidth=1,
            linecolor='#333'
        ),
        yaxis=dict(
            title=dict(text='<b>Share of Hours (%)</b>', font=dict(size=13, color='#333')),
            range=[0, 100],
            showgrid=True,
            gridwidth=1,
            gridcolor='#e0e0e0',
            tickfont=dict(size=11)
        )
    )
    
    fig1.show()
    
    # Save
    output_file1 = OUTPUT_DIR / 'domestic_vs_neighbour_by_tech_country_2025_EU.html'
    fig1.write_html(output_file1)
    print(f"Saved: {output_file1}")
    
    # ========================================
    # VIZ 2: TOP 15 COUNTRIES BY NEIGHBOUR INFLUENCE (WITH BREAKDOWN)
    # ========================================
    
    print(f"\n{'='*80}")
    print("CREATING VISUALIZATION 2: TOP 15 COUNTRIES BY NEIGHBOUR INFLUENCE")
    print(f"{'='*80}")
    
    top15_neighbour = df_country_analysis.nlargest(15, 'neighbour_total_2025').copy()
    
    fig2 = go.Figure()
    
    # Stack neighbour categories
    for cat in categories:
        col_name = f'neighbour_{cat.lower().replace(" ", "_")}_2025'
        
        fig2.add_trace(
            go.Bar(
                name=cat,
                x=top15_neighbour['country'],
                y=top15_neighbour[col_name],
                marker_color=COLORS[cat],
                text=top15_neighbour[col_name].apply(lambda x: f"{x:.0f}%" if x >= 2 else ""),
                textposition='inside',
                textfont=dict(color='white', size=9, family='Arial Black'),
                hovertemplate=f'<b>%{{x}}</b><br>{cat}: %{{y:.1f}}%<br><extra></extra>'
            )
        )
    
    fig2.update_layout(
        title={
            'text': '<b>Top 15 EU Countries by Neighbour Price Setting (2025)',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#333'}
        },
        barmode='stack',
        height=600,
        width=1200,
        plot_bgcolor='white',
        paper_bgcolor='white',
        xaxis=dict(
            title='<b>Country</b>',
            tickfont=dict(size=12)
        ),
        yaxis=dict(
            title='<b>Neighbour Share (%)</b>',
            showgrid=True,
            gridcolor='#e0e0e0'
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    fig2.show()
    
    output_file2 = OUTPUT_DIR / 'top15_neighbour_by_tech_country_2025_EU.html'
    fig2.write_html(output_file2)
    print(f"Saved: {output_file2}")
    
    # [Keep VIZ 3, 4, 5 the same as before - they analyze the detail data]
    
    # ========================================
    # VIZ 3: TOP NEIGHBOUR TECHNOLOGY CATEGORIES (from detail data)
    # ========================================
    
    if len(df_neighbour_techs) > 0:
        print(f"\n{'='*80}")
        print("CREATING VISUALIZATION 3: NEIGHBOUR TECHNOLOGY CATEGORIES (DETAIL)")
        print(f"{'='*80}")
        
        cat_counts = df_neighbour_techs['tech_category'].value_counts()
        total_events = cat_counts.sum()
        cat_percentages = (cat_counts / total_events * 100)
        
        cat_counts_sorted = pd.Series({cat: cat_counts.get(cat, 0) for cat in categories})
        cat_percentages_sorted = pd.Series({cat: cat_percentages.get(cat, 0) for cat in categories})
        
        bar_colors = [COLORS.get(cat, '#999999') for cat in cat_counts_sorted.index]
        
        fig3 = go.Figure()
        
        fig3.add_trace(
            go.Bar(
                x=cat_counts_sorted.index,
                y=cat_percentages_sorted.values,
                marker_color=bar_colors,
                text=cat_percentages_sorted.apply(lambda x: f"{x:.0f}%"),
                textposition='outside',
                hovertemplate=(
                    '<b>%{x}</b><br>' +
                    'Share: %{y:.0f}%<br>' +
                    'Occurrences: %{customdata:,}<br>' +
                    '<extra></extra>'
                ),
                customdata=cat_counts_sorted.values
            )
        )
        
        fig3.update_layout(
            title={
                'text': '<b>Neighbour Technology Categories Setting Prices (2025)</b>',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18, 'color': '#333'}
            },
            height=600,
            width=1200,
            plot_bgcolor='white',
            paper_bgcolor='white',
            xaxis=dict(
                title='<b>Technology Category</b>',
                tickfont=dict(size=11),
                tickangle=-45
            ),
            yaxis=dict(
                title='<b>Share of Events (%)</b>',
                showgrid=True,
                gridcolor='#e0e0e0'
            )
        )
        
        fig3.show()
        
        output_file3 = OUTPUT_DIR / 'neighbour_tech_categories_detail_country_2025_EU.html'
        fig3.write_html(output_file3)
        print(f"Saved: {output_file3}")
        
        # VIZ 4 & 5 same as before...
        # ========================================
# VIZ 4: TOP 15 EXPORTING COUNTRIES BY TECH CATEGORY
# NEW: Shows which countries export their prices most, broken down by tech
# ========================================
 
if len(df_neighbour_techs) > 0:
    print(f"\n{'='*80}")
    print("CREATING VISUALIZATION 4: TOP EXPORTING COUNTRIES BY TECH CATEGORY")
    print(f"{'='*80}")
    
    # Count exports by neighbour country and tech category
    export_matrix = df_neighbour_techs.groupby(['neighbour_country', 'tech_category']).size().reset_index(name='count')
    
    # Get total exports per country
    country_totals = export_matrix.groupby('neighbour_country')['count'].sum().sort_values(ascending=False)
    top15_exporters = country_totals.head(15).index
    
    # Filter to top 15 exporters
    export_matrix_top15 = export_matrix[export_matrix['neighbour_country'].isin(top15_exporters)]
    
    # Pivot to get categories as columns
    export_pivot = export_matrix_top15.pivot(index='neighbour_country', columns='tech_category', values='count').fillna(0)
    
    # Ensure all categories exist
    for cat in categories:
        if cat not in export_pivot.columns:
            export_pivot[cat] = 0
    
    # Reorder columns to match category order
    export_pivot = export_pivot[categories]
    
    # Sort by total (descending)
    export_pivot['Total'] = export_pivot.sum(axis=1)
    export_pivot = export_pivot.sort_values('Total', ascending=False).drop('Total', axis=1)
    
    # Convert to percentages (% of each country's exports)
    export_pivot_pct = export_pivot.div(export_pivot.sum(axis=1), axis=0) * 100
    
    # Get absolute counts for hover
    export_pivot_abs = export_pivot.copy()
    
    fig4 = go.Figure()
    
    # Stack bars by category
    for cat in categories:
        fig4.add_trace(
            go.Bar(
                name=cat,
                x=export_pivot_pct.index,
                y=export_pivot_pct[cat],
                marker_color=COLORS[cat],
                text=export_pivot_pct[cat].apply(lambda x: f"{x:.0f}%" if x >= 3 else ""),
                textposition='inside',
                textfont=dict(color='white', size=9, family='Arial Black'),
                hovertemplate=(
                    '<b>%{x}</b><br>' +
                    f'{cat}: %{{y:.1f}}%<br>' +
                    'Events: %{customdata:,.0f}<br>' +
                    '<extra></extra>'
                ),
                customdata=export_pivot_abs[cat].values
            )
        )
    
    # Add annotations for total exports
    total_exports = export_pivot.sum(axis=1)
    
    fig4.update_layout(
        title={
            'text': '<b>Top 15 Price-Exporting Countries by Technology Category (2025)</b>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#333'}
        },
        barmode='stack',
        height=650,
        width=1400,
        plot_bgcolor='white',
        paper_bgcolor='white',
        xaxis=dict(
            title='<b>Exporting Country</b>',
            tickfont=dict(size=11)
        ),
        yaxis=dict(
            title='<b>Share of Price-Setting Exports (%)</b>',
            showgrid=True,
            gridcolor='#e0e0e0',
            range=[0, 100]
        ),
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02,
            font=dict(size=11)
        )
    )
    
    # Add total export counts as annotations above bars
    for i, country in enumerate(export_pivot_pct.index):
        total = total_exports[country]
        fig4.add_annotation(
            x=i,
            y=105,
            text=f"{total:,.0f}",
            showarrow=False,
            font=dict(size=9, color='#333'),
            xanchor='center'
        )
    
    fig4.show()
    
    output_file4 = OUTPUT_DIR / 'top15_exporting_countries_by_tech_2025_EU.html'
    fig4.write_html(output_file4)
    print(f"Saved: {output_file4}")
    
    # ========================================
    # SUMMARY TABLE: TOP EXPORTERS
    # ========================================
    
    print(f"\n{'='*130}")
    print(f"TOP 10 PRICE-EXPORTING COUNTRIES (2025)")
    print(f"{'='*130}\n")
    
    _h2 = f"  {'Rank':<6s} {'Country':<8s} {'Total':<10s} | "
    _h2 += " ".join(f"{cat[:9]:<10s}" for cat in categories)
    print(_h2)
    print(f"  {'-'*130}")
    
    for i, country in enumerate(export_pivot.head(10).index, 1):
        total = export_pivot.loc[country].sum()
        _pcts = " ".join(f"{export_pivot_pct.loc[country, cat]:>8.1f}%" for cat in categories)
        # Count countries affected
        countries_affected = df_neighbour_techs[df_neighbour_techs['neighbour_country'] == country]['country'].nunique()
        print(f"  {i:<6d} {country:<8s} {total:>9,.0f} | {_pcts}")
        print(f"         (affects {countries_affected} countries)")
    
    print(f"\n{'='*130}")
 
# ========================================
# VIZ 5: TOP 15 NEIGHBOUR COUNTRIES - MAGNITUDE OF EXPORTS BY TECH CATEGORY
# UPDATED: Shows share of each tech category's total exports
# ========================================
 
if len(df_neighbour_techs) > 0:
    print(f"\n{'='*80}")
    print("CREATING VISUALIZATION 5: TOP NEIGHBOUR COUNTRIES - EXPORT MAGNITUDE BY TECH")
    print(f"{'='*80}")
    
    # Count exports per country and tech category (deduplicated by timestamp)
    df_neighbour_unique = df_neighbour_techs.drop_duplicates(subset=['country', 'neighbour_country', 'timestamp', 'tech_category'])
    
    export_by_cat = df_neighbour_unique.groupby(['neighbour_country', 'tech_category']).size().reset_index(name='hours')
    
    # Get top 15 exporters by total
    top15_total = df_neighbour_unique['neighbour_country'].value_counts().head(15)
    top15_countries = top15_total.index
    
    # Filter to top 15
    export_by_cat_top15 = export_by_cat[export_by_cat['neighbour_country'].isin(top15_countries)]
    
    # Pivot to get categories as columns
    export_pivot = export_by_cat_top15.pivot(index='neighbour_country', columns='tech_category', values='hours').fillna(0)
    
    # Ensure all categories exist
    for cat in categories:
        if cat not in export_pivot.columns:
            export_pivot[cat] = 0
    
    # Reorder to match category order
    export_pivot = export_pivot[categories]
    
    # Sort by total (descending)
    export_pivot['Total'] = export_pivot.sum(axis=1)
    export_pivot = export_pivot.sort_values('Total', ascending=False)
    
    total_hours = export_pivot['Total'].copy()
    export_pivot_abs = export_pivot.drop('Total', axis=1).copy()
    
    # Calculate percentage of each CATEGORY's total (e.g., % of all Gas exports)
    category_totals = export_pivot_abs.sum(axis=0)  # Total hours per category across all countries
    export_pivot_cat_pct = export_pivot_abs.div(category_totals, axis=1) * 100  # % of each category
    
    # Get countries affected for each exporter
    countries_affected = []
    for nc in export_pivot_abs.index:
        affected = df_neighbour_unique[df_neighbour_unique['neighbour_country'] == nc]['country'].nunique()
        countries_affected.append(affected)
    
    fig5 = go.Figure()
    
    # Stack bars by category
    for cat in categories:
        fig5.add_trace(
            go.Bar(
                name=cat,
                x=export_pivot_abs.index,
                y=export_pivot_abs[cat],
                marker_color=COLORS[cat],
                text=export_pivot_cat_pct[cat].apply(lambda x: f"{x:.0f}%" if x >= 3 else ""),
                textposition='inside',
                textfont=dict(color='white', size=9, family='Arial Black'),
                hovertemplate=(
                    '<b>%{x}</b><br>' +
                    f'{cat}: %{{y:,.0f}} hours<br>' +
                    f'Share of all {cat} exports: %{{customdata:.1f}}%<br>' +
                    '<extra></extra>'
                ),
                customdata=export_pivot_cat_pct[cat].values
            )
        )
    
    # Add annotations for total hours and MS affected
    for i, country in enumerate(export_pivot_abs.index):
        total = total_hours[country]
        affected = countries_affected[i]
        
        # Total hours above bar
        fig5.add_annotation(
            x=i,
            y=total * 1.02,
            text=f"<b>{total:,.0f}h</b>",
            showarrow=False,
            font=dict(size=10, color='#333', family='Arial Black'),
            xanchor='center',
            yanchor='bottom'
        )
        
        # MS affected above total
        fig5.add_annotation(
            x=i,
            y=total * 1.08,
            text=f"→ {affected} MS",
            showarrow=False,
            font=dict(size=9, color='#666'),
            xanchor='center',
            yanchor='bottom'
        )
    
    # Calculate max for y-axis range
    max_total = total_hours.max()
    
    fig5.update_layout(
        title={
            'text': '<b>Top 15 Price-Exporting Countries: Export Magnitude by Technology (2025)</b><br><sub>Bar height = total hours | Percentages show share of each tech category\'s EU-wide exports</sub>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#333'}
        },
        barmode='stack',
        height=700,
        width=1500,
        plot_bgcolor='white',
        paper_bgcolor='white',
        xaxis=dict(
            title='<b>Exporting Country</b>',
            tickfont=dict(size=11)
        ),
        yaxis=dict(
            title='<b>Total Export Hours (Price-Setting Events)</b>',
            showgrid=True,
            gridcolor='#e0e0e0',
            tickfont=dict(size=11),
            range=[0, max_total * 1.15]
        ),
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02,
            font=dict(size=11)
        ),
        annotations=[
            dict(
                x=0.5,
                y=-0.12,
                xref='paper',
                yref='paper',
                text='<i>Example: "45%" in Gas bar for DE means Germany accounts for 45% of all Gas exports across EU</i>',
                showarrow=False,
                font=dict(size=10, color='#666'),
                xanchor='center'
            )
        ]
    )
    
    fig5.show()
    
    output_file5 = OUTPUT_DIR / 'top15_neighbour_countries_magnitude_by_tech_2025_EU.html'
    fig5.write_html(output_file5)
    print(f"Saved: {output_file5}")
    
    # ========================================
    # SUMMARY TABLE: SHARE OF EACH CATEGORY'S EXPORTS
    # ========================================
    
    print(f"\n{'='*140}")
    print(f"SHARE OF EACH TECHNOLOGY CATEGORY'S TOTAL EXPORTS (% of category)")
    print(f"{'='*140}")
    print(f"  Example: 'Gas: 45%' for DE means Germany accounts for 45% of all Gas exports in EU\n")
    
    _hdr = f"  {'Rank':<6s} {'Country':<8s} {'Total Hrs':>12s} | "
    _hdr += " ".join(f"{cat[:9]:>10s}" for cat in categories)
    _hdr += f" | {'MS':>4s}"
    print(_hdr)
    print(f"  {'-'*140}")
    
    for i, country in enumerate(export_pivot_abs.index, 1):
        total = total_hours[country]
        affected = countries_affected[i-1]
        _cells = " ".join(f"{export_pivot_cat_pct.loc[country, cat]:>9.1f}%" for cat in categories)
        print(f"  {i:<6d} {country:<8s} {total:>12,.0f} | {_cells} | {affected:>4d}")
    
    print(f"\n  {'-'*140}")
    _totals = " ".join(f"{category_totals[cat]:>9,.0f}h" for cat in categories)
    print(f"  {'Category Totals (hours):':<22s} | {_totals}")
    
    print(f"\n{'='*140}")
    
    # ========================================
    # ABSOLUTE VALUES TABLE
    # ========================================
    
    print(f"\n{'='*140}")
    print(f"ABSOLUTE EXPORT HOURS BY TECHNOLOGY CATEGORY")
    print(f"{'='*140}\n")
    
    _h3 = f"  {'Rank':<6s} {'Country':<8s} {'Total':>10s} "
    _h3 += " ".join(f"{cat[:9]:>10s}" for cat in categories)
    _h3 += f" {'MS':>6s}"
    print(_h3)
    print(f"  {'-'*140}")
    
    for i, country in enumerate(export_pivot_abs.index, 1):
        total = total_hours[country]
        affected = countries_affected[i-1]
        _vals = " ".join(f"{export_pivot_abs.loc[country, cat]:>10,.0f}" for cat in categories)
        print(f"  {i:<6d} {country:<8s} {total:>10,.0f} {_vals} {affected:>6d}")
    
    print(f"\n  {'-'*140}")
    
    # Totals
    total_all = total_hours.sum()
    _tvals = " ".join(f"{export_pivot_abs[cat].sum():>10,.0f}" for cat in categories)
    print(f"  {'TOTAL':<15s} {total_all:>10,.0f} {_tvals}")
    
    print(f"\n{'='*140}")
    
    # Key insights
    print(f"\nKey Insights:")
    print(f"  • Top exporter: {export_pivot_abs.index[0]} with {total_hours.iloc[0]:,.0f} hours")
    print(f"  • Most widespread: {export_pivot_abs.index[countries_affected.index(max(countries_affected))]} affecting {max(countries_affected)} Member States")
    
    # Dominance by category
    print(f"\n  Dominance by Technology Category:")
    for cat in categories:
        if category_totals[cat] > 0:
            top_country = export_pivot_cat_pct[cat].idxmax()
            top_pct = export_pivot_cat_pct.loc[top_country, cat]
            top_hours = export_pivot_abs.loc[top_country, cat]
            print(f"    • {cat:<25s}: {top_country} dominates with {top_pct:.1f}% ({top_hours:,.0f} hours)")
    
    # Concentration
    top5_share = total_hours.head(5).sum() / total_hours.sum() * 100
    print(f"\n  • Top 5 exporters account for {top5_share:.1f}% of all cross-border price-setting") 
# ========================================
# VIZ 6: HEATMAP - NEIGHBOUR COUNTRY → DOMESTIC COUNTRY FLOWS
# (This was VIZ 5 before, now becomes VIZ 6)
# ========================================
 
if len(df_neighbour_techs) > 0:
    print(f"\n{'='*80}")
    print("CREATING VISUALIZATION 6: COUNTRY-TO-COUNTRY FLOW HEATMAP")
    print(f"{'='*80}")
    
    # Deduplicate by timestamp
    df_neighbour_unique = df_neighbour_techs.drop_duplicates(subset=['country', 'neighbour_country', 'timestamp'])
    
    # Create pivot table
    flow_matrix = df_neighbour_unique.groupby(['neighbour_country', 'country']).size().reset_index(name='count')
    flow_pivot = flow_matrix.pivot(index='neighbour_country', columns='country', values='count').fillna(0)
 
    flow_pivot.to_csv(OUTPUT_DIR / 'neighbour_flow_matrix_HOURS_FULL_2025_EU.csv')
    
    
    # ALL countries (no top-20 truncation)
    top_neighbour_countries = flow_pivot.sum(axis=1).sort_values(ascending=False).index
    top_domestic_countries = flow_pivot.sum(axis=0).sort_values(ascending=False).index
    
    flow_pivot_subset = flow_pivot.loc[top_neighbour_countries, top_domestic_countries]
    
    # Convert to percentages
    flow_pivot_pct = flow_pivot_subset.div(flow_pivot_subset.sum(axis=1), axis=0) * 100
    
    fig6 = go.Figure(data=go.Heatmap(
        z=flow_pivot_pct.values,
        x=flow_pivot_pct.columns,
        y=flow_pivot_pct.index,
        colorscale='Oranges',
        hovertemplate=(
            'Neighbour Country: %{y}<br>' +
            'Domestic Country: %{x}<br>' +
            'Share: %{z:.0f}%<br>' +
            'Hours: %{customdata:,.0f}<br>' +
            '<extra></extra>'
        ),
        customdata=flow_pivot_subset.values,
        colorbar=dict(title="Share (%)")
    ))
    
    fig6.update_layout(
        title={
            'text': '<b>Neighbour Country → Domestic Country Price-Setting Flows (2025)</b><br><sub>% of each neighbour country\'s price-setting hours</sub>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'color': '#333'}
        },
        height=800,
        width=1200,
        xaxis=dict(
            title='<b>Domestic Country</b>',
            tickfont=dict(size=11)
        ),
        yaxis=dict(
            title='<b>Neighbour Country</b>',
            tickfont=dict(size=11)
        ),
        plot_bgcolor='white',
        paper_bgcolor='white'
    )
    
    fig6.show()
    
    output_file6 = OUTPUT_DIR / 'neighbour_country_flow_heatmap_2025_EU.html'
    fig6.write_html(output_file6)
    print(f"Saved: {output_file6}")
    
    # ========================================
    # TEXT SUMMARY
    # ========================================
    
    print(f"\n{'='*130}")
    print(f"KEY INSIGHTS: NEIGHBOUR PRICE SETTING BY TECHNOLOGY (2025) - COUNTRY PERSPECTIVE")
    print(f"{'='*130}\n")
    
    print(f"📊 TECHNOLOGY BREAKDOWN OF NEIGHBOUR IMPORTS (EU AVERAGE):")
    print(f"{'─'*130}\n")
    
    for cat in categories:
        col_name = f'neighbour_{cat.lower().replace(" ", "_")}_2025'
        avg_share = df_country_analysis[col_name].mean()
        print(f"  • {cat:<25s}: {avg_share:.1f}% of hours")
    
    print(f"\n{'─'*130}")
    print(f"📊 TOP 10 COUNTRIES BY TOTAL NEIGHBOUR INFLUENCE:")
    print(f"{'─'*130}\n")
    
    top10 = df_country_analysis.nlargest(10, 'neighbour_total_2025')
    for i, (idx, row) in enumerate(top10.iterrows(), 1):
        breakdown = []
        for cat in categories:
            col_name = f'neighbour_{cat.lower().replace(" ", "_")}_2025'
            val = row[col_name]
            if val > 1:
                breakdown.append(f"{cat}: {val:.1f}%")
        
        print(f"  {i:2d}. {row['country']:<5s}: Total {row['neighbour_total_2025']:>5.1f}%")
        if breakdown:
            print(f"      ({', '.join(breakdown)})")
    
    print(f"\n{'='*130}")
    print(f"COUNTRY-LEVEL ANALYSIS COMPLETE - VISUALIZATIONS WITH TECH BREAKDOWN")
    print(f"{'='*130}")
    
    # ========================================
    # EXPORT DataFrames
    # ========================================
    
    df_domestic_vs_neighbour_country_2025 = df_country_analysis.copy()
    df_neighbour_technology_detail_country_2025 = df_neighbour_techs.copy()
    
    csv_file1 = OUTPUT_DIR / 'domestic_vs_neighbour_by_tech_country_2025_EU.csv'
    df_domestic_vs_neighbour_country_2025.to_csv(csv_file1, index=False)
    
    csv_file2 = OUTPUT_DIR / 'neighbour_technology_detail_country_2025_EU.csv'
    df_neighbour_technology_detail_country_2025.to_csv(csv_file2, index=False)
    
    print(f"\n💾 CSV files saved:")
    print(f"   • {csv_file1}")
    print(f"   • {csv_file2}")
 
print("\n" + "=" * 80)
 


# In[25]:


# Cell 10.4b - EU RATE OF CROSS-BORDER PRICE-SETTING (binary per zone-hour) + BAR CHART
# For each EU zone and each hour: 1 if the marginal label involves ANY foreign
# bidding zone (co-marginal hours count once; intra-country cross-zone counts).
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

KNOWN_ZONES = sorted(all_zone_results.keys(), key=len, reverse=True)

def is_foreign(label, home):
    for t in str(label).split(' + '):
        t = t.strip()
        for z in KNOWN_ZONES:
            if t.endswith('_' + z) and z != home:
                return True
    return False

rows = []
for zone, res in all_zone_results.items():
    if zone not in ZONES_EU_ONLY or res is None or 'marginal_tech' not in res:
        continue
    lab = res['marginal_tech']
    fz = lab.map(lambda s: is_foreign(s, zone)) & lab.notna() & (lab != 'unknown')
    g = pd.DataFrame({'year': lab.index.year, 'fz': fz.values})
    a = g.groupby('year').agg(neighbour_set=('fz', 'sum'),
                              hours_present=('fz', 'size'))
    a['zone'] = zone
    rows.append(a.reset_index())

d = pd.concat(rows, ignore_index=True)
eu = d.groupby('year').agg(neighbour_set=('neighbour_set', 'sum'),
                           hours_present=('hours_present', 'sum'),
                           n_zones=('zone', 'nunique'))
hours_in_year = lambda y: 8784 if y % 4 == 0 else 8760
eu['denom_8760xN'] = [hours_in_year(y) * n for y, n in zip(eu.index, eu.n_zones)]
eu['pct_of_8760xN'] = eu.neighbour_set / eu.denom_8760xN * 100
eu['pct_of_present'] = eu.neighbour_set / eu.hours_present * 100
print(eu.round(1).to_string())
eu.to_csv(OUTPUT_DIR / 'crossborder_simple_rate_EU.csv')

# ---------------- bar chart: pct_of_present, 2015-2025 ----------------
p = eu[(eu.index >= 2015) & (eu.index <= 2025)]
fig, ax = plt.subplots(figsize=(9, 5))
bars = ax.bar(p.index.astype(int), p['pct_of_present'], color='#1f6fc5',
              edgecolor='white', width=0.7)
for b, v in zip(bars, p['pct_of_present']):
    ax.annotate(f'{v:.0f}%', (b.get_x() + b.get_width() / 2, v),
                textcoords='offset points', xytext=(0, 4),
                ha='center', fontsize=10, fontweight='bold')
ax.set_ylabel('Share of zone-hours (%)', fontsize=11)
ax.set_ylim(0, max(70, p['pct_of_present'].max() * 1.15))
ax.set_xticks(p.index.astype(int))
ax.set_title('Zone-hours with cross-border price-setting in the EU\n'
             '(any foreign bidding zone involved, binary per hour)',
             fontsize=13, fontweight='bold')
for s in ('top', 'right'):
    ax.spines[s].set_visible(False)
ax.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(OUTPUT_DIR / 'crossborder_rate_EU_bars_2015_2025.png',
            dpi=300, bbox_inches='tight')
plt.show()
print(f"Saved: crossborder_simple_rate_EU.csv + crossborder_rate_EU_bars_2015_2025.png")


# In[ ]:


# Cell 12f - EXPORT HOURLY MARGINAL TECH RESULTS FOR ALL ZONES (2015, 2019, 2025 & 2026)
# Updated to export 2015, 2019, 2025, and 2026 data with zone transition handling
# NOW EXTENDS TO 2026
# Modified 19/06/2026: combined export streamed to disk (no giant concat+sort) to avoid MemoryError

if CFG.OUTPUT_HEADLINE_ONLY:
    print("⏭️ Cell 12f skipped (OUTPUT_HEADLINE_ONLY=True)")
else:
    import pandas as pd
    import numpy as np
    from pathlib import Path
    
    print("=" * 80)
    print("EXPORTING HOURLY MARGINAL TECH RESULTS FOR ALL ZONES (2015, 2019, 2025 & 2026)")
    print("=" * 80)
    
    # ========================================
    # CONFIGURATION
    # ========================================
    
    # DEFAULT: Export 2015, 2019, 2025 & 2026
    # years_to_export = [2015, 2019, 2025, 2026]
    
    # OPTIONAL: Uncomment below to export ALL years 2015-2026
    years_to_export = list(range(2015, 2027))
    
    # OPTIONAL: Export only recent years
    # years_to_export = [2019, 2025, 2026]
    
    print(f"\n📅 Exporting years: {years_to_export}")
    print(f"📊 Zones available: {len(all_zone_results) if 'all_zone_results' in locals() else 0}")
    
    # ========================================
    # FUNCTION TO EXPORT SINGLE ZONE
    # ========================================
    
    def export_zone_results(zone, results_df, master_df, srmc_df, years=[2015, 2019, YEAR_H, 2026]):
        """
        Export hourly marginal tech results for a single zone
        Handles zone transitions (DE_AT_LU, DE_LU, AT)
        Updated for 2026
        
        Returns: export_df or None if failed
        """
        
        print(f"\n{'='*60}")
        print(f"EXPORTING: {zone}")
        print(f"{'='*60}")
        
        try:
            # ========== ZONE-SPECIFIC YEAR FILTERING ==========
            available_years = sorted(results_df.index.year.unique())
            
            # Zone transition handling
            if zone == 'DE_AT_LU':
                # Only has 2015-2018 data
                years_for_zone = [y for y in years if y <= 2018 and y in available_years]
                if not years_for_zone:
                    # Use first and last available year
                    years_for_zone = [available_years[0], available_years[-1]] if len(available_years) > 1 else available_years
            
            elif zone in ['DE_LU', 'AT']:
                # Only has 2019+ data (now including 2026)
                years_for_zone = [y for y in years if y >= 2019 and y in available_years]
                if not years_for_zone:
                    # Use first and last available year
                    years_for_zone = [available_years[0], available_years[-1]] if len(available_years) > 1 else available_years
            
            else:
                # Other zones: use requested years if available
                years_for_zone = [y for y in years if y in available_years]
                if not years_for_zone:
                    print(f"   No data found for requested years {years}")
                    print(f"     Available years: {available_years}")
                    return None
            
            print(f"  Years to export: {years_for_zone} (available: {available_years})")
            
            # ========== FILTER TO SPECIFIED YEARS ==========
            mask = results_df.index.year.isin(years_for_zone)
            
            if mask.sum() == 0:
                print(f"   No data found for years {years_for_zone}")
                return None
            
            # ========== BUILD EXPORT DATAFRAME ==========
            export_df = pd.DataFrame(index=results_df[mask].index)
            
            # Add datetime components
            export_df['datetime'] = export_df.index
            export_df['year'] = export_df.index.year
            export_df['month'] = export_df.index.month
            export_df['day'] = export_df.index.day
            export_df['hour'] = export_df.index.hour
            export_df['zone'] = zone
            
            # ========== MARKET PRICE ==========
            if 'price' in master_df.columns:
                export_df['market_price_eur_mwh'] = master_df.loc[export_df.index, 'price']
            
            # ========== MARGINAL TECH IDENTIFICATION ==========
            export_df['marginal_tech'] = results_df.loc[export_df.index, 'marginal_tech']
            
            if 'fallback_mode' in results_df.columns:
                export_df['fallback_mode'] = results_df.loc[export_df.index, 'fallback_mode']
            
            if 'price_setting_zone' in results_df.columns:
                export_df['price_setting_zone'] = results_df.loc[export_df.index, 'price_setting_zone']
            
            if 'flow_convergence' in results_df.columns:
                export_df['flow_convergence'] = results_df.loc[export_df.index, 'flow_convergence']
            
            if 'high_price_flag' in results_df.columns:
                export_df['high_price_flag'] = results_df.loc[export_df.index, 'high_price_flag']
            
            # ========== SRMC VALUES (ALL TECHNOLOGIES) ==========
            
            # Thermal techs (single value)
            thermal_techs = ['gas', 'coal', 'lignite', 'oil', 'oil_shale', 'nuclear', 'biomass', 'chp']
            for tech in thermal_techs:
                srmc_col = f'srmc_{tech}'
                if srmc_col in srmc_df.columns:
                    export_df[f'srmc_{tech}_eur_mwh'] = srmc_df.loc[export_df.index, srmc_col]
            
            # Hydro reservoir (min/max/avg)
            for suffix in ['min', 'max', 'avg']:
                col = f'srmc_hydro_reservoir_{suffix}'
                if col in srmc_df.columns:
                    export_df[f'srmc_hydro_reservoir_{suffix}_eur_mwh'] = srmc_df.loc[export_df.index, col]
            
            # Pumped storage (min/max/avg)
            for suffix in ['min', 'max', 'avg']:
                col = f'srmc_pumped_storage_{suffix}'
                if col in srmc_df.columns:
                    export_df[f'srmc_pumped_storage_{suffix}_eur_mwh'] = srmc_df.loc[export_df.index, col]
            
            # Renewables (static or min/max if available)
            renewable_techs = ['hydro_ror', 'wind_onshore', 'wind_offshore', 'solar', 'waste', 'other_renewable']
            for tech in renewable_techs:
                # Try single value first
                if f'srmc_{tech}' in srmc_df.columns:
                    export_df[f'srmc_{tech}_eur_mwh'] = srmc_df.loc[export_df.index, f'srmc_{tech}']
                # Try min/max/avg
                for suffix in ['min', 'max', 'avg']:
                    col = f'srmc_{tech}_{suffix}'
                    if col in srmc_df.columns:
                        export_df[f'srmc_{tech}_{suffix}_eur_mwh'] = srmc_df.loc[export_df.index, col]
            
            # ========== GENERATION LEVELS (ALL TECHNOLOGIES) ==========
            
            # Map generation columns
            gen_mapping = {
                'gas': None, 'coal': None, 'lignite': None, 'oil': None, 'oil_shale': None,
                'nuclear': None, 'biomass': None, 'chp': None,
                'hydro_reservoir': None, 'hydro_pumped': None, 'hydro_ror': None,
                'wind_onshore': None, 'wind_offshore': None, 'solar': None,
                'waste': None, 'other_renewable': None
            }
            
            for col in master_df.columns:
                col_lower = col.lower()
                if 'gen_gas' in col_lower and 'coal_derived' not in col_lower:
                    gen_mapping['gas'] = col
                elif 'gen_coal' in col_lower and 'lignite' not in col_lower and 'derived' not in col_lower:
                    gen_mapping['coal'] = col
                elif 'gen_lignite' in col_lower or 'brown_coal' in col_lower:
                    gen_mapping['lignite'] = col
                elif 'gen_oil_shale' in col_lower or 'oilshale' in col_lower:
                    gen_mapping['oil_shale'] = col
                elif 'gen_oil' in col_lower:
                    gen_mapping['oil'] = col
                elif 'gen_nuclear' in col_lower:
                    gen_mapping['nuclear'] = col
                elif 'gen_biomass' in col_lower:
                    gen_mapping['biomass'] = col
                elif 'gen_chp' in col_lower or 'gen_other' in col_lower:
                    gen_mapping['chp'] = col
                elif 'hydro_reservoir' in col_lower or 'hydro_water_reservoir' in col_lower:
                    gen_mapping['hydro_reservoir'] = col
                elif 'hydro_pumped' in col_lower or 'pumped_storage' in col_lower:
                    gen_mapping['hydro_pumped'] = col
                elif 'hydro_ror' in col_lower or 'run-of-river' in col_lower or 'run_of_river' in col_lower:
                    gen_mapping['hydro_ror'] = col
                elif 'wind_onshore' in col_lower:
                    gen_mapping['wind_onshore'] = col
                elif 'wind_offshore' in col_lower:
                    gen_mapping['wind_offshore'] = col
                elif 'gen_solar' in col_lower or col_lower == 'solar':
                    gen_mapping['solar'] = col
                elif 'gen_waste' in col_lower:
                    gen_mapping['waste'] = col
                elif 'other_renewable' in col_lower:
                    gen_mapping['other_renewable'] = col
            
            # Add generation values
            for tech, col in gen_mapping.items():
                if col and col in master_df.columns:
                    export_df[f'gen_{tech}_mw'] = master_df.loc[export_df.index, col]
            
            # ========== LOAD & FLOWS ==========
            if 'load' in master_df.columns:
                export_df['load_mw'] = master_df.loc[export_df.index, 'load']
            
            if 'imports' in master_df.columns:
                export_df['imports_mw'] = master_df.loc[export_df.index, 'imports']
            
            if 'exports' in master_df.columns:
                export_df['exports_mw'] = master_df.loc[export_df.index, 'exports']
            
            if 'net_imports' in master_df.columns:
                export_df['net_imports_mw'] = master_df.loc[export_df.index, 'net_imports']
            
            # ========== COMMODITY PRICES ==========
            if 'gas_price' in master_df.columns:
                export_df['gas_price_eur_mwh_th'] = master_df.loc[export_df.index, 'gas_price']
            
            if 'co2_price' in master_df.columns:
                export_df['co2_price_eur_tco2'] = master_df.loc[export_df.index, 'co2_price']
            
            if 'coal_price' in master_df.columns:
                export_df['coal_price_eur_tonne'] = master_df.loc[export_df.index, 'coal_price']
            
            if 'oil_price' in master_df.columns:
                export_df['oil_price_eur_tonne'] = master_df.loc[export_df.index, 'oil_price']
            
            if 'oil_shale_price' in master_df.columns:
                export_df['oil_shale_price_eur_mwh_th'] = master_df.loc[export_df.index, 'oil_shale_price']
            
            # ========== RESERVOIR DATA (IF AVAILABLE) ==========
            if 'reservoir_filling' in master_df.columns:
                export_df['reservoir_filling_pct'] = master_df.loc[export_df.index, 'reservoir_filling']
            
            # ========== SORT BY DATETIME ==========
            export_df = export_df.sort_values('datetime')
            
            # ========== EXPORT TO CSV ==========
            
            # Create filename with actual years exported
            year_range = f"{min(years_for_zone)}_{max(years_for_zone)}" if len(years_for_zone) > 1 else str(years_for_zone[0])
            output_file = OUTPUT_DIR / f'hourly_marginal_analysis_{zone}_{year_range}.csv'
            
            # Save combined file
            export_df.to_csv(output_file, index=False)
            
            print(f"  Saved: {output_file.name}")
            print(f"     Rows: {len(export_df):,}")
            print(f"     Columns: {len(export_df.columns)}")
            
            # Quick stats
            for year in years_for_zone:
                year_data = export_df[export_df['year'] == year]
                if len(year_data) > 0:
                    identified = (year_data['marginal_tech'] != 'unknown').sum()
                    pct = identified / len(year_data) * 100
                    avg_price = year_data['market_price_eur_mwh'].mean()
                    print(f"     {year}: {len(year_data):>6,} hours, {pct:>5.1f}% identified, avg price {avg_price:>6.2f} EUR/MWh")
            
            return export_df
        
        except Exception as e:
            print(f"   Error exporting {zone}: {str(e)[:100]}")
            
            # Save error log
            error_log = OUTPUT_DIR / f'EXPORT_ERROR_{zone}.txt'
            with open(error_log, 'w') as f:
                import traceback
                f.write(f"Zone: {zone}\n")
                f.write(f"Timestamp: {pd.Timestamp.now()}\n\n")
                f.write(traceback.format_exc())
            
            return None
    
    
    # ========================================
    # EXPORT ALL ZONES
    # ========================================
    
    if 'all_zone_results' not in locals() or not all_zone_results:
        print("\n No zone results available!")
        print("   Run Cell 12e or 12e_EU first")
    else:
        export_summary = {
            'success': [],
            'failed': [],
            'exported_dfs': {}
        }
        
        print(f"\n{'='*80}")
        print("STARTING MULTI-ZONE EXPORT")
        print(f"{'='*80}")
        
        for i, zone in enumerate(sorted(all_zone_results.keys()), 1):
            print(f"\n[{i}/{len(all_zone_results)}] Exporting {zone}...")
            
            results_df = all_zone_results[zone]
            master_df = all_zone_masters.get(zone)
            srmc_df = all_zone_srmcs.get(zone)
            
            if master_df is None or srmc_df is None:
                print(f"   Missing master or SRMC data for {zone} - SKIPPING")
                export_summary['failed'].append(zone)
                continue
            
            exported_df = export_zone_results(zone, results_df, master_df, srmc_df, years_to_export)
            
            if exported_df is not None:
                export_summary['success'].append(zone)
                export_summary['exported_dfs'][zone] = exported_df
            else:
                export_summary['failed'].append(zone)
        
        # ========================================
        # CREATE COMBINED EXPORT (ALL ZONES) — memory-safe, streamed to disk
        # Modified 19/06/2026: stream per-zone CSVs to disk instead of one giant
        # concat + sort_values (the ~4M-row x 68-col copy overflowed RAM).
        # Zones written in sorted order, each datetime-sorted -> same layout as
        # the old sort_values(['zone','datetime']), but peak memory = one zone.
        # ========================================
        
        if export_summary['exported_dfs']:
            print(f"\n{'='*80}")
            print("CREATING COMBINED EXPORT (ALL ZONES)")
            print(f"{'='*80}")
            
            # Year range across all exported zones (no concat needed)
            actual_years = sorted({int(y)
                                for df in export_summary['exported_dfs'].values()
                                for y in df['year'].unique()})
            year_range = f"{min(actual_years)}_{max(actual_years)}"
            combined_file = OUTPUT_DIR / f'hourly_marginal_analysis_ALL_ZONES_{year_range}.csv'
            
            # Stream each zone (sorted order) -> grouped by zone, datetime-sorted within zone
            total_rows = 0
            first = True
            for zone in sorted(export_summary['exported_dfs'].keys()):
                zdf = export_summary['exported_dfs'][zone].sort_values('datetime')  # per-zone sort only
                zdf.to_csv(combined_file, mode='w' if first else 'a', header=first, index=False)
                total_rows += len(zdf)
                first = False
            
            print(f"\nCombined export saved: {combined_file.name}")
            print(f"   Total rows: {total_rows:,}")
            print(f"   Zones: {len(export_summary['exported_dfs'])}")
            print(f"   Years: {actual_years}")
            
            # ========== SUMMARY BY ZONE AND YEAR (built per-zone, no monolith) ==========
            print(f"\n{'='*80}")
            print("EXPORT SUMMARY BY ZONE AND YEAR")
            print(f"{'='*80}")
            
            summary_rows = []
            for zone in sorted(export_summary['exported_dfs'].keys()):
                zdf = export_summary['exported_dfs'][zone]
                for year, yg in zdf.groupby('year'):
                    total_h = len(yg)
                    identified = int((yg['marginal_tech'] != 'unknown').sum())
                    avg_price = yg['market_price_eur_mwh'].mean() if 'market_price_eur_mwh' in yg.columns else float('nan')
                    summary_rows.append({
                        'zone': zone,
                        'year': int(year),
                        'total_hours': total_h,
                        'identified_hours': identified,
                        'avg_price_eur_mwh': round(avg_price, 2),
                        'identification_rate_%': round(identified / total_h * 100, 1) if total_h else 0.0,
                    })
            summary_table = pd.DataFrame(summary_rows).set_index(['zone', 'year'])
            print("\n" + summary_table.to_string())
            
            summary_file = OUTPUT_DIR / f'export_summary_by_zone_year_{year_range}.csv'
            summary_table.to_csv(summary_file)
            print(f"\n📊 Summary table saved: {summary_file.name}")
            
            # ========== TOP 3 MARGINAL TECHS BY ZONE (per-zone, raw labels) ==========
            print(f"\n{'='*80}")
            print("TOP 3 MARGINAL TECHNOLOGIES BY ZONE")
            print(f"{'='*80}")
            
            for zone in sorted(export_summary['exported_dfs'].keys()):
                zdf = export_summary['exported_dfs'][zone]
                years_in_zone = sorted(int(y) for y in zdf['year'].unique())
                print(f"\n{zone} ({min(years_in_zone)}-{max(years_in_zone)}):")
                tech_counts = zdf['marginal_tech'].value_counts().head(3)
                for i, (tech, count) in enumerate(tech_counts.items(), 1):
                    pct = count / len(zdf) * 100
                    print(f"  {i}. {tech:30s}: {count:>6,} hours ({pct:>5.1f}%)")
        
        # ========================================
        # FINAL SUMMARY
        # ========================================
        
        print(f"\n{'='*80}")
        print("EXPORT COMPLETE")
        print(f"{'='*80}")
        
        print(f"\nSuccessfully exported: {len(export_summary['success'])} zones")
        if export_summary['success']:
            print(f"   {', '.join(sorted(export_summary['success']))}")
        
        if export_summary['failed']:
            print(f"\n Failed to export: {len(export_summary['failed'])} zones")
            print(f"   {', '.join(sorted(export_summary['failed']))}")
        
        print(f"\n Output directory: {OUTPUT_DIR}")
        print(f"\n{'='*80}")
    
    print("\n" + "=" * 80)
    
    


# In[ ]:




