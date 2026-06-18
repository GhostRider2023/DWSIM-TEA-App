"""
equipment_models.py
Turton, Bailie, Whiting & Shaeiwitz (2012) cost correlations for common
chemical process equipment, updated to 2024 costs using CEPCI index.

Reference: "Analysis, Synthesis and Design of Chemical Processes", 4th Ed.
Table A.1 / Appendix A ─ purchased equipment cost correlations (2001 basis).
"""
from __future__ import annotations

import os
import json
import math
from dataclasses import dataclass, field
from typing import Optional, Union


# ---------------------------------------------------------------------------
# CEPCI Cost Index
# ---------------------------------------------------------------------------
CEPCI_2001 = 394.3   # Turton base year from dwsim_equipment_cost_dataset.json
CEPCI_2024 = 780.0   # Approximate 2024 value (user can override)


def cepci_factor(cepci_year: float = CEPCI_2024) -> float:
    return cepci_year / CEPCI_2001


# ---------------------------------------------------------------------------
# Lang / Hand Installation Factors
# ---------------------------------------------------------------------------

LANG_FACTORS = {
    "Distillation Column": 4.16,
    "Heat Exchanger":      3.17,
    "Pump":                3.30,
    "Compressor":          2.15,
    "Reactor":             4.16,
    "Vessel":              4.16,
    "Mixer":               2.50,
    "Splitter":            2.50,
    "Valve":               2.00,
    "Other":               3.00,
}

BARE_MODULE_FACTORS = {
    # F_BM values from Turton Table A.4 (carbon steel, typical pressure)
    "Distillation Column": 4.16,
    "Heat Exchanger":      3.17,
    "Pump":                3.30,
    "Compressor":          2.15,
    "Reactor":             4.16,
    "Vessel":              4.16,
    "Mixer":               2.50,
    "Splitter":            2.50,
    "Valve":               2.00,
    "Other":               3.00,
}

# OSBL, contingency, engineering fractions applied to ISBL
TOTAL_CAPEX_FACTOR = 1.45   # ISBL -> Total Fixed Capital


# ---------------------------------------------------------------------------
# Utility OPEX unit costs (USD / GJ)
# ---------------------------------------------------------------------------
UTILITY_COSTS = {
    "LP Steam (3 bar)":      6.62,    # $/GJ
    "MP Steam (10 bar)":     8.22,
    "HP Steam (42 bar)":    12.32,
    "Cooling Water":         0.354,
    "Chilled Water":         4.77,
    "Refrigerant (-25°C)":  23.53,
    "Electricity":          16.80,    # $/GJ ≈ 0.06 $/kWh
    "Fuel (natural gas)":    4.47,
}

# Map equipment duty type to likely utility
HEAT_UTILITY_MAP = {
    "reboiler": "MP Steam (10 bar)",
    "condenser": "Cooling Water",
    "heater": "MP Steam (10 bar)",
    "cooler": "Cooling Water",
}


# ---------------------------------------------------------------------------
# Dynamic Costing Database Loading
# ---------------------------------------------------------------------------
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
DATASET_JSON_PATH = os.path.join(DATA_DIR, "dwsim_equipment_cost_dataset.json")


def load_cost_dataset() -> dict:
    if os.path.exists(DATASET_JSON_PATH):
        try:
            with open(DATASET_JSON_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


# Global costing dataset
COST_DATASET = load_cost_dataset()


def normalize_key(name: str) -> str:
    if not name:
        return ""
    s = name.lower().strip()
    s = s.replace(" ", "_").replace("-", "_")
    s = s.replace("(", "").replace(")", "")
    return s


# ---------------------------------------------------------------------------
# Fallback correlations (originally from Turton table)
# ---------------------------------------------------------------------------
@dataclass
class TurtonCorr:
    k1: float
    k2: float
    k3: float
    a_min: float
    a_max: float
    unit: str
    n: float = 0.6

    def purchase_cost_2001(self, A: float) -> float:
        A_c = max(self.a_min, min(self.a_max, A))
        log_a = math.log10(A_c)
        log_c = self.k1 + self.k2 * log_a + self.k3 * log_a ** 2
        cost = 10 ** log_c
        if A < self.a_min:
            cost = cost * ((A / self.a_min) ** self.n)
        elif A > self.a_max:
            cost = cost * ((A / self.a_max) ** self.n)
        return cost


CORRELATIONS: dict[str, TurtonCorr] = {
    "column_vessel":  TurtonCorr(3.4974, 0.4485, 0.1074, 0.3, 520.0, "m3 (volume)", 0.6),
    "sieve_tray":     TurtonCorr(2.9949, 0.4465, 0.3961, 0.7, 12.3, "m2 (area)", 0.86),
    "valve_tray":     TurtonCorr(3.3322, 0.4838, 0.3434, 0.7, 10.5, "m2 (area)", 1.0),
    "hx_shell_tube":  TurtonCorr(4.3247, -0.3030, 0.1634, 10.0, 1000.0, "m2", 0.62),
    "pump_cent":      TurtonCorr(3.3892,  0.0536, 0.1538, 1.0, 300.0, "kW", 0.67),
    "compressor":     TurtonCorr(2.2897,  1.3604, -0.1027, 450.0, 3000.0, "kW", 0.67),
    "vessel_vert":    TurtonCorr(3.4974,  0.4485, 0.1074, 0.3, 520.0, "m3", 0.6),
    "vessel_horiz":   TurtonCorr(3.5565,  0.3776, 0.0905, 0.1, 628.0, "m3", 0.5),
    "kettle_reboiler":TurtonCorr(4.4646, -0.5277, 0.3955, 10.0, 1000.0, "m2", 0.59),
}


# ---------------------------------------------------------------------------
# Sizing and Factor Helper Functions
# ---------------------------------------------------------------------------

def calculate_purchased_cost_c0(category: str, type_name: str, capacity: float) -> tuple[float, list[str]]:
    warnings = []
    cat = normalize_key(category)
    typ = normalize_key(type_name)

    # Use dataset JSON if available
    if COST_DATASET:
        categories = COST_DATASET.get("equipment_categories", {})
        cat_data = categories.get(cat, {})
        types = cat_data.get("types", {})
        type_data = types.get(typ, {})
        if type_data:
            cost_info = type_data.get("cost", {})
            if cost_info:
                k1 = cost_info.get("K1", 0.0)
                k2 = cost_info.get("K2", 0.0)
                k3 = cost_info.get("K3", 0.0)
                a_min = cost_info.get("Amin", 0.1)
                a_max = cost_info.get("Amax", 1e6)
                n = cost_info.get("n", 0.6)
                unit = cost_info.get("unit", "")
                
                if capacity <= 0:
                    capacity = a_min
                    warnings.append(f"w Capacity {capacity} {unit} is zero/invalid; using Amin ({a_min})")
                
                if capacity < a_min:
                    warnings.append(f"w Capacity {capacity:.2f} {unit} is below Amin ({a_min}) {unit}; extrapolated exponentially")
                    log_a = math.log10(a_min)
                    log_c = k1 + k2 * log_a + k3 * (log_a ** 2)
                    c_bound = 10 ** log_c
                    cost = c_bound * ((capacity / a_min) ** n)
                elif capacity > a_max:
                    warnings.append(f"w Capacity {capacity:.2f} {unit} is above Amax ({a_max}) {unit}; extrapolated exponentially")
                    log_a = math.log10(a_max)
                    log_c = k1 + k2 * log_a + k3 * (log_a ** 2)
                    c_bound = 10 ** log_c
                    cost = c_bound * ((capacity / a_max) ** n)
                else:
                    log_a = math.log10(capacity)
                    log_c = k1 + k2 * log_a + k3 * (log_a ** 2)
                    cost = 10 ** log_c
                return cost, warnings

    # Fallback to local CORRELATIONS map
    fallback_key = None
    if cat == "pump":
        fallback_key = "pump_cent"
    elif cat == "compressor":
        fallback_key = "compressor"
    elif cat == "heat_exchanger":
        if typ == "kettle_reboiler":
            fallback_key = "kettle_reboiler"
        else:
            fallback_key = "hx_shell_tube"
    elif cat == "vessel":
        if typ == "horizontal":
            fallback_key = "vessel_horiz"
        else:
            fallback_key = "vessel_vert"
    elif cat == "tray":
        if typ == "valve":
            fallback_key = "valve_tray"
        else:
            fallback_key = "sieve_tray"

    if fallback_key and fallback_key in CORRELATIONS:
        corr = CORRELATIONS[fallback_key]
        if capacity <= 0:
            capacity = corr.a_min
            warnings.append(f"w Capacity {capacity} {corr.unit} is zero/invalid; using Amin ({corr.a_min})")
            
        cost = corr.purchase_cost_2001(capacity)
        if capacity < corr.a_min:
            warnings.append(f"w Capacity {capacity:.2f} {corr.unit} is below Amin ({corr.a_min}) {corr.unit}; extrapolated exponentially")
        elif capacity > corr.a_max:
            warnings.append(f"w Capacity {capacity:.2f} {corr.unit} is above Amax ({corr.a_max}) {corr.unit}; extrapolated exponentially")
        return cost, warnings

    return 10000.0, [f"Warning: Correlation for {category}/{type_name} not found. Using default cost."]


def calculate_pressure_factor_vessel(design_pressure_barg: float, diameter_m: float) -> float:
    if design_pressure_barg < -0.5:
        return 1.25
    pd = design_pressure_barg
    d = diameter_m
    num = (pd + 1.0) * d
    den = 2.0 * (850.0 - 0.6 * (pd + 1.0))
    if den <= 0:
        return 1.0
    ts = (num / den)
    fp = (ts + 0.00315) / 0.0063
    return max(1.0, fp)


def calculate_pressure_factor_ancillary(category: str, type_name: str, pressure_barg: float) -> tuple[float, list[str]]:
    warnings = []
    cat = normalize_key(category)
    typ = normalize_key(type_name)
    
    if pressure_barg <= 0:
        return 1.0, []

    if COST_DATASET:
        categories = COST_DATASET.get("equipment_categories", {})
        cat_data = categories.get(cat, {})
        types = cat_data.get("types", {})
        type_data = types.get(typ, {})
        if type_data:
            pf_info = type_data.get("pressure_factor", {})
            if pf_info:
                c1 = pf_info.get("C1", 0.0)
                c2 = pf_info.get("C2", 0.0)
                c3 = pf_info.get("C3", 0.0)
                p_min = pf_info.get("Pmin_barg", 0.0)
                p_max = pf_info.get("Pmax_barg", 1000.0)
                
                if pressure_barg < p_min:
                    return 1.0, warnings
                
                if pressure_barg > p_max:
                    warnings.append(f"w Operating pressure {pressure_barg:.1f} barg exceeds correlation limit ({p_max} barg) for {category}/{type_name}; pressure factor extrapolated at bound")
                    P_val = p_max
                else:
                    P_val = pressure_barg
                    
                log_p = math.log10(P_val)
                log_fp = c1 + c2 * log_p + c3 * (log_p ** 2)
                fp = max(1.0, 10 ** log_fp)
                return fp, warnings

    # Fallback pressure factors
    if cat == "pump":
        p_min, p_max = 10.0, 100.0
        c1, c2, c3 = -0.3935, 0.3957, -0.00226
    elif cat == "heat_exchanger":
        p_min, p_max = 5.0, 140.0
        c1, c2, c3 = 0.03881, -0.11272, 0.08183
    else:
        return 1.0, []
        
    if pressure_barg < p_min:
        return 1.0, warnings
    
    if pressure_barg > p_max:
        warnings.append(f"w Operating pressure {pressure_barg:.1f} barg exceeds correlation limit ({p_max} barg) for {category}/{type_name}; pressure factor extrapolated at bound")
        P_val = p_max
    else:
        P_val = pressure_barg
        
    log_p = math.log10(P_val)
    log_fp = c1 + c2 * log_p + c3 * (log_p ** 2)
    fp = max(1.0, 10 ** log_fp)
    return fp, warnings


def calculate_bare_module_factor(
    category: str,
    type_name: str,
    material: str,
    pressure_factor: float,
    diameter_m: float = 1.0,
    design_pressure_barg: float = 0.0,
) -> tuple[float, list[str]]:
    warnings = []
    cat = normalize_key(category)
    typ = normalize_key(type_name)
    
    if COST_DATASET:
        categories = COST_DATASET.get("equipment_categories", {})
        cat_data = categories.get(cat, {})
        types = cat_data.get("types", {})
        type_data = types.get(typ, {})
        
        bm_method = cat_data.get("bare_module_method", "standard")
        if type_data and "bare_module_method" in type_data:
            bm_method = type_data["bare_module_method"]
            
        if bm_method == "constant":
            fbm = cat_data.get("bare_module_factor_constant", 1.38)
            if type_data and "bare_module_factor_constant" in type_data:
                fbm = type_data["bare_module_factor_constant"]
            return fbm, []
            
        elif bm_method == "material_lookup_pressure_independent":
            by_mat = type_data.get("bare_module_factor_by_material", {})
            if not by_mat:
                return 2.8, [f"Warning: bare module factor by material not found for {category}/{type_name}. Defaulting to 2.8."]
            mat_key = None
            for k in by_mat.keys():
                if k.lower() == material.lower() or normalize_key(k) == normalize_key(material):
                    mat_key = k
                    break
            if not mat_key:
                first_key = list(by_mat.keys())[0]
                warnings.append(f"w Material {material} not found for {category}/{type_name}; defaulted to {first_key}")
                mat_key = first_key
            return by_mat[mat_key], warnings
            
        elif bm_method == "FBM_equals_material_factor":
            mat_factors = type_data.get("material_factor", {})
            if not mat_factors:
                return 1.0, [f"Warning: Material factor not found for tray {type_name}. Defaulting to 1.0."]
            mat_key = None
            for k in mat_factors.keys():
                if k.lower() == material.lower() or normalize_key(k) == normalize_key(material):
                    mat_key = k
                    break
            if not mat_key:
                first_key = list(mat_factors.keys())[0]
                warnings.append(f"w Material {material} not found for {category}/{type_name}; defaulted to {first_key}")
                mat_key = first_key
            return mat_factors[mat_key], warnings
            
        elif bm_method in ("standard", "standard_with_vessel_pressure_factor"):
            bm_info = type_data.get("bare_module", {})
            if not bm_info:
                bm_info = cat_data.get("bare_module", {})
            b1 = bm_info.get("B1", 1.0)
            b2 = bm_info.get("B2", 1.0)
            
            fm = 1.0
            if cat == "heat_exchanger":
                comb = cat_data.get("material_factor_combinations", {})
                vals = comb.get("values", {})
                mat_key = None
                for k in vals.keys():
                    if k.lower() == material.lower() or normalize_key(k) == normalize_key(material):
                        mat_key = k
                        break
                if not mat_key:
                    parts = material.split("/")
                    if len(parts) == 2:
                        rev_mat = f"{parts[1]}/{parts[0]}"
                        for k in vals.keys():
                            if k.lower() == rev_mat.lower() or normalize_key(k) == normalize_key(rev_mat):
                                mat_key = k
                                break
                if not mat_key:
                    first_key = list(vals.keys())[0]
                    warnings.append(f"w Material combination {material} not found for heat exchanger; defaulted to {first_key}")
                    mat_key = first_key
                fm = vals[mat_key]
            else:
                mat_factors = type_data.get("material_factor", {})
                if not mat_factors:
                    mat_factors = cat_data.get("material_factor", {})
                if mat_factors:
                    mat_key = None
                    for k in mat_factors.keys():
                        if k.lower() == material.lower() or normalize_key(k) == normalize_key(material):
                            mat_key = k
                            break
                    if not mat_key:
                        first_key = list(mat_factors.keys())[0]
                        warnings.append(f"w Material {material} not found for {category}/{type_name}; defaulted to {first_key}")
                        mat_key = first_key
                    fm = mat_factors[mat_key]
            fbm = max(1.0, b1 + b2 * fm * pressure_factor)
            return fbm, warnings

    # Fallback bare module factors
    fbm = BARE_MODULE_FACTORS.get(category, 3.0)
    return fbm, []


def calculate_tray_quantity_factor(n_trays: float) -> float:
    if n_trays >= 20:
        return 1.0
    elif n_trays <= 0:
        return 2.25
    else:
        return 2.25 / (1.0414 ** n_trays)


def _resolve_material_and_factor(
    category: str,
    type_name: str,
    material_factor: float | str,
    material: Optional[str] = None
) -> tuple[str, float]:
    if isinstance(material_factor, str):
        material = material_factor
        material_factor = 1.0
        
    if not material:
        cat = normalize_key(category)
        if cat == "heat_exchanger":
            material = "CS/CS"
        else:
            material = "CS_carbon_steel"
            
    fm = 1.0
    found = False
    
    cat = normalize_key(category)
    typ = normalize_key(type_name)
    
    if COST_DATASET:
        categories = COST_DATASET.get("equipment_categories", {})
        cat_data = categories.get(cat, {})
        types = cat_data.get("types", {})
        type_data = types.get(typ, {})
        
        if cat == "heat_exchanger":
            comb = cat_data.get("material_factor_combinations", {})
            vals = comb.get("values", {})
            for k, v in vals.items():
                if k.lower() == material.lower() or normalize_key(k) == normalize_key(material):
                    fm = v
                    found = True
                    break
            if not found:
                parts = material.split("/")
                if len(parts) == 2:
                    rev_mat = f"{parts[1]}/{parts[0]}"
                    for k, v in vals.items():
                        if k.lower() == rev_mat.lower() or normalize_key(k) == normalize_key(rev_mat):
                            fm = v
                            found = True
                            break
        else:
            mat_factors = type_data.get("material_factor", {}) if type_data else {}
            if not mat_factors:
                mat_factors = cat_data.get("material_factor", {})
            if mat_factors:
                for k, v in mat_factors.items():
                    if k.lower() == material.lower() or normalize_key(k) == normalize_key(material):
                        fm = v
                        found = True
                        break
                        
    if not found:
        # Fallback factors mapping
        mat_lower = material.lower()
        if cat == "heat_exchanger":
            comb_factors = {"cs/cs": 1.0, "cs/ss": 1.8, "ss/ss": 2.9, "cs/ni": 2.8, "ni/ni": 3.8, "cs/ti": 4.6, "ti/ti": 11.4}
            fm = comb_factors.get(normalize_key(material), 1.0)
        elif cat == "vessel":
            vess_factors = {"cs_carbon_steel": 1.0, "ss_stainless_steel": 3.1, "ni_nickel_alloy": 7.1, "ti_titanium": 9.4}
            fm = vess_factors.get(normalize_key(material), 1.0)
        elif cat == "pump":
            pump_factors = {"fe_cast_iron": 1.0, "cs_carbon_steel": 1.6, "ss_stainless_steel": 2.3, "ni_nickel_alloy": 4.4}
            fm = pump_factors.get(normalize_key(material), 1.6)
        elif cat == "tray":
            tray_factors = {"cs_carbon_steel": 1.0, "ss_stainless_steel": 1.8, "ni_nickel_alloy": 5.6}
            fm = tray_factors.get(normalize_key(material), 1.0)
            
    if isinstance(material_factor, (int, float)) and material_factor != 1.0:
        return material, float(material_factor)
        
    return material, fm


# ---------------------------------------------------------------------------
# Equipment cost results
# ---------------------------------------------------------------------------

@dataclass
class EquipmentCost:
    tag: str
    eq_type: str
    size_param: str      # description of size parameter used
    size_value: float
    purchase_cost_2001: float
    purchase_cost_current: float
    bare_module_cost: float
    installed_cost: float
    notes: str = ""
    warnings: list[str] = field(default_factory=list)
    details: dict = field(default_factory=dict)


def _column_volume(diameter_m: float, height_m: float) -> float:
    """Shell volume of a cylindrical column (m3)."""
    return math.pi / 4 * diameter_m ** 2 * height_m


def cost_distillation_column(
    tag: str,
    diameter_m: float,
    height_m: float,
    n_stages: float,
    tray_spacing_m: float = 0.6,
    cepci: float = CEPCI_2024,
    material_factor: float | str = 1.0,
    shell_material: str = "CS_carbon_steel",
    tray_material: str = "CS_carbon_steel",
    tray_type: str = "sieve",
    pressure_barg: float = 0.0,
) -> EquipmentCost:
    """
    Turton shell + tray cost model using loaded data.
    Shell cost: vertical vessel correlation with total volume.
    Tray cost: sieve/valve tray correlation x n_stages x Fq.
    """
    warnings = []
    if not diameter_m or diameter_m <= 0:
        diameter_m = 1.0
        warnings.append("Diameter missing; assumed 1.0 m")
    if not n_stages or n_stages <= 0:
        n_stages = 10
        warnings.append("Stages missing; assumed 10")
    if not tray_spacing_m or tray_spacing_m <= 0:
        tray_spacing_m = 0.6

    # Sizing physical column height dynamically using spacing and allowances
    top_allowance = 1.5
    bottom_allowance = 1.5
    height_actual = (n_stages * tray_spacing_m) + top_allowance + bottom_allowance

    aspect_ratio = height_actual / diameter_m if diameter_m > 0 else 0.0
    if aspect_ratio < 10.0 or aspect_ratio > 40.0:
        warnings.append(
            f"w {tag}: Height/Diameter ratio ({aspect_ratio:.1f}) is outside standard range [10.0 - 40.0]. Check column sizing."
        )

    # Resolve materials
    shell_mat, _ = _resolve_material_and_factor("vessel", "vertical", material_factor if isinstance(material_factor, float) else shell_material, shell_material)
    tray_mat, _ = _resolve_material_and_factor("tray", tray_type, tray_material, tray_material)
    
    # 1. Shell Volume & Base Cost (using height_actual)
    vol = _column_volume(diameter_m, height_actual)
    shell_p_2001, shell_p_warn = calculate_purchased_cost_c0("vessel", "vertical", vol)
    warnings.extend(shell_p_warn)
    
    # Shell Pressure Factor
    shell_fp = calculate_pressure_factor_vessel(pressure_barg, diameter_m)
    
    # Shell Bare Module Factor
    shell_fbm, shell_fbm_warn = calculate_bare_module_factor("vessel", "vertical", shell_mat, shell_fp, diameter_m, pressure_barg)
    warnings.extend(shell_fbm_warn)
    
    shell_cbm_2001 = shell_p_2001 * shell_fbm
 
    # 2. Tray Sizing & Base Cost
    tray_area = math.pi / 4 * diameter_m ** 2
    tray_p_one_2001, tray_p_warn = calculate_purchased_cost_c0("tray", tray_type, tray_area)
    warnings.extend(tray_p_warn)
    
    # Quantity factor discount Fq
    fq = calculate_tray_quantity_factor(n_stages)
    tray_p_total_2001 = tray_p_one_2001 * int(n_stages) * fq
    
    # Tray Bare Module Factor
    tray_fbm, tray_fbm_warn = calculate_bare_module_factor("tray", tray_type, tray_mat, 1.0)
    warnings.extend(tray_fbm_warn)
    
    tray_cbm_2001 = tray_p_total_2001 * tray_fbm

    # 3. Sum & Escalate
    total_p_2001 = shell_p_2001 + tray_p_total_2001
    total_cbm_2001 = shell_cbm_2001 + tray_cbm_2001

    factor = cepci_factor(cepci)
    total_p_current = total_p_2001 * factor
    total_cbm_current = total_cbm_2001 * factor

    return EquipmentCost(
        tag=tag,
        eq_type="Distillation Column",
        size_param=f"D={diameter_m:.2f}m, H={height_actual:.1f}m, N={int(n_stages)} trays",
        size_value=vol,
        purchase_cost_2001=total_p_2001,
        purchase_cost_current=total_p_current,
        bare_module_cost=total_cbm_current,
        installed_cost=total_cbm_current * TOTAL_CAPEX_FACTOR,
        notes=f"Shell ({vol:.2f} m3, {shell_mat}, Fp={shell_fp:.2f}) + {int(n_stages)} {tray_type} trays ({tray_area:.2f} m2 each, {tray_mat}, Fq={fq:.2f})",
        warnings=warnings,
        details={}
    )


def cost_heat_exchanger(
    tag: str,
    area_m2: float,
    cepci: float = CEPCI_2024,
    material_factor: float | str = 1.0,
    material: Optional[str] = None,
    hx_type: str = "fixed_tube_sheet",
    pressure_barg: float = 0.0,
) -> EquipmentCost:
    warnings = []
    hx_mat, _ = _resolve_material_and_factor("heat_exchanger", hx_type, material_factor, material)
    
    # Shell splitting logic (Turton correlations valid up to 1000 m2)
    n_shells = max(1, math.ceil(area_m2 / 1000.0))
    area_per_shell = area_m2 / n_shells
    
    # Calculate purchased cost of single shell
    p_shell_2001, p_warn = calculate_purchased_cost_c0("heat_exchanger", hx_type, area_per_shell)
    warnings.extend(p_warn)
    
    p_2001 = p_shell_2001 * n_shells
    
    # Calculate pressure factor
    fp, fp_warn = calculate_pressure_factor_ancillary("heat_exchanger", hx_type, pressure_barg)
    warnings.extend(fp_warn)
    
    # Calculate bare module factor
    fbm, fbm_warn = calculate_bare_module_factor("heat_exchanger", hx_type, hx_mat, fp)
    warnings.extend(fbm_warn)
    
    factor = cepci_factor(cepci)
    p_cur = p_2001 * factor
    cbm_cur = p_cur * fbm
    
    notes = f"Type: {hx_type}, Material: {hx_mat}, Fp={fp:.2f}"
    if n_shells > 1:
        notes += f", Shells: {n_shells} (Area/shell: {area_per_shell:.1f} m2)"
    
    return EquipmentCost(
        tag=tag,
        eq_type="Heat Exchanger",
        size_param=f"A={area_m2:.2f} m2",
        size_value=area_m2,
        purchase_cost_2001=p_2001,
        purchase_cost_current=p_cur,
        bare_module_cost=cbm_cur,
        installed_cost=cbm_cur * TOTAL_CAPEX_FACTOR,
        notes=notes,
        warnings=warnings,
        details={}
    )


def cost_pump(
    tag: str,
    power_kw: float,
    cepci: float = CEPCI_2024,
    material_factor: float | str = 1.0,
    material: Optional[str] = None,
    pump_type: str = "centrifugal",
    pressure_barg: float = 0.0,
) -> EquipmentCost:
    warnings = []
    pump_mat, _ = _resolve_material_and_factor("pump", pump_type, material_factor, material)
    
    p_2001, p_warn = calculate_purchased_cost_c0("pump", pump_type, power_kw)
    warnings.extend(p_warn)
    
    fp, fp_warn = calculate_pressure_factor_ancillary("pump", pump_type, pressure_barg)
    warnings.extend(fp_warn)
    
    fbm, fbm_warn = calculate_bare_module_factor("pump", pump_type, pump_mat, fp)
    warnings.extend(fbm_warn)
    
    factor = cepci_factor(cepci)
    p_cur = p_2001 * factor
    cbm_cur = p_cur * fbm
    
    return EquipmentCost(
        tag=tag,
        eq_type="Pump",
        size_param=f"P={power_kw:.2f} kW",
        size_value=power_kw,
        purchase_cost_2001=p_2001,
        purchase_cost_current=p_cur,
        bare_module_cost=cbm_cur,
        installed_cost=cbm_cur * TOTAL_CAPEX_FACTOR,
        notes=f"Type: {pump_type}, Material: {pump_mat}, Fp={fp:.2f}",
        warnings=warnings,
    )


def cost_compressor(
    tag: str,
    power_kw: float,
    cepci: float = CEPCI_2024,
    material_factor: float | str = 1.0,
    material: Optional[str] = None,
    compressor_type: str = "centrifugal",
    pressure_barg: float = 0.0,
) -> EquipmentCost:
    warnings = []
    comp_mat, _ = _resolve_material_and_factor("compressor", compressor_type, material_factor, material)
    
    p_2001, p_warn = calculate_purchased_cost_c0("compressor", compressor_type, power_kw)
    warnings.extend(p_warn)
    
    fbm, fbm_warn = calculate_bare_module_factor("compressor", compressor_type, comp_mat, 1.0)
    warnings.extend(fbm_warn)
    
    factor = cepci_factor(cepci)
    p_cur = p_2001 * factor
    cbm_cur = p_cur * fbm
    
    return EquipmentCost(
        tag=tag,
        eq_type="Compressor",
        size_param=f"P={power_kw:.2f} kW",
        size_value=power_kw,
        purchase_cost_2001=p_2001,
        purchase_cost_current=p_cur,
        bare_module_cost=cbm_cur,
        installed_cost=cbm_cur * TOTAL_CAPEX_FACTOR,
        notes=f"Type: {compressor_type}, Material: {comp_mat}",
        warnings=warnings,
    )


def cost_reactor(
    tag: str,
    volume_m3: float,
    cepci: float = CEPCI_2024,
    material_factor: float | str = 1.0,
    material: Optional[str] = None,
    reactor_type: str = "vertical",
    pressure_barg: float = 0.0,
    diameter_m: float = 1.0,
) -> EquipmentCost:
    warnings = []
    react_mat, _ = _resolve_material_and_factor("vessel", reactor_type, material_factor, material)
    
    p_2001, p_warn = calculate_purchased_cost_c0("vessel", reactor_type, volume_m3)
    warnings.extend(p_warn)
    
    fp = calculate_pressure_factor_vessel(pressure_barg, diameter_m)
    
    fbm, fbm_warn = calculate_bare_module_factor("vessel", reactor_type, react_mat, fp, diameter_m, pressure_barg)
    warnings.extend(fbm_warn)
    
    factor = cepci_factor(cepci)
    p_cur = p_2001 * factor
    cbm_cur = p_cur * fbm
    
    return EquipmentCost(
        tag=tag,
        eq_type="Reactor",
        size_param=f"V={volume_m3:.2f} m3",
        size_value=volume_m3,
        purchase_cost_2001=p_2001,
        purchase_cost_current=p_cur,
        bare_module_cost=cbm_cur,
        installed_cost=cbm_cur * TOTAL_CAPEX_FACTOR,
        notes=f"Reactor vessel ({reactor_type}), Material: {react_mat}, Fp={fp:.2f}",
        warnings=warnings,
    )


def cost_generic(
    tag: str,
    eq_type: str,
    cepci: float = CEPCI_2024,
) -> EquipmentCost:
    """Lump-sum estimate for mixers, splitters, valves, etc."""
    c2001 = 10_000.0
    factor = cepci_factor(cepci)
    c_cur = c2001 * factor
    fbm = LANG_FACTORS.get(eq_type, LANG_FACTORS["Other"])
    return EquipmentCost(
        tag=tag,
        eq_type=eq_type,
        size_param="N/A",
        size_value=0.0,
        purchase_cost_2001=c2001,
        purchase_cost_current=c_cur,
        bare_module_cost=c_cur * fbm,
        installed_cost=c_cur * fbm * TOTAL_CAPEX_FACTOR,
        notes=f"Lump-sum estimate for {eq_type}",
    )


def cost_horizontal_vessel(
    tag: str,
    volume_m3: float,
    cepci: float = CEPCI_2024,
    material_factor: float | str = 1.0,
    material: Optional[str] = None,
    pressure_barg: float = 0.0,
    diameter_m: float = 1.0,
) -> EquipmentCost:
    warnings = []
    vess_mat, _ = _resolve_material_and_factor("vessel", "horizontal", material_factor, material)
    
    p_2001, p_warn = calculate_purchased_cost_c0("vessel", "horizontal", volume_m3)
    warnings.extend(p_warn)
    
    fp = calculate_pressure_factor_vessel(pressure_barg, diameter_m)
    
    fbm, fbm_warn = calculate_bare_module_factor("vessel", "horizontal", vess_mat, fp, diameter_m, pressure_barg)
    warnings.extend(fbm_warn)
    
    factor = cepci_factor(cepci)
    p_cur = p_2001 * factor
    cbm_cur = p_cur * fbm
    
    return EquipmentCost(
        tag=tag,
        eq_type="Vessel (Horizontal)",
        size_param=f"V={volume_m3:.2f} m3",
        size_value=volume_m3,
        purchase_cost_2001=p_2001,
        purchase_cost_current=p_cur,
        bare_module_cost=cbm_cur,
        installed_cost=cbm_cur * TOTAL_CAPEX_FACTOR,
        notes=f"Horizontal vessel, Material: {vess_mat}, Fp={fp:.2f}",
        warnings=warnings,
    )


def cost_vertical_vessel(
    tag: str,
    volume_m3: float,
    cepci: float = CEPCI_2024,
    material_factor: float | str = 1.0,
    material: Optional[str] = None,
    pressure_barg: float = 0.0,
    diameter_m: float = 1.0,
) -> EquipmentCost:
    warnings = []
    vess_mat, _ = _resolve_material_and_factor("vessel", "vertical", material_factor, material)
    
    p_2001, p_warn = calculate_purchased_cost_c0("vessel", "vertical", volume_m3)
    warnings.extend(p_warn)
    
    fp = calculate_pressure_factor_vessel(pressure_barg, diameter_m)
    
    fbm, fbm_warn = calculate_bare_module_factor("vessel", "vertical", vess_mat, fp, diameter_m, pressure_barg)
    warnings.extend(fbm_warn)
    
    factor = cepci_factor(cepci)
    p_cur = p_2001 * factor
    cbm_cur = p_cur * fbm
    
    return EquipmentCost(
        tag=tag,
        eq_type="Vessel (Vertical)",
        size_param=f"V={volume_m3:.2f} m3",
        size_value=volume_m3,
        purchase_cost_2001=p_2001,
        purchase_cost_current=p_cur,
        bare_module_cost=cbm_cur,
        installed_cost=cbm_cur * TOTAL_CAPEX_FACTOR,
        notes=f"Vertical vessel, Material: {vess_mat}, Fp={fp:.2f}",
        warnings=warnings,
    )


def cost_kettle_reboiler(
    tag: str,
    area_m2: float,
    cepci: float = CEPCI_2024,
    material_factor: float | str = 1.0,
    material: Optional[str] = None,
    pressure_barg: float = 0.0,
) -> EquipmentCost:
    ec = cost_heat_exchanger(tag, area_m2, cepci, material_factor, material, "kettle_reboiler", pressure_barg)
    ec.eq_type = "Kettle Reboiler"
    return ec


# ---------------------------------------------------------------------------
# OPEX utility cost functions
# ---------------------------------------------------------------------------

def annual_utility_cost(
    duty_kw: float,
    utility_name: str,
    operating_hours: float = 8000.0,
    price_override: float | None = None,
) -> float:
    """
    Calculate annual utility cost in USD/year.
    duty_kw: thermal/electrical duty in kW
    utility_name: key in UTILITY_COSTS
    """
    if price_override is not None:
        rate = price_override
    else:
        rate = UTILITY_COSTS.get(utility_name, 8.0)   # $/GJ
    duty_gj_per_hr = duty_kw * 3600 / 1e6          # GJ per hour
    annual_gj = duty_gj_per_hr * operating_hours
    return annual_gj * rate


def electricity_cost_annual(
    power_kw: float,
    elec_price_per_kwh: float = 0.06,
    operating_hours: float = 8000.0,
) -> float:
    """Annual electricity cost in USD."""
    return power_kw * elec_price_per_kwh * operating_hours


def calculate_lmtd(t_in_hot: float, t_out_hot: float, t_in_cold: float, t_out_cold: float) -> float:
    """Calculate Log Mean Temperature Difference (LMTD) in Kelvin."""
    dt1 = t_in_hot - t_out_cold
    dt2 = t_out_hot - t_in_cold
    # Prevent divide by zero or negative values
    if dt1 <= 0 or dt2 <= 0:
        return 10.0 # fallback default temperature difference
    if abs(dt1 - dt2) < 1e-3:
        return dt1
    return (dt1 - dt2) / math.log(dt1 / dt2)


DEFAULT_UTILITY_DATABASE = {
    "Cooling Water": {
        "type": "cooling",
        "Tin_C": 30.0,
        "Tout_C": 40.0,
        "U_W_m2K": 700.0,
        "price_USD_GJ": 0.354
    },
    "Chilled Water": {
        "type": "cooling",
        "Tin_C": 5.0,
        "Tout_C": 15.0,
        "U_W_m2K": 500.0,
        "price_USD_GJ": 4.77
    },
    "Refrigeration": {
        "type": "cooling",
        "Tin_C": -25.0,
        "Tout_C": -15.0,
        "U_W_m2K": 300.0,
        "price_USD_GJ": 7.89
    },
    "LP Steam": {
        "type": "heating",
        "Tsat_C": 152.0,
        "U_W_m2K": 1000.0,
        "price_USD_GJ": 6.08
    },
    "MP Steam": {
        "type": "heating",
        "Tsat_C": 184.0,
        "U_W_m2K": 900.0,
        "price_USD_GJ": 8.22
    },
    "HP Steam": {
        "type": "heating",
        "Tsat_C": 254.0,
        "U_W_m2K": 800.0,
        "price_USD_GJ": 9.83
    }
}


def select_cooling_utility(process_temp_c: float, db: dict | None = None) -> tuple[str, tuple[float, float], float, float]:
    """
    Select cooling utility based on process temperature and return:
    (name, (Tin, Tout), U_W_m2K, price_USD_GJ)
    """
    if db is None:
        db = DEFAULT_UTILITY_DATABASE

    if process_temp_c > 45.0:
        name = "Cooling Water"
    elif process_temp_c >= 10.0:
        name = "Chilled Water"
    else:
        name = "Refrigeration"

    u_info = db.get(name, {})
    tin = u_info.get("Tin_C", 30.0)
    tout = u_info.get("Tout_C", 40.0)
    u_val = u_info.get("U_W_m2K", 700.0)
    price = u_info.get("price_USD_GJ", 0.354)
    return name, (tin, tout), u_val, price


def select_heating_utility(process_temp_c: float, db: dict | None = None) -> tuple[str, tuple[float, float], float, float]:
    """
    Select heating utility based on process temperature and return:
    (name, (Tsat, Tsat), U_W_m2K, price_USD_GJ)
    """
    if db is None:
        db = DEFAULT_UTILITY_DATABASE

    if process_temp_c < 180.0:
        name = "LP Steam"
    elif process_temp_c < 220.0:
        name = "MP Steam"
    else:
        name = "HP Steam"

    u_info = db.get(name, {})
    tsat = u_info.get("Tsat_C", 184.0)
    u_val = u_info.get("U_W_m2K", 900.0)
    price = u_info.get("price_USD_GJ", 8.22)
    return name, (tsat, tsat), u_val, price
