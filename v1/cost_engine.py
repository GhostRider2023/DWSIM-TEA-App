"""
cost_engine.py
Orchestrates CAPEX and OPEX estimation from parsed DWSIM data.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

from equipment_models import (
    EquipmentCost,
    UTILITY_COSTS, HEAT_UTILITY_MAP, TOTAL_CAPEX_FACTOR,
    CEPCI_2024,
    cost_distillation_column,
    cost_heat_exchanger,
    cost_pump,
    cost_compressor,
    cost_reactor,
    cost_generic,
    annual_utility_cost,
    electricity_cost_annual,
    cepci_factor,
    cost_horizontal_vessel,
    cost_vertical_vessel,
    cost_kettle_reboiler,
    calculate_lmtd,
    select_cooling_utility,
    select_heating_utility,
)
from excel_parser import DWSIMData, _to_float


# ---------------------------------------------------------------------------
# Settings dataclass (user-editable in Streamlit sidebar)
# ---------------------------------------------------------------------------

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


@dataclass
class CostSettings:
    cepci: float = CEPCI_2024
    material_factor: float = 1.0          # 1.0 = carbon steel
    operating_hours: float = 8000.0        # hr/year
    elec_price_per_kwh: float = 0.06       # $/kWh
    steam_utility: str = "MP Steam (10 bar)"
    cooling_utility: str = "Cooling Water"
    contingency_pct: float = 15.0          # % of bare-module cost
    engineering_pct: float = 10.0          # % of bare-module cost
    startup_pct: float = 5.0              # % of total fixed capital
    working_capital_pct: float = 10.0     # % of total fixed capital
    currency_symbol: str = "USD"
    location_factor: float = 1.0           # Country-specific location factor
    currency_conversion: float = 1.0       # USD to local currency factor
    plant_life_years: float = 10.0
    discount_rate_pct: float = 10.0
    tax_rate_pct: float = 0.0             # income tax rate %; 0 = pre-tax only
    # FIX #22: OPEX add-ons
    # labour_cost_annual=None â use N^0.7 heuristic; set to a number to override
    labour_cost_override: Optional[float] = None   # $/year explicit override
    operator_salary: float = 60_000.0     # $/operator-position/year
    shifts_per_position: float = 5.0      # shift operators per position
    maintenance_pct: float = 6.0           # % of TFC/year
    overhead_pct: float = 60.0            # % of labour/year
    insurance_pct: float = 3.0            # % of FCI/year (used in some models)
    raw_material_cost: float = 0.0        # $/year
    opex_method: str = "Turton (COMd)"    # "Turton (COMd)" or "Simple Addition"
    capex_mode: str = "Rigorous (System Expansion)"  # "Rigorous (System Expansion)" or "Strict (Legacy Turton)"
    pump_material: str = "CS_carbon_steel"
    hx_material: str = "CS/CS"
    vessel_material: str = "CS_carbon_steel"
    compressor_material: str = "CS_carbon_steel"
    tray_material: str = "CS_carbon_steel"
    tray_type: str = "sieve"
    reflux_drum_residence_time_s: float = 300.0
    reflux_drum_surge_factor: float = 2.0
    utility_database: dict = field(default_factory=lambda: {k: dict(v) for k, v in DEFAULT_UTILITY_DATABASE.items()})


# ---------------------------------------------------------------------------
# Results container
# ---------------------------------------------------------------------------

@dataclass
class CostReport:
    equipment_costs: list[EquipmentCost] = field(default_factory=list)
    # CAPEX summary (all in USD)
    total_purchase_cost: float = 0.0
    total_bare_module: float = 0.0
    contingency: float = 0.0
    engineering_fee: float = 0.0
    isbl: float = 0.0       # Inside Battery Limits
    osbl_factor: float = 0.10
    osbl: float = 0.0
    total_fixed_capital: float = 0.0
    startup: float = 0.0
    working_capital: float = 0.0
    total_project_cost: float = 0.0
    # OPEX summary (all in USD/year)
    utility_rows: list[dict] = field(default_factory=list)
    total_utility_opex: float = 0.0
    labour_opex: float = 0.0
    maintenance_opex: float = 0.0
    overhead_opex: float = 0.0
    insurance_opex: float = 0.0
    total_opex: float = 0.0
    # Profitability
    annualised_capex: float = 0.0
    total_annual_cost: float = 0.0
    # FIX #22: After-tax fields
    n_operators: int = 0
    depreciation_annual: float = 0.0
    npv_after_tax: Optional[float] = None
    warnings: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Core engine
# ---------------------------------------------------------------------------

class CostEngine:
    def __init__(self, data: DWSIMData, settings: CostSettings | None = None):
        self.data = data
        self.settings = settings or CostSettings()

    def _get_hx_process_info(self, hx_tag: str) -> tuple[str, float, float]:
        if self.data.connections is None or self.data.connections.empty:
            return "heating", 100.0, 100.0
        if self.data.streams is None or self.data.streams.empty:
            return "heating", 100.0, 100.0
            
        df_conn = self.data.connections
        hx_conns = df_conn[
            (df_conn["From"] == hx_tag) | 
            (df_conn["To"] == hx_tag) | 
            (df_conn["ObjectTag"] == hx_tag)
        ]
        
        inlets = []
        outlets = []
        for _, row in hx_conns.iterrows():
            st = row.get("StreamTag") or row.get("OtherTag")
            dir_ = row.get("Direction")
            if st:
                if dir_ == "Inlet":
                    inlets.append(str(st).strip())
                elif dir_ == "Outlet":
                    outlets.append(str(st).strip())
        
        df_str = self.data.streams
        t_in = 25.0
        t_out = 25.0
        
        if inlets:
            t_in_vals = pd.to_numeric(df_str[df_str["Tag"].isin(inlets)]["Temperature_C"], errors="coerce").dropna()
            if not t_in_vals.empty:
                t_in = float(t_in_vals.mean())
        if outlets:
            t_out_vals = pd.to_numeric(df_str[df_str["Tag"].isin(outlets)]["Temperature_C"], errors="coerce").dropna()
            if not t_out_vals.empty:
                t_out = float(t_out_vals.mean())
                
        if t_out < t_in:
            return "cooling", t_in, t_out
        else:
            return "heating", t_in, t_out


    # ------------------------------------------------------------------
    def estimate(self) -> CostReport:
        report = CostReport()
        s = self.settings

        # Helper to get top/bottom temperatures for a column
        def get_column_temps(col_tag: str) -> tuple[float, float]:
            df = self.data.column_stage_profiles
            if df is not None and not df.empty and "Tag" in df.columns:
                df_col = df[df["Tag"] == col_tag]
                if not df_col.empty:
                    df_col = df_col.copy()
                    df_col["Stage"] = pd.to_numeric(df_col["Stage"], errors="coerce")
                    df_col = df_col.dropna(subset=["Stage"]).sort_values(by="Stage")
                    if not df_col.empty:
                        t_top = df_col.iloc[0].get("Temperature_C")
                        t_bot = df_col.iloc[-1].get("Temperature_C")
                        try:
                            return float(t_top), float(t_bot)
                        except (ValueError, TypeError):
                            pass
            return 50.0, 150.0

        # Helper to get top/bottom pressures for a column
        def get_column_pressures(col_tag: str) -> tuple[float, float]:
            df = self.data.column_stage_profiles
            if df is not None and not df.empty and "Tag" in df.columns:
                df_col = df[df["Tag"] == col_tag]
                if not df_col.empty:
                    df_col = df_col.copy()
                    df_col["Stage"] = pd.to_numeric(df_col["Stage"], errors="coerce")
                    df_col = df_col.dropna(subset=["Stage"]).sort_values(by="Stage")
                    if not df_col.empty:
                        p_top = df_col.iloc[0].get("Pressure_bar")
                        p_bot = df_col.iloc[-1].get("Pressure_bar")
                        try:
                            return float(p_top), float(p_bot)
                        except (ValueError, TypeError):
                            pass
            return 1.01325, 1.01325



        # Helper to get column reflux and distillate flow rates
        def get_column_flows(col_tag: str, reflux_ratio: float | None) -> tuple[float, float]:
            conns = []
            if self.data.connections is not None and not self.data.connections.empty:
                df_conn = self.data.connections
                col_conns = df_conn[
                    ((df_conn["From"] == col_tag) | (df_conn["ObjectTag"] == col_tag)) &
                    (df_conn["Direction"] == "Outlet")
                ]
                for _, row in col_conns.iterrows():
                    st = row.get("StreamTag") or row.get("OtherTag")
                    if st:
                        conns.append(str(st).strip())
            
            dist_flow = 0.0
            if conns and self.data.streams is not None and not self.data.streams.empty:
                df_str = self.data.streams
                col_streams = df_str[df_str["Tag"].isin(conns)].copy()
                if not col_streams.empty:
                    col_streams["Temperature_C"] = pd.to_numeric(col_streams["Temperature_C"], errors="coerce")
                    col_streams = col_streams.sort_values(by="Temperature_C")
                    dist_stream = col_streams.iloc[0]
                    dist_flow = _to_float(dist_stream.get("VolumetricFlow_m3_s")) or 0.0
            
            if dist_flow <= 0.0:
                dist_flow = 0.002  # fallback
                
            r = reflux_ratio if (reflux_ratio is not None and reflux_ratio > 0) else 1.5
            return r * dist_flow, dist_flow

        inventory = self.data.equipment_inventory()
        virtual_inventory = []

        # 1. Per-equipment purchase costs & system expansion
        for item in inventory:
            tag = item["tag"]
            eq_type = item["type"]

            if eq_type == "Distillation Column" and s.capex_mode == "Rigorous (System Expansion)":
                # Primary Column (Shell + Trays)
                diam = item.get("diameter_m", 1.0) or 1.0
                ht   = item.get("height_m", 3.0) or 3.0
                trays= item.get("n_stages", 10) or 10
                sp   = item.get("tray_spacing_m", 0.6) or 0.6
                
                # Check aspect ratio
                aspect_ratio = ht / diam if diam > 0 else 0.0
                if aspect_ratio < 10.0 or aspect_ratio > 40.0:
                    report.warnings.append(
                        f"â ï¸ {tag}: Height/Diameter ratio ({aspect_ratio:.1f}) is outside standard range [10.0 - 40.0]. Check column sizing."
                    )
                
                # Get column temperatures, pressures and flows
                top_temp, bottom_temp = get_column_temps(tag)
                top_p_bar, bot_p_bar = get_column_pressures(tag)
                top_p_barg = max(0.0, top_p_bar - 1.01325)
                bot_p_barg = max(0.0, bot_p_bar - 1.01325)
                max_col_p_barg = max(top_p_barg, bot_p_barg)

                # Main Column cost
                col_cost = cost_distillation_column(
                    tag, diam, ht, trays, sp, s.cepci, s.material_factor,
                    shell_material=s.vessel_material, tray_material=s.tray_material,
                    tray_type=s.tray_type, pressure_barg=max_col_p_barg
                )
                report.equipment_costs.append(col_cost)
                report.warnings.extend(col_cost.warnings)

                reflux_flow, dist_flow = get_column_flows(tag, item.get("reflux_ratio"))

                # (a) Condenser
                cond_duty = item.get("condenser_duty_kw", 0.0) or 0.0
                if cond_duty > 0:
                    cool_util, cool_temps, cool_u, cool_price = select_cooling_utility(top_temp, s.utility_database)
                    if top_temp < 10.0 and cool_util == "Cooling Water":
                        report.warnings.append(
                            f"❌ {tag}_Condenser: Condenser temperature ({top_temp:.1f}°C) < 10°C with cooling water selected. High risk of freeze-up."
                        )
                    
                    lmtd_cond = calculate_lmtd(top_temp, top_temp, cool_temps[0], cool_temps[1])
                    lmtd_cond = max(0.1, lmtd_cond)
                    
                    # Area calculation: Area = Q / (U * LMTD), converting Q to W
                    area_cond = (cond_duty * 1000.0) / (cool_u * lmtd_cond)
                    if area_cond > 1000.0 or area_cond < 10.0:
                        report.warnings.append(
                            f"⚠️ {tag}_Condenser: Exchanger area ({area_cond:.1f} m²) is outside standard range [10.0 - 1000.0] m². Split or combine shells."
                        )
                        
                    condenser = cost_heat_exchanger(
                        tag + "_Condenser", area_cond, s.cepci, s.material_factor,
                        material=s.hx_material, pressure_barg=top_p_barg
                    )
                    condenser.eq_type = "Condenser"
                    
                    # Populating details dict
                    n_shells = max(1, math.ceil(area_cond / 1000.0))
                    condenser.details = {
                        "duty_kw": cond_duty,
                        "utility_name": cool_util,
                        "utility_temp_str": f"{cool_temps[0]:.1f} -> {cool_temps[1]:.1f}°C",
                        "u_value_w_m2k": cool_u,
                        "lmtd_k": lmtd_cond,
                        "area_m2": area_cond,
                        "n_shells": n_shells,
                        "purchase_cost": condenser.purchase_cost_current,
                        "bare_module_cost": condenser.bare_module_cost,
                        "utility_price_usd_gj": cool_price
                    }
                    
                    # extract actual notes Fp
                    fp_part = condenser.notes.split("Fp=")[1] if "Fp=" in condenser.notes else "1.00"
                    condenser.notes = f"Condenser (Cooling utility: {cool_util}, Material: {s.hx_material}, Fp={fp_part})"
                    report.equipment_costs.append(condenser)
                    report.warnings.extend(condenser.warnings)
                    
                    # Store dynamic utility details for OPEX step
                    virtual_inventory.append({
                        "tag": tag + "_Condenser",
                        "type": "Condenser",
                        "duty_kw": cond_duty,
                        "utility": cool_util,
                        "price": cool_price
                    })

                # (b) Reboiler
                reb_duty = item.get("reboiler_duty_kw", 0.0) or 0.0
                if reb_duty > 0:
                    heat_util, heat_temps, heat_u, heat_price = select_heating_utility(bottom_temp, s.utility_database)
                    
                    lmtd_reb = calculate_lmtd(heat_temps[0], heat_temps[1], bottom_temp, bottom_temp)
                    lmtd_reb = max(0.1, lmtd_reb)
                    
                    # Area calculation: Area = Q / (U * LMTD), converting Q to W
                    area_reb = (reb_duty * 1000.0) / (heat_u * lmtd_reb)
                    if area_reb > 1000.0 or area_reb < 10.0:
                        report.warnings.append(
                            f"⚠️ {tag}_Reboiler: Exchanger area ({area_reb:.1f} m²) is outside standard range [10.0 - 1000.0] m². Split or combine shells."
                        )
                        
                    reboiler = cost_kettle_reboiler(
                        tag + "_Reboiler", area_reb, s.cepci, s.material_factor,
                        material=s.hx_material, pressure_barg=bot_p_barg
                    )
                    reboiler.eq_type = "Reboiler"
                    
                    # Populating details dict
                    n_shells = max(1, math.ceil(area_reb / 1000.0))
                    reboiler.details = {
                        "duty_kw": reb_duty,
                        "utility_name": heat_util,
                        "utility_temp_str": f"{heat_temps[0]:.1f}°C (Saturated)",
                        "u_value_w_m2k": heat_u,
                        "lmtd_k": lmtd_reb,
                        "area_m2": area_reb,
                        "n_shells": n_shells,
                        "purchase_cost": reboiler.purchase_cost_current,
                        "bare_module_cost": reboiler.bare_module_cost,
                        "utility_price_usd_gj": heat_price
                    }
                    
                    fp_part = reboiler.notes.split("Fp=")[1] if "Fp=" in reboiler.notes else "1.00"
                    reboiler.notes = f"Kettle Reboiler (Heating utility: {heat_util}, Material: {s.hx_material}, Fp={fp_part})"
                    report.equipment_costs.append(reboiler)
                    report.warnings.extend(reboiler.warnings)
                    
                    virtual_inventory.append({
                        "tag": tag + "_Reboiler",
                        "type": "Reboiler",
                        "duty_kw": reb_duty,
                        "utility": heat_util,
                        "price": heat_price
                    })

                # (c) Reflux Drum
                drum_vol = s.reflux_drum_residence_time_s * (reflux_flow + dist_flow) * s.reflux_drum_surge_factor
                drum = cost_horizontal_vessel(
                    tag + "_Reflux_Drum", drum_vol, s.cepci, s.material_factor,
                    material=s.vessel_material, pressure_barg=top_p_barg, diameter_m=1.0
                )
                drum.eq_type = "Reflux Drum"
                fp_part = drum.notes.split("Fp=")[1] if "Fp=" in drum.notes else "1.00"
                drum.notes = f"Horizontal reflux accumulator, volume = {drum_vol:.2f} mÂ³, Material: {s.vessel_material}, Fp={fp_part}"
                report.equipment_costs.append(drum)
                report.warnings.extend(drum.warnings)

                # (d) Reflux Pump
                pump_pwr = max(0.1, reflux_flow * 500.0)
                pump = cost_pump(
                    tag + "_Reflux_Pump", pump_pwr, s.cepci, s.material_factor,
                    material=s.pump_material, pressure_barg=top_p_barg
                )
                pump.eq_type = "Reflux Pump"
                fp_part = pump.notes.split("Fp=")[1] if "Fp=" in pump.notes else "1.00"
                pump.notes = f"Centrifugal reflux pump, power = {pump_pwr:.2f} kW, Material: {s.pump_material}, Fp={fp_part}"
                report.equipment_costs.append(pump)
                report.warnings.extend(pump.warnings)
                
                virtual_inventory.append({
                    "tag": tag + "_Reflux_Pump",
                    "type": "Pump",
                    "power_kw": pump_pwr,
                    "utility": "Electricity",
                })

            else:
                # Normal costing (legacy or not a column)
                ec = self._cost_one_item(item)
                if ec is not None:
                    # In Strict mode, add warning for Distillation Column
                    if eq_type == "Distillation Column":
                        report.warnings.append(
                            f"â ï¸ {tag}: Operating in Strict CAPCOST mode. Auxiliary equipment (Condenser, Reboiler, Reflux Drum, Pumps) are not costed."
                        )
                        
                    # Standard check for standalone heat exchangers
                    if eq_type == "Heat Exchanger":
                        area = item.get("area_m2", 10.0) or 10.0
                        if area > 1000.0 or area < 10.0:
                            report.warnings.append(
                                f"â ï¸ {tag}: Exchanger area ({area:.1f} mÂ²) is outside standard range [10.0 - 1000.0] mÂ². Split or combine shells."
                            )
                    report.equipment_costs.append(ec)
                    report.warnings.extend(ec.warnings)

        # Apply Location Factor and Currency Conversion to all Equipment Costs
        for ec in report.equipment_costs:
            ec.purchase_cost_current *= s.location_factor * s.currency_conversion
            ec.bare_module_cost *= s.location_factor * s.currency_conversion
            ec.installed_cost *= s.location_factor * s.currency_conversion

        # 2. CAPEX roll-up
        report.total_purchase_cost = sum(e.purchase_cost_current for e in report.equipment_costs)
        report.total_bare_module   = sum(e.bare_module_cost       for e in report.equipment_costs)

        if s.capex_mode == "Rigorous (System Expansion)":
            # Rigorous grassroots investment roll-up
            report.isbl = 1.18 * report.total_bare_module  # ISBL is Total Module Cost C_TM
            report.contingency = 0.15 * report.total_bare_module
            report.engineering_fee = 0.03 * report.total_bare_module
            report.osbl = 0.0
            report.total_fixed_capital = report.isbl + 0.35 * report.total_purchase_cost
        else:
            # Legacy roll-up
            report.contingency    = report.total_bare_module * s.contingency_pct / 100
            report.engineering_fee= report.total_bare_module * s.engineering_pct / 100
            report.isbl = report.total_bare_module + report.contingency + report.engineering_fee
            report.osbl = report.isbl * report.osbl_factor
            report.total_fixed_capital = report.isbl + report.osbl

        report.startup         = report.total_fixed_capital * s.startup_pct / 100
        report.working_capital = report.total_fixed_capital * s.working_capital_pct / 100
        report.total_project_cost = report.total_fixed_capital + report.startup + report.working_capital

        # 3. OPEX roll-up
        # Estimate utilities (including dynamically generated virtual ones)
        report.utility_rows = self._estimate_utilities_dynamic(report.equipment_costs, virtual_inventory)
        
        # Scale utilities by currency conversion
        for row in report.utility_rows:
            row["annual_cost_usd"] *= s.currency_conversion
            
        report.total_utility_opex = sum(r["annual_cost_usd"] for r in report.utility_rows)
        
        # Plant-size-aware labor using N^0.7 heuristic (or override)
        n_eq = len(report.equipment_costs)
        if s.labour_cost_override is not None and s.labour_cost_override > 0:
            report.labour_opex = float(s.labour_cost_override) * s.currency_conversion
            report.n_operators = 0
        else:
            import math as _math
            n_op = max(1, _math.ceil(0.5 * max(n_eq, 1) ** 0.7))
            total_op = n_op * s.shifts_per_position
            report.labour_opex = total_op * s.operator_salary * s.currency_conversion
            report.n_operators = n_op

        scaled_raw_material_cost = s.raw_material_cost * s.currency_conversion

        if s.opex_method == "Turton (COMd)":
            # Scale utilities and labor as per Turton COMd breakdown
            report.labour_opex = 2.73 * report.labour_opex
            report.total_utility_opex = 1.23 * report.total_utility_opex
            
            # Maintenance and Insurance & Tax are 6% and 3% of FCI respectively
            report.maintenance_opex = 0.06 * report.total_fixed_capital
            report.insurance_opex = 0.03 * report.total_fixed_capital
            
            # Overhead is 60% of (Labour + Maintenance)
            report.overhead_opex = 0.60 * (report.labour_opex + report.maintenance_opex)
            
            report.total_opex = (
                report.total_utility_opex + report.labour_opex
                + report.maintenance_opex + report.overhead_opex
                + report.insurance_opex + scaled_raw_material_cost * 1.23
            )
        else:
            # Simple Addition mode (matching opex_calculator.py logic)
            report.maintenance_opex = report.total_fixed_capital * s.maintenance_pct / 100
            report.insurance_opex = report.total_fixed_capital * s.insurance_pct / 100
            report.overhead_opex = (report.labour_opex + report.maintenance_opex) * s.overhead_pct / 100
            
            report.total_opex = (
                report.total_utility_opex + report.labour_opex
                + report.maintenance_opex + report.overhead_opex
                + report.insurance_opex + scaled_raw_material_cost
            )

        # 4. Annualised CAPEX (capital recovery factor)
        r = s.discount_rate_pct / 100
        n = s.plant_life_years
        if r > 0:
            crf = r * (1 + r) ** n / ((1 + r) ** n - 1)
        else:
            crf = 1.0 / n
        report.annualised_capex = report.total_fixed_capital * crf
        report.total_annual_cost = report.annualised_capex + report.total_opex

        # After-tax NPV with straight-line depreciation
        tr = s.tax_rate_pct / 100.0
        if tr > 0:
            depr = report.total_fixed_capital / n
            report.depreciation_annual = depr
            report.warnings.append(
                f"After-tax mode enabled ({s.tax_rate_pct:.1f}%); "
                f"supply annual_revenue to compute after-tax NPV in the Streamlit app."
            )

        # Unified Warnings & Validation Rules
        for ec in report.equipment_costs:
            if ec.eq_type in ("Heat Exchanger", "Condenser", "Reboiler", "Kettle Reboiler"):
                d = ec.details
                if d:
                    area = d.get("area_m2", 0.0)
                    lmtd = d.get("lmtd_k", 0.0)
                    u_val = d.get("u_value_w_m2k", 0.0)
                    
                    if area > 1000.0:
                        w = f"⚠️ {ec.tag}: Exchanger area ({area:.1f} m²) is outside standard range [10.0 - 1000.0] m². Split or combine shells."
                        if w not in report.warnings:
                            report.warnings.append(w)
                            
                    if lmtd < 10.0:
                        report.warnings.append(
                            f"⚠️ {ec.tag}: LMTD ({lmtd:.1f} K) is below 10 K. Sizing may be uneconomical due to low thermal driving force."
                        )
                        
                    if u_val < 100.0 or u_val > 2000.0:
                        report.warnings.append(
                            f"⚠️ {ec.tag}: Overall heat transfer coefficient U ({u_val:.0f} W/m²K) is outside typical range [100 - 2000] W/m²K."
                        )
            
            # Reboiler duty check
            if ec.eq_type in ("Reboiler", "Kettle Reboiler"):
                d = ec.details
                if d:
                    duty_kw = d.get("duty_kw", 0.0)
                    if duty_kw > 20000.0: # 20 MW
                        report.warnings.append(
                            f"⚠️ {ec.tag}: Reboiler duty is unusually high ({duty_kw/1000.0:.1f} MW). Verify column thermal profile."
                        )

        return report

    # ------------------------------------------------------------------
    def _get_equipment_pressure_barg(self, tag: str) -> float:
        """Find the maximum pressure in barg among all connected material streams."""
        if self.data.connections is None or self.data.connections.empty:
            return 0.0
        if self.data.streams is None or self.data.streams.empty:
            return 0.0
            
        df_conn = self.data.connections
        eq_conns = df_conn[
            (df_conn["From"] == tag) | 
            (df_conn["To"] == tag) | 
            (df_conn["ObjectTag"] == tag)
        ]
        
        stream_tags = []
        for _, row in eq_conns.iterrows():
            st = row.get("StreamTag") or row.get("OtherTag")
            if st:
                stream_tags.append(str(st).strip())
                
        if not stream_tags:
            return 0.0
            
        df_str = self.data.streams
        eq_streams = df_str[df_str["Tag"].isin(stream_tags)]
        if eq_streams.empty:
            return 0.0
            
        pressures = pd.to_numeric(eq_streams["Pressure_bar"], errors="coerce").dropna()
        if pressures.empty:
            return 0.0
            
        max_p_bar = float(pressures.max())
        return max(0.0, max_p_bar - 1.01325)

    # ------------------------------------------------------------------
    def _cost_one_item(self, item: dict) -> EquipmentCost | None:
        s = self.settings
        tag = item["tag"]
        eq_type = item["type"]
        p_barg = self._get_equipment_pressure_barg(tag)

        if eq_type == "Distillation Column":
            diam = item.get("diameter_m", 1.0) or 1.0
            ht   = item.get("height_m", 3.0) or 3.0
            trays= item.get("n_stages", 10) or 10
            sp   = item.get("tray_spacing_m", 0.6) or 0.6
            return cost_distillation_column(
                tag, diam, ht, trays, sp, s.cepci, s.material_factor,
                shell_material=s.vessel_material, tray_material=s.tray_material,
                tray_type=s.tray_type, pressure_barg=p_barg
            )

        elif eq_type == "Heat Exchanger":
            area = item.get("area_m2", 10.0) or 10.0
            duty_kw = item.get("duty_kw", 0.0) or 0.0
            p_barg = self._get_equipment_pressure_barg(tag)
            
            h_type, t_in, t_out = self._get_hx_process_info(tag)
            if h_type == "cooling":
                cool_util, cool_temps, cool_u, cool_price = select_cooling_utility(t_out, s.utility_database)
                lmtd_val = calculate_lmtd(t_in, t_out, cool_temps[0], cool_temps[1])
                lmtd_val = max(0.1, lmtd_val)
                u_val = cool_u
                util_name = cool_util
                util_price = cool_price
                util_temp_str = f"{cool_temps[0]:.1f} -> {cool_temps[1]:.1f}°C"
            else:
                heat_util, heat_temps, heat_u, heat_price = select_heating_utility(t_out, s.utility_database)
                lmtd_val = calculate_lmtd(heat_temps[0], heat_temps[1], t_in, t_out)
                lmtd_val = max(0.1, lmtd_val)
                u_val = heat_u
                util_name = heat_util
                util_price = heat_price
                util_temp_str = f"{heat_temps[0]:.1f}°C (Saturated)"
            
            # If area is not provided but duty is, size it!
            if (area is None or area <= 0.0) and duty_kw > 0.0:
                area = (duty_kw * 1000.0) / (u_val * lmtd_val)
                
            ec = cost_heat_exchanger(
                tag, area, s.cepci, s.material_factor,
                material=s.hx_material, pressure_barg=p_barg
            )
            ec.details = {
                "duty_kw": duty_kw,
                "utility_name": util_name,
                "utility_temp_str": util_temp_str,
                "u_value_w_m2k": u_val,
                "lmtd_k": lmtd_val,
                "area_m2": area,
                "n_shells": max(1, math.ceil(area / 1000.0)),
                "purchase_cost": ec.purchase_cost_current,
                "bare_module_cost": ec.bare_module_cost,
                "utility_price_usd_gj": util_price
            }
            return ec

        elif eq_type == "Pump":
            pwr = item.get("power_kw", 1.0) or 1.0
            return cost_pump(
                tag, pwr, s.cepci, s.material_factor,
                material=s.pump_material, pressure_barg=p_barg
            )

        elif eq_type == "Compressor":
            pwr = item.get("power_kw", 10.0) or 10.0
            return cost_compressor(
                tag, pwr, s.cepci, s.material_factor,
                material=s.compressor_material, pressure_barg=p_barg
            )

        elif "Reactor" in eq_type:
            vol = item.get("volume_m3", 1.0) or 1.0
            diam = self.data.get_costing_value(tag, "EstimatedDiameter", "Diameter", "VesselDiameter") or 1.0
            return cost_reactor(
                tag, vol, s.cepci, s.material_factor,
                material=s.vessel_material, pressure_barg=p_barg, diameter_m=diam
            )

        elif any(k in eq_type.lower() for k in ("vessel", "separator", "tank", "drum", "accumulator")):
            vol = item.get("volume_m3") or self.data.get_costing_value(tag, "Volume", "VesselVolume", "Vessel_Volume") or 1.0
            diam = self.data.get_costing_value(tag, "EstimatedDiameter", "Diameter", "VesselDiameter") or 1.0
            vess_type = "horizontal" if "horizontal" in eq_type.lower() or "tank" in eq_type.lower() else "vertical"
            if vess_type == "horizontal":
                return cost_horizontal_vessel(
                    tag, vol, s.cepci, s.material_factor,
                    material=s.vessel_material, pressure_barg=p_barg, diameter_m=diam
                )
            else:
                return cost_vertical_vessel(
                    tag, vol, s.cepci, s.material_factor,
                    material=s.vessel_material, pressure_barg=p_barg, diameter_m=diam
                )

        else:
            # Mixer, Splitter, Valve, Custom, etc.
            return cost_generic(tag, eq_type, s.cepci)

    # ------------------------------------------------------------------
    def _estimate_utilities(self, equipment_costs: list[EquipmentCost]) -> list[dict]:
        return self._estimate_utilities_dynamic(equipment_costs, [])

    # ------------------------------------------------------------------
    def _estimate_utilities_dynamic(self, equipment_costs: list[EquipmentCost], virtual_inventory: list[dict]) -> list[dict]:
        """Build per-stream utility cost rows from distillation + HX duties."""
        s = self.settings
        rows = []
        inv = self.data.equipment_inventory()

        for item in inv:
            tag = item["tag"]
            eq_type = item["type"]

            if eq_type == "Distillation Column":
                if s.capex_mode == "Rigorous (System Expansion)":
                    continue
                else:
                    # Condenser â cooling utility
                    cond = item.get("condenser_duty_kw", 0.0) or 0.0
                    if cond > 0:
                        rows.append({
                            "tag": tag,
                            "duty_type": "Condenser (cooling)",
                            "utility": s.cooling_utility,
                            "duty_kw": cond,
                            "annual_cost_usd": annual_utility_cost(cond, s.cooling_utility, s.operating_hours),
                        })
                    # Reboiler â steam utility
                    reb = item.get("reboiler_duty_kw", 0.0) or 0.0
                    if reb > 0:
                        rows.append({
                            "tag": tag,
                            "duty_type": "Reboiler (heating)",
                            "utility": s.steam_utility,
                            "duty_kw": reb,
                            "annual_cost_usd": annual_utility_cost(reb, s.steam_utility, s.operating_hours),
                        })

            elif eq_type == "Heat Exchanger":
                duty = item.get("duty_kw", 0.0) or 0.0
                if duty > 0:
                    ec_match = next((ec for ec in equipment_costs if ec.tag == tag), None)
                    if ec_match and ec_match.details:
                        ut_name = ec_match.details.get("utility_name")
                        ut_price = ec_match.details.get("utility_price_usd_gj")
                        rows.append({
                            "tag": tag,
                            "duty_type": "Heat Exchanger",
                            "utility": ut_name,
                            "duty_kw": duty,
                            "annual_cost_usd": annual_utility_cost(duty, ut_name, s.operating_hours, price_override=ut_price),
                        })
                    else:
                        rows.append({
                            "tag": tag,
                            "duty_type": "Heat Exchanger",
                            "utility": s.steam_utility,
                            "duty_kw": duty,
                            "annual_cost_usd": annual_utility_cost(duty, s.steam_utility, s.operating_hours),
                        })

            elif eq_type in ("Pump", "Compressor"):
                pwr = item.get("power_kw", 0.0) or 0.0
                if pwr > 0:
                    rows.append({
                        "tag": tag,
                        "duty_type": f"{eq_type} (electricity)",
                        "utility": "Electricity",
                        "duty_kw": pwr,
                        "annual_cost_usd": electricity_cost_annual(pwr, s.elec_price_per_kwh, s.operating_hours),
                    })

        # Add virtual inventory utilities
        for v in virtual_inventory:
            v_tag = v["tag"]
            v_type = v["type"]
            if v_type == "Condenser":
                rows.append({
                    "tag": v_tag,
                    "duty_type": "Condenser (cooling)",
                    "utility": v["utility"],
                    "duty_kw": v["duty_kw"],
                    "annual_cost_usd": annual_utility_cost(v["duty_kw"], v["utility"], s.operating_hours, price_override=v.get("price")),
                })
            elif v_type == "Reboiler":
                rows.append({
                    "tag": v_tag,
                    "duty_type": "Reboiler (heating)",
                    "utility": v["utility"],
                    "duty_kw": v["duty_kw"],
                    "annual_cost_usd": annual_utility_cost(v["duty_kw"], v["utility"], s.operating_hours, price_override=v.get("price")),
                })
            elif v_type == "Pump":
                rows.append({
                    "tag": v_tag,
                    "duty_type": "Pump (electricity)",
                    "utility": "Electricity",
                    "duty_kw": v["power_kw"],
                    "annual_cost_usd": electricity_cost_annual(v["power_kw"], s.elec_price_per_kwh, s.operating_hours),
                })

        return rows

    # ------------------------------------------------------------------
    def to_dataframes(self, report: CostReport) -> dict[str, pd.DataFrame]:
        """Return dict of DataFrames for export."""
        s = self.settings
        
        # Equipment CAPEX table
        capex_rows = []
        for ec in report.equipment_costs:
            capex_rows.append({
                "Tag": ec.tag,
                "Equipment Type": ec.eq_type,
                "Size Parameter": ec.size_param,
                "Purchase Cost (2001 USD)": round(ec.purchase_cost_2001, 0),
                "Purchase Cost (Current USD)": round(ec.purchase_cost_current, 0),
                "Bare-Module Cost (USD)": round(ec.bare_module_cost, 0),
                "Notes": ec.notes,
                "Warnings": "; ".join(ec.warnings),
            })
        capex_df = pd.DataFrame(capex_rows)

        # Utility OPEX table
        util_rows = []
        for r in report.utility_rows:
            util_rows.append({
                "Tag": r["tag"],
                "Duty Type": r["duty_type"],
                "Utility": r["utility"],
                "Duty (kW)": round(r["duty_kw"], 2),
                "Price ($/GJ)": 0.0,
                "Annual Cost (USD/yr)": round(r["annual_cost_usd"], 0),
            })
        util_df = pd.DataFrame(util_rows)

        # CAPEX Summary
        capex_summary = pd.DataFrame([
            {"Item": "Total Purchase Cost",             "Value (USD)": round(report.total_purchase_cost, 0)},
            {"Item": "Total Bare-Module Cost",          "Value (USD)": round(report.total_bare_module, 0)},
            {"Item": "Contingency",                     "Value (USD)": round(report.contingency, 0)},
            {"Item": "Engineering Fee",                 "Value (USD)": round(report.engineering_fee, 0)},
            {"Item": "ISBL (Inside Battery Limits)",    "Value (USD)": round(report.isbl, 0)},
            {"Item": "OSBL (Outside Battery Limits)",   "Value (USD)": round(report.osbl, 0)},
            {"Item": "Total Fixed Capital Investment",  "Value (USD)": round(report.total_fixed_capital, 0)},
            {"Item": "Startup Cost",                    "Value (USD)": round(report.startup, 0)},
            {"Item": "Working Capital",                 "Value (USD)": round(report.working_capital, 0)},
            {"Item": "Total Project Cost",              "Value (USD)": round(report.total_project_cost, 0)},
        ])

        # OPEX Summary
        opex_summary = pd.DataFrame([
            {"Item": "Raw Materials",                   "Value (USD/yr)": round(self.settings.raw_material_cost * (1.23 if self.settings.opex_method == "Turton (COMd)" else 1.0), 0)},
            {"Item": "Utilities",                       "Value (USD/yr)": round(report.total_utility_opex, 0)},
            {"Item": "Labour",                          "Value (USD/yr)": round(report.labour_opex, 0)},
            {"Item": "Maintenance",                     "Value (USD/yr)": round(report.maintenance_opex, 0)},
            {"Item": "Overhead",                        "Value (USD/yr)": round(report.overhead_opex, 0)},
            {"Item": "Insurance & Tax",                 "Value (USD/yr)": round(report.insurance_opex, 0)},
            {"Item": "Total OPEX",                      "Value (USD/yr)": round(report.total_opex, 0)},
            {"Item": "Annualised CAPEX",                "Value (USD/yr)": round(report.annualised_capex, 0)},
            {"Item": "Total Annual Cost",               "Value (USD/yr)": round(report.total_annual_cost, 0)},
        ])

        return {
            "Equipment_CAPEX": capex_df,
            "Utility_OPEX": util_df,
            "CAPEX_Summary": capex_summary,
            "OPEX_Summary": opex_summary,
        }
