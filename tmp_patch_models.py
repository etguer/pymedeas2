"""
Patch pymedeas_w and pymedeas_qc model files to support both 14-sector and
16-sector SECTORS_and_HOUSEHOLDS classifications.

Detection: _SECTORS_16 = "Agriculture_Forestry" in _subscript_dict["SECTORS_and_HOUSEHOLDS"]

Changes:
  ccs.py               - _ext_lookup_ccs_tech_share uses loop over subscript dict for 16sec;
                         co2_captured_by_sector_energy_related() Mining special case commented
                         out for 16sec (no equivalent sector)
  process_emissions.py - total_process_emissions() Coke sector usage commented out for 16sec
                         (no equivalent sector); returns 0
  fe_intensity_sectors.py - fe_bottom_up_activation() Transport special case commented out
                         for 16sec (no equivalent sector)
"""

def patch(path, old, new, marker=None):
    """Apply a replacement; marker is a unique string present in new but not in old."""
    text = open(path, encoding='utf-8').read()
    unique = marker or new[:80]
    if unique in text:
        print(f'  Already patched: {path}')
        return
    if old not in text:
        raise AssertionError(f'Pattern not found in {path}:\n{repr(old[:200])}')
    open(path, 'w', encoding='utf-8').write(text.replace(old, new, 1))
    print(f'  Patched: {path}')


# ── ccs.py ────────────────────────────────────────────────────────────────────

SECTORS_16_FLAG = (
    '# [Added for 16-sector compatibility - 2026-06]\n'
    '_SECTORS_16 = "Agriculture_Forestry" in _subscript_dict.get("SECTORS_and_HOUSEHOLDS", [])\n\n'
)

def ccs_old_extlookup(sheet):
    return f'''\
_ext_lookup_ccs_tech_share = ExtLookup(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_hh",
    {{"SECTORS_and_HOUSEHOLDS": ["Households"], "CCS_tech": _subscript_dict["CCS_tech"]}},
    _root,
    {{
        "SECTORS_and_HOUSEHOLDS": _subscript_dict["SECTORS_and_HOUSEHOLDS"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
    "_ext_lookup_ccs_tech_share",
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_agr",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Agriculture"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_mqes",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Mining_quarrying_and_energy_supply"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_fbt",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Food_Beverages_and_Tobacco"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_tex",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Textiles_and_leather_etc"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_coke",
    {{
        "SECTORS_and_HOUSEHOLDS": [
            "Coke_refined_petroleum_nuclear_fuel_and_chemicals_etc"
        ],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_eoete",
    {{
        "SECTORS_and_HOUSEHOLDS": [
            "Electrical_and_optical_equipment_and_Transport_equipment"
        ],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_om",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Other_manufacturing"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_cons",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Construction"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_dist",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Distribution"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_hr",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Hotels_and_restaurant"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_tsc",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Transport_storage_and_communication"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_fi",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Financial_Intermediation"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_re",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Real_estate_renting_and_busine_activitie"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)

_ext_lookup_ccs_tech_share.add(
    r"../climate.xlsx",
    "{sheet}",
    "year_ccs_tech",
    "ccs_tech_share_nms",
    {{
        "SECTORS_and_HOUSEHOLDS": ["Non_Market_Service"],
        "CCS_tech": _subscript_dict["CCS_tech"],
    }},
)'''

def ccs_new_extlookup(sheet):
    old_indented = ccs_old_extlookup(sheet).replace('\n', '\n    ')
    return (
        SECTORS_16_FLAG +
        f'''\
# [Added for 16-sector compatibility - 2026-06]
if _SECTORS_16:
    # 16-sector: load ccs_tech_share_{{SectorName}} ranges for each sector in subscript dict
    _ext_lookup_ccs_tech_share = ExtLookup(
        r"../climate.xlsx",
        "{sheet}",
        "year_ccs_tech",
        "ccs_tech_share_Households",
        {{"SECTORS_and_HOUSEHOLDS": ["Households"], "CCS_tech": _subscript_dict["CCS_tech"]}},
        _root,
        {{
            "SECTORS_and_HOUSEHOLDS": _subscript_dict["SECTORS_and_HOUSEHOLDS"],
            "CCS_tech": _subscript_dict["CCS_tech"],
        }},
        "_ext_lookup_ccs_tech_share",
    )
    for _sector in _subscript_dict["SECTORS_and_HOUSEHOLDS"]:
        if _sector != "Households":
            _ext_lookup_ccs_tech_share.add(
                r"../climate.xlsx",
                "{sheet}",
                "year_ccs_tech",
                f"ccs_tech_share_{{_sector}}",
                {{"SECTORS_and_HOUSEHOLDS": [_sector], "CCS_tech": _subscript_dict["CCS_tech"]}},
            )
else:
    {old_indented}'''
    )

# co2_captured_by_sector_energy_related: Mining special case
# pymedeas_w version (has share_beccs term)
CCS_MINING_OLD_W = '''\
    except_subs.loc[["Mining_quarrying_and_energy_supply"]] = False
    value.values[except_subs.values] = np.minimum(
        co2_policy_captured_sector_ccs() * share_ccs_energy_related(time()),
        co2_emissions_households_and_sectors_fossil_fuels() * (1 + share_beccs(time())),
    ).values[except_subs.values]
    value.loc[["Mining_quarrying_and_energy_supply"]] = float(
        np.minimum(
            float(
                co2_policy_captured_sector_ccs().loc[
                    "Mining_quarrying_and_energy_supply"
                ]
            )
            * float(
                share_ccs_energy_related(time()).loc[
                    "Mining_quarrying_and_energy_supply"
                ]
            ),
            (
                float(
                    co2_emissions_households_and_sectors_fossil_fuels().loc[
                        "Mining_quarrying_and_energy_supply"
                    ]
                )
                + float(co2_emissions_per_fuel().loc["electricity"])
                + float(co2_emissions_per_fuel().loc["heat"])
            )
            * (1 + share_beccs(time())),
        )
    )'''

CCS_MINING_NEW_W = '''\
    # [Added for 16-sector compatibility - 2026-06]
    if _SECTORS_16:
        # 16-sector: Mining_quarrying_and_energy_supply has no equivalent in the 16-sector
        # classification; special CCS treatment for electricity/heat co-emissions commented
        # out. Standard formula applied uniformly to all sectors.
        value.values[except_subs.values] = np.minimum(
            co2_policy_captured_sector_ccs() * share_ccs_energy_related(time()),
            co2_emissions_households_and_sectors_fossil_fuels() * (1 + share_beccs(time())),
        ).values[except_subs.values]
    else:
        except_subs.loc[["Mining_quarrying_and_energy_supply"]] = False
        value.values[except_subs.values] = np.minimum(
            co2_policy_captured_sector_ccs() * share_ccs_energy_related(time()),
            co2_emissions_households_and_sectors_fossil_fuels() * (1 + share_beccs(time())),
        ).values[except_subs.values]
        value.loc[["Mining_quarrying_and_energy_supply"]] = float(
            np.minimum(
                float(
                    co2_policy_captured_sector_ccs().loc[
                        "Mining_quarrying_and_energy_supply"
                    ]
                )
                * float(
                    share_ccs_energy_related(time()).loc[
                        "Mining_quarrying_and_energy_supply"
                    ]
                ),
                (
                    float(
                        co2_emissions_households_and_sectors_fossil_fuels().loc[
                            "Mining_quarrying_and_energy_supply"
                        ]
                    )
                    + float(co2_emissions_per_fuel().loc["electricity"])
                    + float(co2_emissions_per_fuel().loc["heat"])
                )
                * (1 + share_beccs(time())),
            )
        )'''

# pymedeas_qc version (no share_beccs term)
CCS_MINING_OLD_QC = '''\
    except_subs.loc[["Mining_quarrying_and_energy_supply"]] = False
    value.values[except_subs.values] = np.minimum(
        co2_policy_captured_sector_ccs() * share_ccs_energy_related(time()),
        co2_emissions_households_and_sectors_fossil_fuels(),
    ).values[except_subs.values]
    value.loc[["Mining_quarrying_and_energy_supply"]] = float(
        np.minimum(
            float(
                co2_policy_captured_sector_ccs().loc[
                    "Mining_quarrying_and_energy_supply"
                ]
            )
            * float(
                share_ccs_energy_related(time()).loc[
                    "Mining_quarrying_and_energy_supply"
                ]
            ),
            float(
                co2_emissions_households_and_sectors_fossil_fuels().loc[
                    "Mining_quarrying_and_energy_supply"
                ]
            )
            + float(co2_emissions_per_fuel().loc["electricity"])
            + float(co2_emissions_per_fuel().loc["heat"]),
        )
    )'''

CCS_MINING_NEW_QC = '''\
    # [Added for 16-sector compatibility - 2026-06]
    if _SECTORS_16:
        # 16-sector: Mining_quarrying_and_energy_supply has no equivalent in the 16-sector
        # classification; special CCS treatment for electricity/heat co-emissions commented
        # out. Standard formula applied uniformly to all sectors.
        value.values[except_subs.values] = np.minimum(
            co2_policy_captured_sector_ccs() * share_ccs_energy_related(time()),
            co2_emissions_households_and_sectors_fossil_fuels(),
        ).values[except_subs.values]
    else:
        except_subs.loc[["Mining_quarrying_and_energy_supply"]] = False
        value.values[except_subs.values] = np.minimum(
            co2_policy_captured_sector_ccs() * share_ccs_energy_related(time()),
            co2_emissions_households_and_sectors_fossil_fuels(),
        ).values[except_subs.values]
        value.loc[["Mining_quarrying_and_energy_supply"]] = float(
            np.minimum(
                float(
                    co2_policy_captured_sector_ccs().loc[
                        "Mining_quarrying_and_energy_supply"
                    ]
                )
                * float(
                    share_ccs_energy_related(time()).loc[
                        "Mining_quarrying_and_energy_supply"
                    ]
                ),
                float(
                    co2_emissions_households_and_sectors_fossil_fuels().loc[
                        "Mining_quarrying_and_energy_supply"
                    ]
                )
                + float(co2_emissions_per_fuel().loc["electricity"])
                + float(co2_emissions_per_fuel().loc["heat"]),
            )
        )'''


# ── process_emissions.py ──────────────────────────────────────────────────────

PE_FLAG = (
    '# [Added for 16-sector compatibility - 2026-06]\n'
    '_SECTORS_16_PE = "Agriculture_Forestry" in _subscript_dict.get("SECTORS_and_HOUSEHOLDS", [])\n\n\n'
)

def pe_func_body_old(output_fn):
    return f'''\
    return (
        float(
            {output_fn}().loc[
                "Coke_refined_petroleum_nuclear_fuel_and_chemicals_etc"
            ]
        )
        * m_to_t()
        * process_emissions_intensity()
    )'''

def pe_func_body_new(output_fn):
    return f'''\
    # [Added for 16-sector compatibility - 2026-06]
    if _SECTORS_16_PE:
        # 16-sector: Coke_refined_petroleum_nuclear_fuel_and_chemicals_etc has no equivalent
        # in the 16-sector classification; process emissions calculation commented out.
        # Returns 0 for 16-sector runs.
        return 0
    return (
        float(
            {output_fn}().loc[
                "Coke_refined_petroleum_nuclear_fuel_and_chemicals_etc"
            ]
        )
        * m_to_t()
        * process_emissions_intensity()
    )'''


# ── fe_intensity_sectors.py ───────────────────────────────────────────────────

FE_FLAG = (
    '# [Added for 16-sector compatibility - 2026-06]\n'
    '_SECTORS_16_FE = "Agriculture_Forestry" in _subscript_dict.get("SECTORS_and_HOUSEHOLDS", [])\n\n\n'
)

FE_OLD = '''\
    except_subs.loc[["Transport_storage_and_communication"]] = False
    except_subs.loc[["Households"]] = False
    value.values[except_subs.values] = 0
    value.loc[["Transport_storage_and_communication"]] = 0
    value.loc[["Households"]] = 0'''

FE_NEW = '''\
    # [Added for 16-sector compatibility - 2026-06]
    if _SECTORS_16_FE:
        # 16-sector: Transport_storage_and_communication has no equivalent in the 16-sector
        # classification; bottom-up activation for transport commented out.
        # All sectors (including Households) use top-down method.
        except_subs.loc[["Households"]] = False
        value.values[except_subs.values] = 0
        value.loc[["Households"]] = 0
    else:
        except_subs.loc[["Transport_storage_and_communication"]] = False
        except_subs.loc[["Households"]] = False
        value.values[except_subs.values] = 0
        value.loc[["Transport_storage_and_communication"]] = 0
        value.loc[["Households"]] = 0'''


# ── apply patches ─────────────────────────────────────────────────────────────

mining_patches = {
    "pymedeas_w":  (CCS_MINING_OLD_W,  CCS_MINING_NEW_W),
    "pymedeas_qc": (CCS_MINING_OLD_QC, CCS_MINING_NEW_QC),
}

for model, sheet in [("pymedeas_w", "World"), ("pymedeas_qc", "Europe")]:
    base = f"models/{model}/modules_{model}"
    print(f'\n=== {model} ===')

    # ccs.py
    ccs_path = f"{base}/climate/ccs.py"
    patch(ccs_path, ccs_old_extlookup(sheet), ccs_new_extlookup(sheet),
          marker='# [Added for 16-sector compatibility - 2026-06]\nif _SECTORS_16:')
    mining_old, mining_new = mining_patches[model]
    patch(ccs_path, mining_old, mining_new,
          marker='# 16-sector: Mining_quarrying_and_energy_supply has no equivalent')

    # process_emissions.py
    pe_path = f"{base}/climate/process_emissions.py"
    pe_text = open(pe_path, encoding='utf-8').read()
    if '_SECTORS_16_PE' not in pe_text:
        # insert flag before the @component.add decorator for Total_process_emissions
        pe_text = pe_text.replace(
            '@component.add(\n    name="Total_process_emissions"',
            PE_FLAG + '@component.add(\n    name="Total_process_emissions"',
            1,
        )
        open(pe_path, 'w', encoding='utf-8').write(pe_text)
        print(f'  Added flag: {pe_path}')
    pe_output_fn = ("required_total_output_by_sector" if model == "pymedeas_w"
                    else "total_output_required_by_sector")
    patch(pe_path, pe_func_body_old(pe_output_fn), pe_func_body_new(pe_output_fn),
          marker='# 16-sector: Coke_refined_petroleum_nuclear_fuel_and_chemicals_etc has no equivalent')

    # fe_intensity_sectors.py
    fe_path = f"{base}/energy/demand/fe_intensity_sectors.py"
    fe_text = open(fe_path, encoding='utf-8').read()
    if '_SECTORS_16_FE' not in fe_text:
        fe_text = fe_text.replace(
            '@component.add(\n    name="Activate_BOTTOM_UP_method"',
            FE_FLAG + '@component.add(\n    name="Activate_BOTTOM_UP_method"',
            1,
        )
        open(fe_path, 'w', encoding='utf-8').write(fe_text)
        print(f'  Added flag: {fe_path}')
    patch(fe_path, FE_OLD, FE_NEW,
          marker='# 16-sector: Transport_storage_and_communication has no equivalent')

print('\nAll patches applied.')
