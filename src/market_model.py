"""The same market-sizing logic as the Excel model, in Python, so the workbook can be audited.

tests/test_model.py checks that the workbook's calculated values match this code, catching
any broken formula. Inputs mirror the Inputs sheet.
"""
SALES = {2022: 252555, 2023: 728271, 2024: 948561, 2025: 1149334, 2026: 2550000 * 0.58}
SCENARIOS = {"Bear": {"growth": 0.10, "spend": 2000, "battery_share": 0.40},
             "Base": {"growth": 0.20, "spend": 3000, "battery_share": 0.50},
             "Bull": {"growth": 0.30, "spend": 4000, "battery_share": 0.60}}
SCRAP, BATTERY_PRICE_FY26, BATTERY_DECLINE, REPLACE_AGE = 0.03, 55000, 0.07, 6
INDEPENDENT_SHARE, KARNATAKA_SHARE = 0.25, 0.09
YEARS = range(2022, 2032)
CRORE = 1e7


def run(name: str) -> dict:
    p = SCENARIOS[name]
    sales = dict(SALES)
    for y in range(2027, 2032):
        sales[y] = sales[y - 1] * (1 + p["growth"])
    fleet, out = 0.0, {}
    for y in YEARS:
        fleet = fleet * (1 - SCRAP) + sales[y]
        paying = fleet - sales[y]
        service = paying * p["spend"] / CRORE
        units = sales.get(y - REPLACE_AGE, 0) * p["battery_share"] * (1 - SCRAP) ** REPLACE_AGE
        price = BATTERY_PRICE_FY26 * (1 - BATTERY_DECLINE) ** (y - 2026)
        battery = units * price / CRORE
        out[y] = {"sales": sales[y], "fleet": fleet, "service": service, "battery": battery,
                  "total": service + battery, "independent": (service + battery) * INDEPENDENT_SHARE,
                  "karnataka": (service + battery) * KARNATAKA_SHARE}
    return out


if __name__ == "__main__":
    for n in SCENARIOS:
        r = run(n)
        print(f"{n}: FY26 Rs {r[2026]['total']:,.0f} Cr -> FY31 Rs {r[2031]['total']:,.0f} Cr")
