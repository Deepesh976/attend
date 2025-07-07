from dataclasses import dataclass

@dataclass
class SalaryInputs:
    actual_ctc: float  # J2
    cons_salary: float  # S2
    basic: float
    hra: float
    cca: float
    trp_alw: float
    lop: float
    total_days: int
    days_paid: int
    gross_pay: float  # AE2
    plb: float  # AF2
    cca5: float
    trp_alw6: float
    o_alw1: float
    o_alw17: float
    tds: float
    oth_deds: float

def calculate_consolidated_salary(ctc: float) -> float:
    if ctc < 15285:
        return ctc / 1.1758
    elif 15285 <= ctc < 23758:
        return ctc / 1.1638
    elif 23758 <= ctc < 34299:
        return ctc / 1.1313
    elif ctc >= 34299:
        return (ctc - 1800) / 1.0833
    else:
        raise ValueError("Invalid CTC for consolidated salary")

def calculate_basic(cons_salary: float) -> float:
    if cons_salary > 30000:
        return 15000
    elif 13000 < cons_salary <= 30000:
        return cons_salary * 0.40
    elif cons_salary < 13000:
        return cons_salary * 0.50
    else:
        raise ValueError("Basic salary cannot be determined for exact 13000")

def calculate_hra(basic: float) -> float:
    return basic * 0.40

def calculate_o_alw1(cons_salary, basic, hra, cca, trp_alw):
    return cons_salary - (basic + hra + cca + trp_alw)

def prorate_component(component: float, total_days: int, days_paid: int) -> float:
    return (component / total_days) * days_paid if total_days else 0

def calculate_pf(basic3: float) -> float:
    return basic3 * 0.12

def calculate_esi_employee(cons_salary: float, gross_pay: float, plb: float) -> float:
    if cons_salary > 21000:
        return 0
    return (gross_pay + plb) * 0.0075

def calculate_esi_employer(cons_salary: float, gross_pay: float, plb: float) -> float:
    if cons_salary > 21000:
        return 0
    return (gross_pay + plb) * 0.0325

def calculate_pt(gross_pay: float) -> float:
    if gross_pay > 20000:
        return 200
    elif 15000 < gross_pay <= 20000:
        return 150
    else:
        return 100

def calculate_gpap(cons_salary: float) -> float:
    return ((cons_salary * 12 * 5 * 0.12) / 100) / 12

def calculate_bonus(gross_pay: float) -> float:
    return gross_pay * 0.0833

def calculate_net_pay(gross_pay: float, plb: float, pf: float, esi: float, pt: float, tds: float, gpap: float, oth_deds: float) -> float:
    total_earnings = gross_pay + plb
    total_deductions = pf + esi + pt + tds + gpap + oth_deds
    return total_earnings - total_deductions

def calculate_lpo_ctc(gross_pay: float, pf_employer: float, esi_employer: float, bonus: float) -> float:
    return gross_pay + pf_employer + esi_employer + bonus

# Example usage:
if __name__ == "__main__":
    # Provide mock values for testing
    s = SalaryInputs(
        actual_ctc=128878,
        cons_salary=117321,
        basic=15000,
        hra=6000,
        cca=2000,
        trp_alw=3000,
        lop=2,
        total_days=30,
        days_paid=28,
        gross_pay=29000,
        plb=2000,
        cca5=1900,
        trp_alw6=2800,
        o_alw1=2000,
        o_alw17=1900,
        tds=1000,
        oth_deds=500
    )

    pf = calculate_pf(s.basic)
    esi = calculate_esi_employee(s.cons_salary, s.gross_pay, s.plb)
    pt = calculate_pt(s.gross_pay)
    gpap = calculate_gpap(s.cons_salary)
    bonus = calculate_bonus(s.gross_pay)
    pf_employer = s.basic * 0.12
    esi_employer = calculate_esi_employer(s.cons_salary, s.gross_pay, s.plb)
    net_pay = calculate_net_pay(s.gross_pay, s.plb, pf, esi, pt, s.tds, gpap, s.oth_deds)
    lpo_ctc = calculate_lpo_ctc(s.gross_pay, pf_employer, esi_employer, bonus)

    print(f"PF: {pf}")
    print(f"ESI (Employee): {esi}")
    print(f"PT: {pt}")
    print(f"GPAP: {gpap}")
    print(f"Bonus: {bonus}")
    print(f"Net Pay: {net_pay}")
    print(f"LPO CTC: {lpo_ctc}")
