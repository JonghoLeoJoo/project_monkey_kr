"""
Korean Financial Model Builder
================================
Fetches financial statements from DART (dart.fss.or.kr) and market data from
KRX via pykrx.  Generates an Excel workbook with:

  - Sheet 1: Financial Statements (IS / BS / CF, 4 historical years + LTM)
  - Sheet 2: WACC (Korean market inputs)
  - Sheet 3: DCF Model (5-year projection + terminal value)
  - Sheet 4: P/E Comparable Companies
  - Sheet 5: Data Validation

Usage:
    python main.py 005930              # Samsung Electronics (by stock code)
    python main.py "삼성전자"           # by Korean name
    python main.py --bulk-test         # Run validation on 20 largest companies

Requires:
    DART_API_KEY in .env or as environment variable.
    Register at https://opendart.fss.or.kr/
"""

import sys
import os
import json
import time

sys.stdout.reconfigure(encoding='utf-8')

from dart_fetcher import (
    search_company,
    get_fiscal_years,
    extract_financial_data,
    compute_ltm,
    get_stock_price,
    get_historical_prices,
    get_current_price,
    get_market_cap,
    get_kr_bond_yield,
    get_kr_erp,
    get_industry_peers,
    get_pe_comps_data,
)
from excel_builder import create_excel


# ~20 largest KOSPI companies by market cap (diverse sectors)
BULK_TEST_CODES = [
    '005930',  # 삼성전자 (Samsung Electronics)
    '000660',  # SK하이닉스 (SK Hynix)
    '005380',  # 현대자동차 (Hyundai Motor)
    '035420',  # NAVER
    '005490',  # POSCO홀딩스
    '051910',  # LG화학 (LG Chem)
    '006400',  # 삼성SDI (Samsung SDI)
    '035720',  # 카카오 (Kakao)
    '000270',  # 기아 (Kia)
    '068270',  # 셀트리온 (Celltrion)
    '105560',  # KB금융 (KB Financial)
    '055550',  # 신한지주 (Shinhan Financial)
    '012330',  # 현대모비스 (Hyundai Mobis)
    '028260',  # 삼성물산 (Samsung C&T)
    '066570',  # LG전자 (LG Electronics)
    '003550',  # LG (LG Corp)
    '034730',  # SK (SK Inc)
    '032830',  # 삼성생명 (Samsung Life)
    '096770',  # SK이노베이션 (SK Innovation)
    '003670',  # 포스코퓨처엠 (POSCO Future M)
]


def _safe_name(name: str) -> str:
    """Create a filesystem-safe version of a Korean company name."""
    import re
    # Keep Korean chars, alphanumeric, space, hyphen
    safe = re.sub(r'[^\w\s\-]', '', name, flags=re.UNICODE)
    safe = safe.strip().replace(' ', '_')
    return safe or 'company'


def build_model(company_query: str, skip_prices: bool = False,
                auto_select: bool = False, price_date: str = None):
    """Run the full pipeline for one company.

    Returns ``(company_info, validation_results)`` or *None* on failure.
    """
    # Step 1: Search company on DART
    company_info = search_company(company_query)
    if not company_info:
        print(f"  [SKIP] Could not find '{company_query}' on DART.")
        return None

    corp_code = company_info['corp_code']
    corp_name = company_info['corp_name']
    stock_code = company_info['stock_code']
    market = company_info['market']

    print(f"  [OK] {corp_name} ({stock_code}), Market: {market}")

    # Step 2: Determine fiscal years
    years = get_fiscal_years(corp_code, n_years=4)
    if not years:
        print(f"  [SKIP] No annual data found for {corp_name}.")
        return None

    years_display = ', '.join(f'FY{y}' for y in reversed(years))
    print(f"  [OK] Found data for: {years_display}")

    # Step 3: Extract financial data from DART
    financial_data = extract_financial_data(corp_code, years, stock_code=stock_code)

    # Step 4: Compute LTM if quarterly data available
    latest_year = years[0]
    ltm_year = compute_ltm(corp_code, latest_year, financial_data)
    if ltm_year:
        print(f"  [OK] LTM data added for year {ltm_year}")
    years = financial_data['years']  # May have been updated with LTM

    # Step 5: Stock prices, market cap, WACC inputs
    if not skip_prices:
        closing_prices = get_historical_prices(stock_code, years)
    else:
        closing_prices = {yr: None for yr in years}

    # Market cap from prices * shares
    market_cap_data = {}
    shares = financial_data['income_statement']['shares_diluted']
    for yr in years:
        price = closing_prices.get(yr)
        shr = shares.get(yr)
        if price is not None and shr is not None and shr > 0:
            market_cap_data[yr] = price * shr
        else:
            market_cap_data[yr] = None

    financial_data['market_cap'] = market_cap_data
    financial_data['stock_prices'] = closing_prices

    # WACC components
    if not skip_prices:
        risk_free = get_kr_bond_yield()
        erp = get_kr_erp()
    else:
        risk_free = 0.035
        erp = 0.065

    # Comparable companies
    comp_data = []
    industry_name = 'Unknown'
    if not skip_prices:
        comp_data, industry_name = get_industry_peers(
            stock_code, auto_select=auto_select,
        )
        if comp_data:
            print(f"  Comparable companies ({industry_name}): {len(comp_data)} peers")

    # Price date for WACC
    if not auto_select and not skip_prices and price_date is None:
        print()
        price_date = input(
            "  Share price date for WACC (YYYY-MM-DD, or Enter for latest): "
        ).strip() or None

    if not skip_prices:
        current_price_data = get_current_price(stock_code, price_date)
    else:
        current_price_data = {'price': None, 'date': None}

    # Shares breakdown -- use most recent year that has actual shares data
    # (LTM year may not have shares populated)
    latest = years[0]
    inc = financial_data['income_statement']
    bs = financial_data['balance_sheet']

    basic_sh = 0
    diluted_sh = 0
    for yr in years:
        b = inc['shares_basic'].get(yr)
        if b and b > 0:
            basic_sh = b
            diluted_sh = inc['shares_diluted'].get(yr) or b
            break

    # Implied cost of debt: interest expense / average total debt
    # Uses average of current and prior period debt outstanding.
    # Korean IFRS 이자비용 often includes lease interest and other items,
    # so cap at 12% to avoid unreasonable values.
    def _total_debt(yr):
        return ((bs.get('st_debt', {}).get(yr) or 0)
                + (bs.get('lt_debt', {}).get(yr) or 0)
                + (bs.get('current_lt_debt', {}).get(yr) or 0))

    int_exp = abs(inc['interest_expense'].get(latest) or 0)
    debt_cur = _total_debt(latest)
    prior = years[1] if len(years) > 1 else None
    debt_pri = _total_debt(prior) if prior else 0
    avg_debt = (debt_cur + debt_pri) / 2 if (debt_cur + debt_pri) > 0 else 0
    if avg_debt > 0:
        raw_cod = int_exp / avg_debt
        implied_cod = round(min(raw_cod, 0.12), 4)  # Cap at 12%
    else:
        implied_cod = 0.05  # Default 5%

    financial_data['wacc_inputs'] = {
        'current_price': current_price_data,
        'treasury_yield': risk_free,  # Korean 10Y bond
        'kroll_erp': erp,             # Korean ERP
        'comparables': comp_data,
        'shares_breakdown': {
            'basic': basic_sh,
            'rsus': max(diluted_sh - basic_sh, 0),
            'options': 0,
            'conv_debt': 0,
            'conv_pref': 0,
        },
        'implied_cod': implied_cod,
    }

    # Step 6: P/E comparable companies
    pe_comps = []
    if not skip_prices and comp_data:
        print("  Fetching P/E comparable data...")
        pe_comps = get_pe_comps_data(comp_data, stock_code)
        valid_pe = sum(1 for p in pe_comps if p.get('trailing_pe'))
        print(f"  [OK] P/E data for {valid_pe}/{len(pe_comps)} peers")

    financial_data['pe_comps'] = pe_comps

    # Step 7: Build Excel workbook
    safe = _safe_name(corp_name)
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'models')
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"{safe}_Financial_Model.xlsx")

    # JSON backup
    def _stringify_keys(obj):
        if isinstance(obj, dict):
            return {str(k): _stringify_keys(v) for k, v in obj.items()}
        return obj

    dart_dir = os.path.join(output_dir, 'dart')
    os.makedirs(dart_dir, exist_ok=True)
    dart_file = os.path.join(dart_dir, f"{safe}_dart.json")
    with open(dart_file, 'w', encoding='utf-8') as f:
        json.dump({
            'company': company_info,
            'fiscal_years': years,
            'financial_data': _stringify_keys(financial_data),
        }, f, indent=2, ensure_ascii=False)

    validation_results = create_excel(company_info, financial_data, output_file)

    return company_info, validation_results


def bulk_test():
    """Run the model on ~20 largest KOSPI companies and print a validation report."""
    print()
    print("=" * 70)
    print("  BULK VALIDATION TEST -- 20 Largest KOSPI Companies")
    print("=" * 70)
    print()

    results = []
    for i, code in enumerate(BULK_TEST_CODES, 1):
        print(f"\n{'-' * 70}")
        print(f"  [{i}/{len(BULK_TEST_CODES)}]  Processing: {code}")
        print(f"{'-' * 70}")

        try:
            result = build_model(code, skip_prices=False, auto_select=True)
            if result is None:
                results.append({
                    'code': code, 'status': 'SKIP',
                    'reason': 'not found / no data',
                })
            else:
                company_info, val = result
                results.append({
                    'code': code,
                    'name': company_info['corp_name'],
                    'status': 'OK',
                    'total': val['total'],
                    'passed': val['passed'],
                    'failed': val['failed'],
                    'checks': val.get('checks', []),
                })
        except Exception as e:
            results.append({'code': code, 'status': 'ERROR', 'reason': str(e)})
            import traceback
            traceback.print_exc()

        time.sleep(0.3)  # Rate limiting

    # -- Consolidated Report -----------------------------------------------
    print()
    print()
    print("=" * 70)
    print("  BULK VALIDATION REPORT")
    print("=" * 70)
    print()

    grand_total = 0
    grand_passed = 0
    grand_failed = 0

    for r in results:
        if r['status'] == 'OK':
            grand_total += r['total']
            grand_passed += r['passed']
            grand_failed += r['failed']
            print(f"  {r['code']:>6s}:  {r['passed']}/{r['total']} checks passed  "
                  f"- {r['name']}")
        elif r['status'] == 'SKIP':
            print(f"  {r['code']:>6s}:  SKIPPED - {r['reason']}")
        else:
            print(f"  {r['code']:>6s}:  ERROR   - {r['reason']}")

    print()
    print(f"  {'-' * 50}")
    if grand_total > 0:
        pct = grand_passed / grand_total * 100
        print(f"  TOTAL:  {grand_passed}/{grand_total} checks passed  ({pct:.1f}%)")
        print(f"  FAILED: {grand_failed}")
    else:
        print("  No checks were run.")
    print()


def main():
    print()
    print("=" * 60)
    print("  Financial Model Builder  -  DART / KRX Edition")
    print("=" * 60)
    print()

    # Check for --bulk-test flag
    if '--bulk-test' in sys.argv:
        bulk_test()
        return

    # Allow stock code or company name as command-line argument
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if args:
        company_query = ' '.join(args).strip()
        print(f"  Company: {company_query}")
    else:
        company_query = input(
            "  Enter stock code (e.g. 005930) or company name: "
        ).strip()

    if not company_query:
        print("  No company provided. Exiting.")
        sys.exit(1)

    # Step 1: Search DART
    print()
    print("[1/7] Searching DART...")
    company_info = search_company(company_query)
    if not company_info:
        print("\n  Could not find the company. Try the 6-digit stock code")
        print("  (e.g. 005930) or the Korean company name (e.g. 삼성전자).")
        sys.exit(1)

    corp_code = company_info['corp_code']
    corp_name = company_info['corp_name']
    stock_code = company_info['stock_code']
    market = company_info['market']

    print(f"\n  [OK] Company    : {corp_name}")
    print(f"       Stock Code : {stock_code}")
    print(f"       Market     : {market}")

    # Step 2: Determine fiscal years
    print()
    print("[2/7] Finding available fiscal years...")
    years = get_fiscal_years(corp_code, n_years=4)
    if not years:
        print("\n  Could not find annual filing data for this company.")
        sys.exit(1)

    years_display = ', '.join(f'FY{y}' for y in reversed(years))
    print(f"  [OK] Found data for: {years_display}")

    # Step 3: Extract financial data
    print()
    print("[3/7] Extracting financial statements from DART...")
    financial_data = extract_financial_data(corp_code, years, stock_code=stock_code)

    # Quick sanity summary
    inc = financial_data['income_statement']
    latest_yr = years[0]
    rev = inc['revenue'].get(latest_yr)
    ni = inc['net_income'].get(latest_yr)
    scale = 1e8  # Display in 억원 (100 million KRW)
    print(f"  [OK] Latest year FY{latest_yr}:")
    if rev:
        print(f"         Revenue    : {rev/scale:,.0f} 억원")
    if ni:
        print(f"         Net Income : {ni/scale:,.0f} 억원")

    # Step 4: Compute LTM
    print()
    print("[4/7] Checking for quarterly data (LTM)...")
    ltm_year = compute_ltm(corp_code, latest_yr, financial_data)
    if ltm_year:
        years = financial_data['years']
        ltm_rev = inc['revenue'].get(ltm_year)
        ltm_ni = inc['net_income'].get(ltm_year)
        print(f"  [OK] LTM {ltm_year}:")
        if ltm_rev:
            print(f"         Revenue    : {ltm_rev/scale:,.0f} 억원")
        if ltm_ni:
            print(f"         Net Income : {ltm_ni/scale:,.0f} 억원")
    else:
        print("  [--] No quarterly data available; using annual figures only.")

    # Step 5: Stock prices & WACC inputs
    print()
    print("[5/7] Fetching stock prices and WACC inputs...")
    closing_prices = get_historical_prices(stock_code, years)

    market_cap_data = {}
    shares = financial_data['income_statement']['shares_diluted']
    for yr in years:
        price = closing_prices.get(yr)
        shr = shares.get(yr)
        if price is not None and shr is not None and shr > 0:
            market_cap_data[yr] = price * shr
            print(f"  [OK] FY{yr} Market Cap: {price * shr / scale:,.0f} 억원 "
                  f"(₩{price:,.0f} x {shr/1e6:,.1f}M shares)")
        else:
            market_cap_data[yr] = None
            print(f"  [--] FY{yr} Market Cap: unavailable")

    financial_data['market_cap'] = market_cap_data
    financial_data['stock_prices'] = closing_prices

    risk_free = get_kr_bond_yield()
    erp = get_kr_erp()

    print()
    price_date = input(
        "  Share price date for WACC (YYYY-MM-DD, or Enter for latest): "
    ).strip() or None

    print("  Finding comparable companies...")
    comp_data, industry_name = get_industry_peers(stock_code)
    if comp_data:
        print(f"  [OK] Comparable companies ({industry_name}): {len(comp_data)} peers")
        for cd in comp_data[:5]:
            print(f"    {cd['name']} ({cd['ticker']}): beta={cd['beta']:.2f}")
    else:
        print("  [--] No comparable companies found")

    current_price_data = get_current_price(stock_code, price_date)
    if current_price_data['price']:
        print(f"  [OK] Share Price: ₩{current_price_data['price']:,.0f} "
              f"({current_price_data['date']})")
    else:
        print("  [--] Share price unavailable")

    latest = years[0]
    bs = financial_data['balance_sheet']

    # Use most recent year with actual shares data (LTM year may lack it)
    basic_sh = 0
    diluted_sh = 0
    for yr in years:
        b = inc['shares_basic'].get(yr)
        if b and b > 0:
            basic_sh = b
            diluted_sh = inc['shares_diluted'].get(yr) or b
            break

    # Implied cost of debt: interest expense / average total debt
    def _total_debt(yr):
        return ((bs.get('st_debt', {}).get(yr) or 0)
                + (bs.get('lt_debt', {}).get(yr) or 0)
                + (bs.get('current_lt_debt', {}).get(yr) or 0))

    int_exp = abs(inc['interest_expense'].get(latest) or 0)
    debt_cur = _total_debt(latest)
    prior = years[1] if len(years) > 1 else None
    debt_pri = _total_debt(prior) if prior else 0
    avg_debt = (debt_cur + debt_pri) / 2 if (debt_cur + debt_pri) > 0 else 0
    if avg_debt > 0:
        raw_cod = int_exp / avg_debt
        implied_cod = round(min(raw_cod, 0.12), 4)
    else:
        implied_cod = 0.05

    financial_data['wacc_inputs'] = {
        'current_price': current_price_data,
        'treasury_yield': risk_free,
        'kroll_erp': erp,
        'comparables': comp_data,
        'shares_breakdown': {
            'basic': basic_sh,
            'rsus': max(diluted_sh - basic_sh, 0),
            'options': 0,
            'conv_debt': 0,
            'conv_pref': 0,
        },
        'implied_cod': implied_cod,
    }

    # Step 6: P/E comparables
    print()
    print("[6/7] Fetching P/E comparable company data...")
    pe_comps = []
    if comp_data:
        pe_comps = get_pe_comps_data(comp_data, stock_code)
        valid_pe = sum(1 for p in pe_comps if p.get('trailing_pe'))
        print(f"  [OK] P/E data for {valid_pe}/{len(pe_comps)} comparable companies")
        for p in pe_comps[:5]:
            if p.get('trailing_pe'):
                print(f"    {p['name']}: Trailing P/E = {p['trailing_pe']:.1f}x")
    else:
        print("  [--] No peers available for P/E analysis")

    financial_data['pe_comps'] = pe_comps

    # Step 7: Build Excel workbook
    print()
    print("[7/7] Building Excel financial model...")

    safe = _safe_name(corp_name)
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'models')
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"{safe}_Financial_Model.xlsx")

    # JSON backup
    def _stringify_keys(obj):
        if isinstance(obj, dict):
            return {str(k): _stringify_keys(v) for k, v in obj.items()}
        return obj

    dart_dir = os.path.join(output_dir, 'dart')
    os.makedirs(dart_dir, exist_ok=True)
    dart_file = os.path.join(dart_dir, f"{safe}_dart.json")
    with open(dart_file, 'w', encoding='utf-8') as f:
        json.dump({
            'company': company_info,
            'fiscal_years': years,
            'financial_data': _stringify_keys(financial_data),
        }, f, indent=2, ensure_ascii=False)
    print(f"  [OK] DART data saved to: models/dart/{safe}_dart.json")

    validation_results = create_excel(company_info, financial_data, output_file)

    abs_path = os.path.abspath(output_file)
    print()
    print("=" * 60)
    print("  Done!")
    print(f"  File saved to: {abs_path}")
    print("=" * 60)
    print()
    print("  Excel workbook contents:")
    print("  - Sheet 'Financial Statements': 3-statement model (historical + LTM)")
    print("  - Sheet 'WACC'               : Weighted avg cost of capital")
    print("  - Sheet 'DCF Model'           : 5-year DCF valuation")
    print("  - Sheet 'PE Comps'            : P/E comparable companies")
    print("  - Sheet 'Data Validation'     : Cross-checks (PASS/FAIL)")
    print()
    if validation_results:
        p = validation_results['passed']
        t = validation_results['total']
        f_count = validation_results['failed']
        print(f"  Validation: {p}/{t} checks passed", end='')
        if f_count > 0:
            print(f"  ({f_count} failed - see Data Validation sheet)")
        else:
            print("  (all passed)")
    print()
    print("  Tip: Yellow cells in the WACC and DCF sheets are editable inputs.")
    print("       Change WACC assumptions and growth rates to run scenarios.")
    print()


if __name__ == '__main__':
    main()
