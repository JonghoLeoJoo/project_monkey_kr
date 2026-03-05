"""
DART Financial Data Fetcher
============================
Fetches financial statements from Korea's DART (dart.fss.or.kr) and market
data from KRX via pykrx.  Produces the same ``financial_data`` dictionary
structure used by the Excel builder so the workbook generator is largely
data-source-agnostic.

Dependencies:
    pip install opendartreader pykrx python-dotenv yfinance requests
"""

from __future__ import annotations

import io
import os
import re
import sys
import time
import warnings
from contextlib import contextmanager
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple


@contextmanager
def _suppress_stdout():
    """Temporarily redirect stdout/stderr to suppress noisy library debug prints."""
    old_out, old_err = sys.stdout, sys.stderr
    sys.stdout = io.StringIO()
    sys.stderr = io.StringIO()
    try:
        yield
    finally:
        sys.stdout = old_out
        sys.stderr = old_err

import requests

try:
    import OpenDartReader
    _HAS_DART = True
except ImportError:
    _HAS_DART = False

try:
    from pykrx import stock as krx_stock
    _HAS_PYKRX = True
except ImportError:
    _HAS_PYKRX = False

try:
    import yfinance as yf
    _HAS_YFINANCE = True
except ImportError:
    _HAS_YFINANCE = False

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ---------------------------------------------------------------------------
# Module-level DART reader (initialised lazily)
# ---------------------------------------------------------------------------
_dart: Optional[OpenDartReader.OpenDartReader] = None


def _get_dart() -> OpenDartReader.OpenDartReader:
    global _dart
    if _dart is None:
        api_key = os.environ.get('DART_API_KEY', '')
        if not api_key:
            raise RuntimeError(
                "DART_API_KEY not set. Place it in .env or export it:\n"
                "  export DART_API_KEY=<your 40-char key>\n"
                "Register at https://opendart.fss.or.kr/"
            )
        _dart = OpenDartReader(api_key)
    return _dart


# =========================================================================
# 1.  COMPANY SEARCH
# =========================================================================

def search_company(query: str) -> Optional[Dict]:
    """Look up a company by 6-digit stock code or Korean name.

    Returns ``{corp_code, corp_name, stock_code, market}`` or *None*.
    """
    dart = _get_dart()
    codes = dart.corp_codes

    # Try exact stock code match first
    query = query.strip()
    if re.fullmatch(r'\d{6}', query):
        match = codes[codes['stock_code'] == query]
        if not match.empty:
            row = match.iloc[0]
            market = _detect_market(query)
            return {
                'corp_code': row['corp_code'],
                'corp_name': row['corp_name'],
                'stock_code': query,
                'market': market,
            }

    # Fuzzy name search
    match = codes[codes['corp_name'].str.contains(query, na=False)]
    # Prefer listed companies (stock_code not empty / NaN)
    listed = match[match['stock_code'].notna() & (match['stock_code'] != '')]
    if listed.empty and match.empty:
        return None
    target = listed if not listed.empty else match
    row = target.iloc[0]
    stock_code = str(row.get('stock_code', ''))
    market = _detect_market(stock_code) if stock_code else 'UNKNOWN'
    return {
        'corp_code': row['corp_code'],
        'corp_name': row['corp_name'],
        'stock_code': stock_code,
        'market': market,
    }


def _detect_market(stock_code: str) -> str:
    """Determine whether a stock is on KOSPI or KOSDAQ.

    Uses DART ``company()`` API: ``corp_cls`` Y = KOSPI, K = KOSDAQ.
    """
    try:
        dart = _get_dart()
        codes = dart.corp_codes
        match = codes[codes['stock_code'] == stock_code]
        if not match.empty:
            corp_code = match.iloc[0]['corp_code']
            info = dart.company(corp_code)
            if isinstance(info, dict):
                cls = info.get('corp_cls', '')
                if cls == 'Y':
                    return 'KOSPI'
                elif cls == 'K':
                    return 'KOSDAQ'
    except Exception:
        pass
    return 'UNKNOWN'


# =========================================================================
# 2.  FINANCIAL STATEMENT EXTRACTION
# =========================================================================

# --- Korean account-name fallback chains -----------------------------------
# Each key maps to a list of Korean account names (tried in order).
# The first match wins, similar to XBRL concept fallback chains.

IS_MAP: Dict[str, List[str]] = {
    'revenue':          ['매출액', '영업수익', '수익(매출액)', '매출', '이자수익'],
    'cogs':             ['매출원가'],
    'gross_profit':     ['매출총이익'],
    'sga_expense':      ['판매비와관리비', '판매비와 관리비'],
    'rd_expense':       ['연구개발비', '경상연구개발비', '연구비'],
    'operating_income': ['영업이익', '영업이익(손실)'],
    'interest_expense': ['금융비용', '이자비용', '금융원가'],
    'interest_income':  ['금융수익', '이자수익'],
    'other_income':     ['기타수익', '기타이익'],
    'other_expense':    ['기타비용', '기타손실'],
    'pretax_income':    ['법인세비용차감전순이익', '법인세비용차감전순이익(손실)',
                         '법인세비용차감전순이익(손실)', '법인세차감전순이익'],
    'tax_expense':      ['법인세비용', '법인세비용(수익)'],
    'net_income':       ['당기순이익', '당기순이익(손실)', '지배기업 소유지분',
                         '지배기업소유주지분'],
    'net_income_total': ['당기순이익', '당기순이익(손실)'],
    'minority_interest': ['비지배지분', '비지배주주지분'],
    'eps_basic':        ['기본주당이익', '기본주당순이익', '기본주당이익(손실)'],
    'eps_diluted':      ['희석주당이익', '희석주당순이익', '희석주당이익(손실)'],
    'equity_method':    ['지분법이익', '지분법손익'],
}

BS_MAP: Dict[str, List[str]] = {
    'total_assets':     ['자산총계'],
    'total_current_a':  ['유동자산'],
    'cash':             ['현금및현금성자산'],
    'st_investments':   ['단기금융상품', '단기상각후원가금융자산',
                         '단기당기손익-공정가치금융자산'],
    'accounts_rec':     ['매출채권', '매출채권 및 기타유동채권'],
    'inventory':        ['재고자산'],
    'other_current_a':  ['기타유동자산'],
    'total_noncurrent_a': ['비유동자산'],
    'ppe_net':          ['유형자산'],
    'goodwill':         ['영업권'],
    'intangibles':      ['무형자산'],
    'lt_investments':   ['장기금융상품', '기타포괄손익-공정가치금융자산',
                         '관계기업 및 공동기업 투자'],
    'other_noncurrent_a': ['기타비유동자산'],
    'deferred_tax_a':   ['이연법인세자산'],

    'total_liabilities': ['부채총계'],
    'total_current_l':  ['유동부채'],
    'accounts_pay':     ['매입채무', '매입채무 및 기타유동채무'],
    'st_debt':          ['단기차입금'],
    'current_lt_debt':  ['유동성장기부채'],
    'accrued_liab':     ['미지급비용', '미지급금'],
    'other_current_l':  ['기타유동부채', '충당부채'],
    'deferred_rev_cur': ['선수금', '예수금'],
    'total_noncurrent_l': ['비유동부채'],
    'lt_debt':          ['장기차입금', '사채'],
    'deferred_tax_l':   ['이연법인세부채'],
    'other_noncurrent_l': ['기타비유동부채', '장기충당부채', '순확정급여부채',
                           '장기미지급금'],

    'total_equity':     ['자본총계'],
    'controlling_eq':   ['지배기업 소유주지분'],
    'common_stock':     ['자본금', '보통주자본금'],
    'apic':             ['주식발행초과금'],
    'retained_earnings': ['이익잉여금'],
    'treasury_stock':   ['기타자본항목'],
    'minority_equity':  ['비지배지분'],
}

CF_MAP: Dict[str, List[str]] = {
    'operating_cf':     ['영업활동현금흐름', '영업활동으로인한현금흐름',
                         '영업활동으로 인한 현금흐름'],
    'cf_from_ops':      ['영업에서 창출된 현금흐름', '영업에서 창출된 현금'],
    'net_income_cf':    ['당기순이익'],
    'adjustments':      ['조정'],
    'wc_changes':       ['영업활동으로 인한 자산부채의 변동'],
    'interest_paid':    ['이자의 지급'],
    'interest_received': ['이자의 수취'],
    'tax_paid':         ['법인세 납부액', '법인세의 납부'],
    'dividends_received': ['배당금 수입'],
    'investing_cf':     ['투자활동현금흐름', '투자활동으로인한현금흐름',
                         '투자활동으로 인한 현금흐름'],
    'capex':            ['유형자산의 취득'],
    'intangible_acq':   ['무형자산의 취득'],
    'ppe_disposal':     ['유형자산의 처분'],
    'acquisitions':     ['사업결합으로 인한 현금유출액'],
    'financing_cf':     ['재무활동현금흐름', '재무활동으로인한현금흐름',
                         '재무활동으로 인한 현금흐름'],
    'dividends_paid':   ['배당금의지급', '배당금의 지급', '배당금지급'],
    'repurchases':      ['자기주식의 취득'],
    'debt_issuance':    ['장기차입금의 차입', '단기차입금의 순증가(감소)'],
    'debt_repay':       ['사채 및 장기차입금의 상환'],
    'fx_effect':        ['외화환산으로 인한 현금의 변동',
                         '현금및현금성자산의 환율변동효과'],
    'net_cash_change':  ['현금및현금성자산의 증가(감소)', '현금및현금성자산의 증가'],
    'beginning_cash':   ['기초현금및현금성자산', '기초의 현금및현금성자산'],
    'ending_cash':      ['기말현금및현금성자산', '기말의 현금및현금성자산'],
}


def _parse_amount(val) -> Optional[float]:
    """Parse DART amount string into a float (raw KRW)."""
    if val is None:
        return None
    s = str(val).strip()
    if s in ('', 'None', '-', 'nan'):
        return None
    # Remove commas
    s = s.replace(',', '')
    # Handle parenthesised negatives: (123) -> -123
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return None


# PPE sub-item prefixes for capex fallback
_PPE_ACQ_CATEGORIES = (
    '토지', '건물', '구축물', '기계장치', '차량운반구',
    '사무용비품', '비품', '기타유형자산', '건설중인자산',
    '시설장치', '공기구비품', '집기비품',
)


def _sum_ppe_acquisitions(df, col: str = 'thstrm_amount') -> Optional[float]:
    """Sum individual PPE acquisition items from CF rows.

    Some companies break down capex into individual asset categories
    (건물의 취득, 사무용비품의 취득, …) instead of a single 유형자산의 취득.
    Returns negative total (capex convention) or None if nothing found.
    """
    if df is None or df.empty:
        return None
    cf_rows = df[df['sj_div'] == 'CF'] if 'sj_div' in df.columns else df
    total = 0.0
    found = False
    for _, row in cf_rows.iterrows():
        acct = str(row.get('account_nm', ''))
        if any(acct.startswith(cat) for cat in _PPE_ACQ_CATEGORIES) and '취득' in acct:
            v = _parse_amount(row.get(col))
            if v is not None:
                total += abs(v)
                found = True
    return -total if found else None


def _extract_from_df(df, account_names: List[str],
                     col: str = 'thstrm_amount',
                     sj_div: Optional[str] = None) -> Optional[float]:
    """Find the first matching account name in the DataFrame and return its value.

    When *sj_div* is ``'IS'``, also searches ``'CIS'`` (Comprehensive Income
    Statement) as a fallback — some companies only file under CIS.

    Matching strategy:
      1. Exact match on ``account_nm``
      2. Normalised match — strips Roman numeral prefixes (Ⅰ.~Ⅶ.), leading
         digits/dots, and whitespace so that e.g.
         ``'Ⅰ. 영업활동으로 인한 현금흐름'`` matches ``'영업활동으로 인한 현금흐름'``.
    """
    if df is None or df.empty:
        return None
    # Build list of sj_div values to try
    divs_to_try = [sj_div] if sj_div else [None]
    if sj_div == 'IS':
        divs_to_try.append('CIS')

    def _normalise(s: str) -> str:
        """Strip Roman numeral / numeric prefixes and collapse spaces."""
        s = re.sub(r'^[ⅠⅡⅢⅣⅤⅥⅦ0-9]+\.\s*', '', s)
        return s.replace(' ', '')

    for div in divs_to_try:
        subset = df
        if div:
            subset = subset[subset['sj_div'] == div]
        # Pass 1: exact match
        for name in account_names:
            rows = subset[subset['account_nm'] == name]
            if not rows.empty:
                return _parse_amount(rows.iloc[0][col])
        # Pass 2: normalised match (handles prefixes / spacing differences)
        for name in account_names:
            norm_name = _normalise(name)
            for _, row in subset.iterrows():
                if _normalise(str(row['account_nm'])) == norm_name:
                    return _parse_amount(row[col])
    return None


def _build_stmt_dict(df, mapping: Dict[str, List[str]],
                     sj_div: str,
                     year_cols: List[Tuple[int, str]]) -> Dict[str, Dict[int, Optional[float]]]:
    """Build {metric: {year: value}} from a DataFrame using the mapping.

    *year_cols* is a list of ``(year, column_name)`` tuples, e.g.
    ``[(2024, 'thstrm_amount'), (2023, 'frmtrm_amount'), ...]``.
    """
    result: Dict[str, Dict[int, Optional[float]]] = {}
    for key, names in mapping.items():
        vals: Dict[int, Optional[float]] = {}
        for yr, col in year_cols:
            vals[yr] = _extract_from_df(df, names, col=col, sj_div=sj_div)
        result[key] = vals
    return result


def get_fiscal_years(corp_code: str, n_years: int = 4) -> List[int]:
    """Determine the *n_years* most recent fiscal years with DART data.

    Returns years in descending order (newest first).

    For recently-listed companies that only have one annual filing, the
    prior-year data embedded in that filing (frmtrm / bfefrmtrm columns)
    is also included so that we can compute growth rates.
    """
    dart = _get_dart()
    current_year = datetime.today().year
    years_found: List[int] = []
    for y in range(current_year, current_year - 6, -1):
        try:
            with _suppress_stdout():
                fs = dart.finstate(corp_code, y, reprt_code='11011')
            if fs is not None and not fs.empty:
                years_found.append(y)
                if len(years_found) >= n_years:
                    break
        except Exception:
            pass
        time.sleep(0.1)

    # If we have fewer years than requested, try to recover prior-year data
    # from the earliest report's frmtrm / bfefrmtrm columns.
    if 0 < len(years_found) < n_years:
        earliest = min(years_found)
        try:
            with _suppress_stdout():
                df = dart.finstate_all(corp_code, earliest, reprt_code='11011')
            if df is not None and not df.empty:
                for col, offset in [('frmtrm_amount', 1), ('bfefrmtrm_amount', 2)]:
                    prior_yr = earliest - offset
                    if prior_yr not in years_found and col in df.columns:
                        # Check if there's actual data in this column
                        non_null = df[col].dropna()
                        non_null = non_null[non_null != '']
                        if not non_null.empty:
                            years_found.append(prior_yr)
                            if len(years_found) >= n_years:
                                break
        except Exception:
            pass

    return sorted(years_found, reverse=True)


def extract_financial_data(corp_code: str, years: List[int],
                           stock_code: str = '') -> Dict:
    """Fetch annual reports from DART and build the financial_data dict.

    Uses ``finstate_all`` which returns consolidated (CFS) data and gives
    current, prior, and two-periods-ago amounts in a single call.
    """
    dart = _get_dart()
    years = sorted(years, reverse=True)  # newest first

    # We fetch annual reports to cover all years.  Each report gives up to
    # 3 years of data (thstrm, frmtrm, bfefrmtrm).
    all_dfs: Dict[int, object] = {}
    for y in years:
        try:
            with _suppress_stdout():
                df = dart.finstate_all(corp_code, y, reprt_code='11011')
            if df is not None and not df.empty:
                all_dfs[y] = df
                print(f"  Fetched FY{y} annual report ({len(df)} items)")
        except Exception as e:
            print(f"  Warning: could not fetch FY{y}: {e}")
        time.sleep(0.15)

    if not all_dfs:
        raise RuntimeError("No annual report data available from DART.")

    # Build year -> value dicts for each metric.
    # Strategy: use the most recent report that covers each year.
    # For the latest report, thstrm=latest, frmtrm=prior, bfefrmtrm=2-prior.
    is_data: Dict[str, Dict[int, Optional[float]]] = {k: {} for k in IS_MAP}
    bs_data: Dict[str, Dict[int, Optional[float]]] = {k: {} for k in BS_MAP}
    cf_data: Dict[str, Dict[int, Optional[float]]] = {k: {} for k in CF_MAP}

    for report_year in sorted(all_dfs.keys(), reverse=True):
        df = all_dfs[report_year]
        # Map columns to years
        col_year_map = [
            (report_year, 'thstrm_amount'),
            (report_year - 1, 'frmtrm_amount'),
            (report_year - 2, 'bfefrmtrm_amount'),
        ]

        for key, names in IS_MAP.items():
            for yr, col in col_year_map:
                if yr in years and (yr not in is_data[key] or is_data[key][yr] is None):
                    val = _extract_from_df(df, names, col=col, sj_div='IS')
                    if val is not None:
                        is_data[key][yr] = val

        for key, names in BS_MAP.items():
            for yr, col in col_year_map:
                if yr in years and (yr not in bs_data[key] or bs_data[key][yr] is None):
                    val = _extract_from_df(df, names, col=col, sj_div='BS')
                    if val is not None:
                        bs_data[key][yr] = val

        for key, names in CF_MAP.items():
            for yr, col in col_year_map:
                if yr in years and (yr not in cf_data[key] or cf_data[key][yr] is None):
                    val = _extract_from_df(df, names, col=col, sj_div='CF')
                    if val is not None:
                        cf_data[key][yr] = val

    # Ensure every year key exists even if None
    for mapping_data in [is_data, bs_data, cf_data]:
        for key in mapping_data:
            for yr in years:
                mapping_data[key].setdefault(yr, None)

    # --- Capex fallback ---
    # Some companies break down capex into individual PPE items (건물의 취득,
    # 사무용비품의 취득, 기타유형자산의 취득, etc.) instead of reporting a
    # single 유형자산의 취득.  When the primary capex extraction is empty,
    # scan CF rows for all individual asset acquisitions and sum them.
    capex_all_none = all(cf_data['capex'].get(yr) is None for yr in years)
    if capex_all_none:
        for report_year in sorted(all_dfs.keys(), reverse=True):
            df = all_dfs[report_year]
            col_year_map = [
                (report_year, 'thstrm_amount'),
                (report_year - 1, 'frmtrm_amount'),
                (report_year - 2, 'bfefrmtrm_amount'),
            ]
            for yr, col in col_year_map:
                if yr not in years or cf_data['capex'].get(yr) is not None:
                    continue
                val = _sum_ppe_acquisitions(df, col=col)
                if val is not None:
                    cf_data['capex'][yr] = val

    # --- Derived metrics ---
    # Gross Profit = Revenue - COGS (if missing)
    for yr in years:
        rev = is_data['revenue'].get(yr)
        cogs = is_data['cogs'].get(yr)
        gp = is_data['gross_profit'].get(yr)
        if gp is None and rev is not None and cogs is not None:
            is_data['gross_profit'][yr] = rev - cogs
        elif cogs is None and rev is not None and gp is not None:
            is_data['cogs'][yr] = rev - gp

    # EBITDA = Operating Income + D&A (use CF D&A if IS doesn't have it)
    da_data: Dict[int, Optional[float]] = {}
    amort_data: Dict[int, Optional[float]] = {}
    for yr in years:
        # Try IS depreciation first, then CF adjustments
        da_val = is_data.get('da', {}).get(yr)
        if da_val is None:
            # Use CF adjustments as proxy
            da_val = cf_data.get('adjustments', {}).get(yr)
            # Actually adjustments includes more than D&A.  Better: try capex proxy.
            da_val = None  # Leave as None; will derive from CF
        da_data[yr] = da_val
        amort_data[yr] = is_data.get('amortization', {}).get(yr)

    ebitda: Dict[int, Optional[float]] = {}
    for yr in years:
        oi = is_data['operating_income'].get(yr)
        da = da_data.get(yr) or 0
        am = amort_data.get(yr) or 0
        if oi is not None:
            ebitda[yr] = oi + abs(da) + abs(am)
        else:
            ebitda[yr] = None

    # FCF = Operating CF - Capex
    fcf: Dict[int, Optional[float]] = {}
    for yr in years:
        opcf = cf_data['operating_cf'].get(yr)
        capex = cf_data['capex'].get(yr)
        if opcf is not None and capex is not None:
            fcf[yr] = opcf - abs(capex)
        else:
            fcf[yr] = None

    # Shares outstanding from DART
    shares_basic_data = _get_shares(stock_code, years)

    # Diluted shares: derive from basic shares * dilution ratio (EPS-based)
    # Note: NI from financial statements is consolidated total (incl. minority),
    # but EPS is based on parent-attributable NI, so we must use the EPS ratio.
    shares_diluted_data: Dict[int, Optional[float]] = {}
    for yr in years:
        basic_sh = shares_basic_data.get(yr)
        eps_b = is_data.get('eps_basic', {}).get(yr)
        eps_d = is_data.get('eps_diluted', {}).get(yr)
        if basic_sh and eps_b and eps_d and eps_d != 0:
            # diluted_shares = basic_shares * (basic_eps / diluted_eps)
            diluted_sh = basic_sh * (abs(eps_b) / abs(eps_d))
            shares_diluted_data[yr] = max(diluted_sh, basic_sh)
        else:
            shares_diluted_data[yr] = basic_sh

    # SBC placeholder (Korean filings rarely separate this)
    sbc: Dict[int, Optional[float]] = {yr: None for yr in years}

    # Build the financial_data dict matching project_monkey structure
    financial_data: Dict = {
        'years': years,
        'income_statement': {
            'revenue':          is_data['revenue'],
            'cogs':             is_data['cogs'],
            'gross_profit':     is_data['gross_profit'],
            'rd_expense':       is_data.get('rd_expense', {yr: None for yr in years}),
            'sga_expense':      is_data['sga_expense'],
            'operating_income': is_data['operating_income'],
            'da':               da_data,
            'amortization':     amort_data,
            'transformation_costs': {yr: None for yr in years},
            'debt_extinguishment': {yr: None for yr in years},
            'ebitda':           ebitda,
            'interest_expense': is_data['interest_expense'],
            'interest_income':  is_data['interest_income'],
            'other_income':     is_data.get('other_income', {yr: None for yr in years}),
            'pretax_income':    is_data['pretax_income'],
            'tax_expense':      is_data['tax_expense'],
            'net_income':       is_data['net_income'],
            'eps_basic':        is_data.get('eps_basic', {yr: None for yr in years}),
            'eps_diluted':      is_data.get('eps_diluted', {yr: None for yr in years}),
            'shares_basic':     shares_basic_data,
            'shares_diluted':   shares_diluted_data,
        },
        'balance_sheet': {
            'cash':             bs_data['cash'],
            'st_investments':   bs_data['st_investments'],
            'accounts_rec':     bs_data['accounts_rec'],
            'inventory':        bs_data['inventory'],
            'other_current_a':  bs_data['other_current_a'],
            'total_current_a':  bs_data['total_current_a'],
            'ppe_net':          bs_data['ppe_net'],
            'goodwill':         bs_data.get('goodwill', {yr: None for yr in years}),
            'intangibles':      bs_data.get('intangibles', {yr: None for yr in years}),
            'lt_investments':   bs_data.get('lt_investments', {yr: None for yr in years}),
            'other_noncurrent_a': bs_data.get('other_noncurrent_a', {yr: None for yr in years}),
            'total_assets':     bs_data['total_assets'],
            'accounts_pay':     bs_data['accounts_pay'],
            'accrued_liab':     bs_data.get('accrued_liab', {yr: None for yr in years}),
            'other_current_l':  bs_data.get('other_current_l', {yr: None for yr in years}),
            'st_debt':          bs_data.get('st_debt', {yr: None for yr in years}),
            'deferred_rev_cur': bs_data.get('deferred_rev_cur', {yr: None for yr in years}),
            'total_current_l':  bs_data['total_current_l'],
            'lt_debt':          bs_data.get('lt_debt', {yr: None for yr in years}),
            'deferred_tax_l':   bs_data.get('deferred_tax_l', {yr: None for yr in years}),
            'other_noncurrent_l': bs_data.get('other_noncurrent_l', {yr: None for yr in years}),
            'total_liabilities': bs_data['total_liabilities'],
            'common_stock':     bs_data.get('common_stock', {yr: None for yr in years}),
            'apic':             bs_data.get('apic', {yr: None for yr in years}),
            'retained_earnings': bs_data.get('retained_earnings', {yr: None for yr in years}),
            'treasury_stock':   bs_data.get('treasury_stock', {yr: None for yr in years}),
            'total_equity':     bs_data['total_equity'],
        },
        'cash_flow': {
            'net_income':   cf_data.get('net_income_cf', is_data['net_income']),
            'da':           da_data,
            'sbc':          sbc,
            'operating_cf': cf_data['operating_cf'],
            'capex':        cf_data.get('capex', {yr: None for yr in years}),
            'acquisitions': cf_data.get('acquisitions', {yr: None for yr in years}),
            'investing_cf': cf_data['investing_cf'],
            'dividends':    cf_data.get('dividends_paid', {yr: None for yr in years}),
            'repurchases':  cf_data.get('repurchases', {yr: None for yr in years}),
            'debt_issuance': cf_data.get('debt_issuance', {yr: None for yr in years}),
            'debt_repay':   cf_data.get('debt_repay', {yr: None for yr in years}),
            'financing_cf': cf_data['financing_cf'],
            'fx_effect':    cf_data.get('fx_effect', {yr: None for yr in years}),
            'fcf':          fcf,
        },
    }

    # Also fetch D&A from cash flow for the IS if IS didn't have it
    _fill_da_from_cf(financial_data, years)

    return financial_data


def _fill_da_from_cf(fd: Dict, years: List[int]) -> None:
    """If IS D&A is missing, try to use CF capex and PPE schedule to infer it.

    Fallback: use capex * 0.7 as a rough D&A proxy (many Korean cos don't
    separately tag D&A in IS).
    """
    da = fd['income_statement']['da']
    capex = fd['cash_flow']['capex']
    ppe = fd['balance_sheet']['ppe_net']

    for yr in years:
        if da.get(yr) is not None:
            continue
        # D&A = Beginning PPE + Capex - Ending PPE (if we have consecutive years)
        sorted_years = sorted(years)
        idx = sorted_years.index(yr) if yr in sorted_years else -1
        if idx > 0:
            prior_yr = sorted_years[idx - 1]
            ppe_begin = ppe.get(prior_yr)
            ppe_end = ppe.get(yr)
            capex_yr = capex.get(yr)
            if ppe_begin is not None and ppe_end is not None and capex_yr is not None:
                inferred_da = ppe_begin + abs(capex_yr) - ppe_end
                if inferred_da > 0:
                    da[yr] = inferred_da
                    continue
        # Rough fallback: 70% of capex
        capex_yr = capex.get(yr)
        if capex_yr is not None and abs(capex_yr) > 0:
            da[yr] = abs(capex_yr) * 0.7

    # Update EBITDA with corrected D&A
    ebitda = fd['income_statement']['ebitda']
    oi = fd['income_statement']['operating_income']
    am = fd['income_statement']['amortization']
    for yr in years:
        if oi.get(yr) is not None:
            d = abs(da.get(yr) or 0)
            a = abs(am.get(yr) or 0)
            ebitda[yr] = oi[yr] + d + a


def _get_shares(stock_code: str, years: List[int]) -> Dict[int, Optional[float]]:
    """Fetch shares outstanding via DART ``report('주식총수')`` or EPS derivation.

    Primary: DART share capital report (gives exact common share count).
    Fallback: Derive from Net Income / EPS Basic from financial statements.
    """
    result: Dict[int, Optional[float]] = {yr: None for yr in years}
    if not stock_code:
        return result

    dart = _get_dart()
    codes = dart.corp_codes
    match = codes[codes['stock_code'] == stock_code]
    if match.empty:
        return result
    corp_code = match.iloc[0]['corp_code']

    for yr in years:
        # Method 1: DART share capital report (주식총수)
        try:
            with _suppress_stdout():
                shares_df = dart.report(corp_code, '주식총수', yr)
            if shares_df is not None and not shares_df.empty:
                # Look for 보통주 (common stock) row -> distb_stock_co (distributable)
                for _, row in shares_df.iterrows():
                    se_val = str(row.get('se', ''))
                    if '보통주' in se_val or 'common' in se_val.lower():
                        distb = str(row.get('distb_stock_co', '0')).replace(',', '').strip()
                        if distb and distb != '-' and distb != '':
                            shares = float(distb)
                            if shares > 0:
                                result[yr] = shares
                                break
                # If no common stock row found, try total (합계)
                if result[yr] is None:
                    for _, row in shares_df.iterrows():
                        se_val = str(row.get('se', ''))
                        if '합계' in se_val or 'total' in se_val.lower():
                            distb = str(row.get('distb_stock_co', '0')).replace(',', '').strip()
                            if distb and distb != '-' and distb != '':
                                shares = float(distb)
                                if shares > 0:
                                    result[yr] = shares
                                    break
        except Exception:
            pass
        time.sleep(0.1)

        if result[yr] is not None:
            continue

        # Method 2: Derive from Net Income / EPS (from financial statements)
        try:
            with _suppress_stdout():
                fs = dart.finstate_all(corp_code, yr, reprt_code='11011')
            if fs is not None and not fs.empty:
                ni_val = _extract_from_df(
                    fs, ['당기순이익', '당기순이익(손실)'],
                    col='thstrm_amount', sj_div='IS')
                eps_val = _extract_from_df(
                    fs, ['기본주당이익', '기본주당순이익'],
                    col='thstrm_amount', sj_div='IS')
                if ni_val and eps_val and eps_val != 0:
                    result[yr] = abs(ni_val / eps_val)
        except Exception:
            pass

    # Forward-fill: if some years are still None, use the nearest available
    sorted_yrs = sorted(years)
    last_known = None
    for yr in sorted_yrs:
        if result[yr] is not None:
            last_known = result[yr]
        elif last_known is not None:
            result[yr] = last_known

    return result


# =========================================================================
# 3.  LTM (LAST TWELVE MONTHS) ANNUALIZATION
# =========================================================================

def compute_ltm(corp_code: str, latest_annual_year: int,
                financial_data: Dict) -> Optional[int]:
    """Attempt to compute LTM financials from quarterly reports.

    If a Q3 report exists for ``latest_annual_year + 1``, computes:
        LTM = Annual(Y) + Cumulative-Q3(Y+1) - Cumulative-Q3(Y)

    Returns the LTM label year (e.g. ``latest_annual_year + 1``) if successful,
    or *None* if no quarterly data is available.
    """
    dart = _get_dart()
    next_year = latest_annual_year + 1

    # Try Q3 (reprt_code='11014'), then semi-annual ('11012'), then Q1 ('11013')
    for reprt_code, label in [('11014', 'Q3'), ('11012', 'H1'), ('11013', 'Q1')]:
        try:
            with _suppress_stdout():
                q_current = dart.finstate_all(corp_code, next_year, reprt_code=reprt_code)
        except Exception:
            q_current = None
        if q_current is None or q_current.empty:
            continue

        # Also need same quarter of the prior year
        try:
            with _suppress_stdout():
                q_prior = dart.finstate_all(corp_code, latest_annual_year,
                                            reprt_code=reprt_code)
        except Exception:
            q_prior = None
        if q_prior is None or q_prior.empty:
            continue

        print(f"  Computing LTM using {label} {next_year} data...")

        # For IS and CF items: LTM = Annual(Y) + Q-cumulative(Y+1) - Q-cumulative(Y)
        # For BS items: just use the quarterly BS (point-in-time)
        years = financial_data['years']
        ltm_year = next_year  # label for the LTM column

        # Add ltm_year to years list
        if ltm_year not in years:
            years.insert(0, ltm_year)
            financial_data['years'] = years

        # IS metrics
        for key, names in IS_MAP.items():
            fd_key = key
            if fd_key not in financial_data['income_statement']:
                continue
            annual_val = financial_data['income_statement'][fd_key].get(latest_annual_year)
            q_cur_val = _extract_from_df(q_current, names, col='thstrm_amount', sj_div='IS')
            q_pri_val = _extract_from_df(q_prior, names, col='thstrm_amount', sj_div='IS')
            if annual_val is not None and q_cur_val is not None and q_pri_val is not None:
                financial_data['income_statement'][fd_key][ltm_year] = (
                    annual_val + q_cur_val - q_pri_val
                )
            else:
                financial_data['income_statement'][fd_key][ltm_year] = None

        # CF metrics
        for key, names in CF_MAP.items():
            fd_key = key
            if fd_key not in financial_data['cash_flow']:
                continue
            annual_val = financial_data['cash_flow'][fd_key].get(latest_annual_year)
            q_cur_val = _extract_from_df(q_current, names, col='thstrm_amount', sj_div='CF')
            q_pri_val = _extract_from_df(q_prior, names, col='thstrm_amount', sj_div='CF')
            if annual_val is not None and q_cur_val is not None and q_pri_val is not None:
                financial_data['cash_flow'][fd_key][ltm_year] = (
                    annual_val + q_cur_val - q_pri_val
                )
            else:
                financial_data['cash_flow'][fd_key][ltm_year] = None

        # Capex fallback for LTM: sum individual PPE items if primary failed
        if financial_data['cash_flow'].get('capex', {}).get(ltm_year) is None:
            annual_capex = financial_data['cash_flow'].get('capex', {}).get(latest_annual_year)
            q_cur_capex = _sum_ppe_acquisitions(q_current, col='thstrm_amount')
            q_pri_capex = _sum_ppe_acquisitions(q_prior, col='thstrm_amount')
            if annual_capex is not None and q_cur_capex is not None and q_pri_capex is not None:
                financial_data['cash_flow']['capex'][ltm_year] = (
                    annual_capex + q_cur_capex - q_pri_capex
                )

        # BS metrics: use quarterly BS directly (point-in-time)
        for key, names in BS_MAP.items():
            if key not in financial_data['balance_sheet']:
                continue
            val = _extract_from_df(q_current, names, col='thstrm_amount', sj_div='BS')
            financial_data['balance_sheet'][key][ltm_year] = val

        # Recompute derived metrics for LTM year
        # First fill D&A (was not in IS_MAP, needs PPE schedule inference).
        # Pass full years so the PPE-schedule method can find the prior year.
        _fill_da_from_cf(financial_data, financial_data['years'])
        _recompute_derived(financial_data, ltm_year)

        return ltm_year

    print("  No quarterly data available for LTM.")
    return None


def _recompute_derived(fd: Dict, yr: int) -> None:
    """Recompute EBITDA, GP, FCF for a given year."""
    inc = fd['income_statement']
    cf = fd['cash_flow']

    # Gross Profit
    rev = inc['revenue'].get(yr)
    cogs = inc['cogs'].get(yr)
    gp = inc['gross_profit'].get(yr)
    if gp is None and rev and cogs:
        inc['gross_profit'][yr] = rev - cogs

    # EBITDA
    oi = inc['operating_income'].get(yr)
    da = inc['da'].get(yr) or 0
    am = inc['amortization'].get(yr) or 0
    if oi is not None:
        inc['ebitda'][yr] = oi + abs(da) + abs(am)

    # FCF
    opcf = cf['operating_cf'].get(yr)
    capex = cf['capex'].get(yr)
    if opcf is not None and capex is not None:
        cf['fcf'][yr] = opcf - abs(capex)


# =========================================================================
# 4.  STOCK PRICES & MARKET DATA (pykrx)
# =========================================================================

def get_stock_price(stock_code: str, date: str = None) -> Optional[float]:
    """Fetch the closing price from pykrx.  *date* format: ``YYYYMMDD``."""
    if not _HAS_PYKRX:
        return None
    if date is None:
        date = datetime.today().strftime('%Y%m%d')
    # Try the exact date, then look back up to 7 days for holidays
    for offset in range(8):
        try:
            dt = datetime.strptime(date, '%Y%m%d') - timedelta(days=offset)
            d = dt.strftime('%Y%m%d')
            df = krx_stock.get_market_ohlcv_by_date(d, d, stock_code)
            if df is not None and not df.empty:
                close = float(df.iloc[0]['종가'])
                if close > 0:
                    return close
        except Exception:
            pass
    return None


def get_historical_prices(stock_code: str,
                          years: List[int]) -> Dict[int, Optional[float]]:
    """Fetch closing prices at each fiscal year-end (Dec 31)."""
    result: Dict[int, Optional[float]] = {}
    for yr in years:
        result[yr] = get_stock_price(stock_code, f'{yr}1231')
    return result


def get_current_price(stock_code: str,
                      target_date: str = None) -> Dict:
    """Fetch a closing price, returning ``{price, date}``."""
    if target_date:
        d = target_date.replace('-', '')
    else:
        d = datetime.today().strftime('%Y%m%d')
    price = get_stock_price(stock_code, d)
    return {'price': price, 'date': target_date or datetime.today().strftime('%Y-%m-%d')}


def get_market_cap(stock_code: str, date: str = None) -> Optional[float]:
    """Fetch market capitalisation from pykrx."""
    if not _HAS_PYKRX:
        return None
    if date is None:
        date = datetime.today().strftime('%Y%m%d')
    for offset in range(8):
        try:
            dt = datetime.strptime(date, '%Y%m%d') - timedelta(days=offset)
            d = dt.strftime('%Y%m%d')
            df = krx_stock.get_market_cap_by_date(d, d, stock_code)
            if df is not None and not df.empty:
                cap = float(df.iloc[0].get('시가총액', 0))
                if cap > 0:
                    return cap
        except Exception:
            pass
    return None


# =========================================================================
# 5.  WACC COMPONENTS
# =========================================================================

def get_kr_bond_yield() -> float:
    """Fetch the Korean 10-year government bond yield.

    Tries BOK ECOS statistics page, then falls back to a default.
    """
    # Try yfinance for Korean 10Y bond (symbol varies)
    if _HAS_YFINANCE:
        for sym in ['^KS10Y', '148070.KS']:  # Korea 10Y ETFs/indices
            try:
                with _suppress_stdout():
                    t = yf.Ticker(sym)
                    hist = t.history(period='5d')
                if hist is not None and not hist.empty:
                    val = float(hist['Close'].iloc[-1])
                    if 0 < val < 20:
                        print(f"  Korean 10Y bond yield (yfinance {sym}): {val:.2f}%")
                        return val / 100.0
            except Exception:
                pass

    # Default
    print("  Using default Korean 10Y bond yield: 3.50%")
    return 0.035


def get_kr_erp() -> float:
    """Return the Korean equity risk premium.

    Uses a research-based default.  Editable in the Excel WACC sheet.
    """
    print("  Korean Equity Risk Premium: 6.50% (default)")
    return 0.065


def _get_industry_name(induty_code: str) -> str:
    """Map a KSIC industry code to a Korean/English name."""
    # Top-level KSIC groupings (2-digit prefix)
    _KSIC_MAP = {
        '10': '식품 (Food)',
        '11': '음료 (Beverages)',
        '13': '섬유 (Textiles)',
        '14': '의복 (Apparel)',
        '20': '화학 (Chemicals)',
        '21': '의약품 (Pharmaceuticals)',
        '22': '고무·플라스틱 (Rubber & Plastics)',
        '23': '비금속광물 (Non-metallic Minerals)',
        '24': '금속 (Metals)',
        '25': '금속가공 (Fabricated Metals)',
        '26': '전자·반도체 (Electronics & Semiconductors)',
        '27': '의료기기 (Medical Devices)',
        '28': '전기장비 (Electrical Equipment)',
        '29': '기계 (Machinery)',
        '30': '자동차 (Automotive)',
        '31': '운송장비 (Transportation Equipment)',
        '32': '가구 (Furniture)',
        '33': '기타제조 (Other Manufacturing)',
        '35': '전기·가스 (Electricity & Gas)',
        '41': '건설 (Construction)',
        '45': '자동차판매 (Auto Dealers)',
        '46': '도매 (Wholesale)',
        '47': '소매 (Retail)',
        '49': '운수 (Transportation)',
        '52': '창고·물류 (Warehousing & Logistics)',
        '58': '출판 (Publishing)',
        '59': '영상·방송 (Broadcasting & Film)',
        '60': '방송통신 (Telecommunications)',
        '61': '통신 (Telecom)',
        '62': '소프트웨어 (Software)',
        '63': 'IT서비스 (IT Services & Portals)',
        '64': '금융 (Finance)',
        '65': '보험 (Insurance)',
        '66': '금융서비스 (Financial Services)',
        '68': '부동산 (Real Estate)',
        '70': '지주회사 (Holding Companies)',
        '71': '과학기술서비스 (Science & Tech Services)',
        '72': 'R&D (Research & Development)',
        '73': '광고 (Advertising)',
        '74': '전문서비스 (Professional Services)',
        '75': '사업지원 (Business Support)',
        '85': '교육 (Education)',
        '86': '보건의료 (Healthcare)',
        '90': '예술·스포츠 (Arts & Sports)',
    }
    # Sub-industry overrides (3+ digit prefixes)
    _KSIC_SUB = {
        '204': '화장품 (Cosmetics)',
    }
    if not induty_code:
        return 'Unknown'
    # Try longer prefix first
    for plen in (4, 3):
        sub = _KSIC_SUB.get(induty_code[:plen])
        if sub:
            return sub
    prefix2 = induty_code[:2]
    return _KSIC_MAP.get(prefix2, f'산업코드 {induty_code}')


def _search_listed_companies(keyword: str) -> list:
    """Search DART corp_codes for listed companies matching *keyword*.

    Returns list of ``{corp_code, corp_name, stock_code}`` dicts.
    """
    dart = _get_dart()
    codes = dart.corp_codes
    listed = codes[
        (codes['stock_code'].notna())
        & (codes['stock_code'] != '')
        & (codes['stock_code'].str.len() == 6)
    ]
    matches = listed[listed['corp_name'].str.contains(keyword, na=False)]
    results = []
    for _, row in matches.iterrows():
        results.append({
            'corp_code': row['corp_code'],
            'corp_name': row['corp_name'],
            'stock_code': row['stock_code'],
        })
    return results


def get_industry_peers(stock_code: str,
                       max_peers: int = 10,
                       industry_keyword: str = None,
                       auto_select: bool = False) -> Tuple[List[Dict], str]:
    """Find comparable companies by industry keyword search.

    If *industry_keyword* is given, searches DART for listed companies
    whose names contain the keyword.  Otherwise prompts the user for
    a search term (unless *auto_select* is True, in which case the
    DART industry code is used to pick a default keyword).
    """
    dart = _get_dart()
    codes = dart.corp_codes

    # Step 1: Determine target company's industry
    match = codes[codes['stock_code'] == stock_code]
    if match.empty:
        return [], 'Unknown'
    corp_code = match.iloc[0]['corp_code']
    corp_name = match.iloc[0]['corp_name']

    induty_code = ''
    try:
        info = dart.company(corp_code)
        induty_code = info.get('induty_code', '')
    except Exception:
        pass
    industry_name = _get_industry_name(induty_code)

    # Step 2: Determine search keyword(s)
    if not industry_keyword and not auto_select:
        print(f"\n  대상기업: {corp_name} ({stock_code})")
        print(f"  산업분류: {industry_name} (KSIC {induty_code})")
        print()
        print("  비교기업을 찾기 위한 산업 키워드를 입력하세요.")
        print("  예시: 반도체, 전자, 자동차, 화학, 바이오, 금융, IT, 화장품 등")
        print("  (여러 키워드는 쉼표로 구분, Enter = 자동 선택)")
        print()
        user_input = input("  산업 키워드 (Industry keywords): ").strip()
        if user_input:
            industry_keyword = user_input

    # Auto-select: derive keywords from industry code or company name
    if not industry_keyword:
        # Map common KSIC prefixes to search keywords
        _AUTO_KEYWORDS = {
            # 3-digit or longer prefixes (checked first for specificity)
            '204': ['화장품', '코스메틱', '뷰티', '클리오', '아모레',
                    '에이블', '마녀공장', '브이티', '삐아', '아이패밀리',
                    '디와이디', '제이준'],
            # 2-digit prefixes
            '26': ['전자', '반도체', '디스플레이', '하이닉스', 'SDI'],
            '20': ['화학', '석유화학', '케미칼'],
            '21': ['제약', '바이오', '셀트리온', '약품'],
            '28': ['배터리', '전기', '에너지', '2차전지'],
            '30': ['자동차', '모빌리티', '기아', '모비스'],
            '63': ['인터넷', '플랫폼', '포털', '카카오', '네이버'],
            '62': ['소프트웨어', 'IT', '시스템'],
            '64': ['금융', '은행', '증권', '지주'],
            '65': ['보험', '생명', '화재'],
            '10': ['식품', '음식', '제과', '음료'],
            '14': ['의류', '패션', '화장품'],
        }
        # Try longer prefixes first for specificity (e.g. '204' before '20')
        keywords = []
        if induty_code:
            for plen in (4, 3, 2):
                prefix = induty_code[:plen]
                keywords = _AUTO_KEYWORDS.get(prefix, [])
                if keywords:
                    break
        if not keywords:
            # Fall back to generic search using company name fragments
            keywords = [corp_name[:2]]  # first two chars of company name
        industry_keyword = ','.join(keywords)

    # Step 3: Search for matching companies
    keywords_list = [k.strip() for k in industry_keyword.split(',') if k.strip()]
    candidates = []
    seen_codes = set()
    seen_codes.add(stock_code)  # exclude target

    for kw in keywords_list:
        for company in _search_listed_companies(kw):
            sc = company['stock_code']
            if sc not in seen_codes:
                seen_codes.add(sc)
                candidates.append(company)

    if not candidates:
        print(f"  키워드 '{industry_keyword}'로 비교기업을 찾을 수 없습니다.")
        return [], industry_name

    print(f"  Found {len(candidates)} candidate companies for keywords: {industry_keyword}")

    # Step 4: Fetch peer data for candidates (scan all, keep top by mcap)
    peers = []
    for comp in candidates:
        peer = _get_peer_data(comp['stock_code'], comp['corp_code'],
                              comp['corp_name'])
        if peer and peer.get('price', 0) > 0 and peer.get('shares', 0) > 0:
            peers.append(peer)
        time.sleep(0.15)

    # Sort by market cap descending
    peers.sort(key=lambda p: p.get('shares', 0) * p.get('price', 0),
               reverse=True)

    return peers[:max_peers], industry_name


def _get_peer_data(stock_code: str, corp_code: str = '',
                   name: str = '') -> Optional[Dict]:
    """Fetch comparable company data for a single peer.

    Uses pykrx OHLCV (reliable) for price, DART for shares and financials,
    and yfinance for beta.
    """
    dart = _get_dart()

    # Resolve corp_code if not given
    if not corp_code:
        codes = dart.corp_codes
        match = codes[codes['stock_code'] == stock_code]
        if match.empty:
            return None
        corp_code = match.iloc[0]['corp_code']
        name = name or match.iloc[0]['corp_name']

    # Price from pykrx OHLCV (this API works reliably)
    price = get_stock_price(stock_code)
    if price is None or price <= 0:
        return None

    # Shares from DART -- try recent years (annual reports may lag 3-4 months)
    current_year = datetime.today().year
    try_years = [current_year - 1, current_year - 2, current_year - 3]
    shares_dict = _get_shares(stock_code, try_years)
    shares = 0
    for yr in try_years:
        shares = shares_dict.get(yr) or 0
        if shares > 0:
            break
    if shares <= 0:
        return None

    # Beta from yfinance
    beta = 1.0
    if _HAS_YFINANCE:
        for suffix in ['.KS', '.KQ']:
            try:
                yf_t = yf.Ticker(f'{stock_code}{suffix}')
                b = yf_t.info.get('beta')
                if b and 0.1 < b < 5.0:
                    beta = b
                    break
            except Exception:
                pass

    # Financial data from DART (latest annual report)
    # Use finstate_all for granular items (finstate only has summaries)
    total_debt = 0
    total_cash = 0
    net_income = None
    tax_rate = 0.22
    for yr in [current_year - 1, current_year - 2]:
        try:
            with _suppress_stdout():
                fs = dart.finstate_all(corp_code, yr, reprt_code='11011')
            if fs is None or fs.empty:
                continue

            def _get_val(names):
                for n in names:
                    rows = fs[fs['account_nm'] == n]
                    if not rows.empty:
                        return _parse_amount(rows.iloc[0]['thstrm_amount'])
                return None

            def _sum_debt():
                """Sum all interest-bearing debt items from BS."""
                debt_names = ['단기차입금', '장기차입금', '사채',
                              '단기사채', '장기사채', '유동성장기부채',
                              '유동성장기차입금']
                total = 0
                found = False
                for nm in debt_names:
                    v = _get_val([nm])
                    if v:
                        total += abs(v)
                        found = True
                if found:
                    return total
                # Fallback: some companies use generic 차입금 for both
                # current and non-current borrowings
                bs_rows = fs[fs['sj_div'] == 'BS']
                borrow = bs_rows[bs_rows['account_nm'] == '차입금']
                if not borrow.empty:
                    for _, row in borrow.iterrows():
                        v = _parse_amount(row.get('thstrm_amount'))
                        if v:
                            total += abs(v)
                    return total
                return 0

            total_debt = _sum_debt()
            cash = _get_val(['현금및현금성자산']) or 0
            total_cash = abs(cash)

            ni = _get_val(['당기순이익', '당기순이익(손실)'])
            if ni is not None:
                net_income = ni

            pretax = _get_val(['법인세비용차감전순이익',
                               '법인세비용차감전순이익(손실)'])
            tax = _get_val(['법인세비용', '법인세비용(수익)'])
            if pretax and pretax > 0 and tax:
                tax_rate = abs(tax) / abs(pretax)
            break
        except Exception:
            pass
        time.sleep(0.1)

    return {
        'name': name,
        'ticker': stock_code,
        'beta': beta,
        'price': price,
        'shares': shares,
        'total_debt': total_debt,
        'total_cash': total_cash,
        'net_debt': total_debt - total_cash,
        'tax_rate': min(tax_rate, 0.50),  # cap at 50%
        'net_income': net_income,
    }


# =========================================================================
# 6.  P/E COMPARABLE COMPANIES
# =========================================================================

def get_pe_comps_data(peers: List[Dict],
                      stock_code: str = '') -> List[Dict]:
    """Enrich peer data with P/E metrics.

    For each peer, fetches:
    - Market cap (already available from peer data)
    - FY Net Income (from DART)
    - LTM Net Income (from DART quarterly if available)
    - Forward EPS (from yfinance analyst estimates)
    - Trailing P/E and Forward P/E
    """
    dart = _get_dart()
    results = []
    current_year = datetime.today().year

    for peer in peers[:10]:
        code = peer['ticker']
        name = peer['name']
        market_cap = peer.get('shares', 0) * peer.get('price', 0)

        # Get FY Net Income from DART
        fy_ni = None
        try:
            codes = dart.corp_codes
            match = codes[codes['stock_code'] == code]
            if not match.empty:
                corp_code = match.iloc[0]['corp_code']
                for yr in [current_year - 1, current_year - 2]:
                    try:
                        with _suppress_stdout():
                            fs = dart.finstate(corp_code, yr, reprt_code='11011')
                        if fs is not None and not fs.empty:
                            ni_rows = fs[fs['account_nm'].isin(
                                ['당기순이익', '당기순이익(손실)'])]
                            if not ni_rows.empty:
                                fy_ni = _parse_amount(ni_rows.iloc[0]['thstrm_amount'])
                                break
                    except Exception:
                        pass
                    time.sleep(0.1)
        except Exception:
            pass

        # Forward EPS from yfinance
        forward_eps = None
        if _HAS_YFINANCE:
            try:
                yf_t = yf.Ticker(f'{code}.KS')
                forward_eps = yf_t.info.get('forwardEps')
            except Exception:
                pass

        # Compute P/E ratios
        trailing_pe = None
        forward_pe = None
        shares = peer.get('shares', 0)

        if fy_ni and fy_ni > 0 and market_cap > 0:
            trailing_pe = market_cap / fy_ni

        if forward_eps and forward_eps > 0 and peer.get('price', 0) > 0:
            forward_pe = peer['price'] / forward_eps

        results.append({
            'name': name,
            'stock_code': code,
            'market_cap': market_cap,
            'fy_ni': fy_ni,
            'ltm_ni': fy_ni,  # Same as FY unless LTM computed
            'trailing_pe': trailing_pe,
            'forward_eps': forward_eps,
            'forward_pe': forward_pe,
        })

        time.sleep(0.15)  # Rate limiting

    return results
