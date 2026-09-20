#!/usr/bin/env python3
"""
build_workbook.py

Builds the Overview page of the Channel Ledger dashboard for Tableau, styled to
match dashboard_mockup/channel_ledger.html:

    channel_performance.hyper   Tableau extract made from cleaned_data/channel_performance.csv
    marketing_dashboard.twbx    packaged workbook (workbook XML + the extract)

    Overview = masthead, Period / Channel controls, page title with a live scope label,
               6 KPI tiles with sparklines, Revenue & Spend line, ROAS line, Revenue by Channel

Tableau Public only opens extracts (not live Excel/CSV connections), hence the .hyper.

How the controls work: Period and Channel are dashboard parameters, and every measure is written as
SUM(IF [In scope] THEN ... END), so the whole dashboard follows them. Worksheet filters were tried first
and Tableau rejected every form written from this script, hence parameters.

Charts | Tables: two dashboards ("Overview" and "Overview - Tables") with the same header and tiles.
A pair of navigation buttons in the filter bar switches between them; Period and Channel carry across.
The file declares Tableau's newer format features (document-format-change-manifest) and gives every
sheet, dashboard and window an id, because buttons point at their target by id.

Known limits (Tableau, not the data): Channel is single-select; Tableau shows "N nulls" badges on charts
when a period leaves months empty (right-click the badge > Hide Indicator); the ROAS axis cannot be given
an "x" suffix from the file; hover dimming and styled tooltips are not available; in Tableau's editor a
navigation button needs Alt+click (in presentation mode / on the web a normal click works).

Requires:  pip install tableauhyperapi pandas
Run:       python tableau/build_workbook.py            (dark theme, as in the mockup)
           python tableau/build_workbook.py light      (light theme)
Then open marketing_dashboard.twbx in Tableau Desktop or Tableau Public.
"""

import os
import sys
import uuid
import xml.etree.ElementTree as ET
import zipfile
from xml.sax.saxutils import escape

import pandas as pd
from tableauhyperapi import (Connection, CreateMode, Date, HyperProcess, Inserter, SqlType,
                             TableDefinition, TableName, Telemetry)

def guid(key):
    """Stable GUID for a sheet/window, in the {UPPER-CASE} form Tableau uses."""
    return "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, "channel-ledger/" + key)).upper() + "}"


HERE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(HERE, "..", "cleaned_data", "channel_performance.csv")
HYPER = os.path.join(HERE, "channel_performance.hyper")
TWB_NAME = "marketing_dashboard.twb"
TWBX = os.path.join(HERE, "marketing_dashboard.twbx")
HYPER_IN_TWBX = "Data/Extracts/channel_performance.hyper"

DS = "federated.1q2w3e4r5t6y7u8i9o0p1a2s3d4f"   # datasource id
CONN = "hyper.0z9x8c7v6b5n4m3l2k1j0h9g8f7d"      # named-connection id
Q = "&quot;"

# ---------------------------------------------------------------- theme (tokens from the mockup)
THEMES = {
    "dark": dict(page="#0d1013", surface="#171b20", ink="#f2f4f6", ink2="#b9c0c9", muted="#8c949e",
                 grid="#262b32", axis="#39404a", hair="#2c3137", spark="#5a626d", slate="#8f99a6",
                 ch=["#3987e5", "#d95926", "#199e70", "#c98500", "#d55181"]),
    "light": dict(page="#eceff2", surface="#fcfcfd", ink="#12171e", ink2="#454c57", muted="#646c77",
                  grid="#e6e9ed", axis="#c8cdd4", hair="#e1e4e8", spark="#a3abb5", slate="#5b6673",
                  ch=["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"]),
}
THEME = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] in THEMES else "dark"
T = THEMES[THEME]
SANS, SANS_SB, MONO = "Segoe UI", "Segoe UI Semibold", "Consolas"
CHANNELS = ["Google Ads", "Meta", "Email", "Organic", "Affiliate"]
CHANNEL_COLORS = list(zip(CHANNELS, T["ch"]))

# Calculated fields. name -> (caption, formula)
SCOPE = "[Calculation_5000000000000001]"   # boolean: row is inside the chosen Period and Channel
CALCS = {
    "[Calculation_1000000000000001]": ("ROAS", f"SUM(IF {SCOPE} AND [spend_missing]=0 THEN [revenue] END) / SUM(IF {SCOPE} THEN [spend] END)"),
    "[Calculation_1000000000000002]": ("AOV", f"SUM(IF {SCOPE} THEN [revenue] END) / SUM(IF {SCOPE} THEN [orders] END)"),
    "[Calculation_1000000000000003]": ("Conv Rate", f"SUM(IF {SCOPE} THEN [orders] END) / SUM(IF {SCOPE} THEN [sessions] END)"),
}
ROAS, AOV, CVR = list(CALCS)
DEFAULT_FMT = {"[month_start]": "*mmm yy", ROAS: '*0"x"'}


def dfmt(name):
    return f" default-format='{DEFAULT_FMT[name].replace(chr(34), '&quot;')}'" if name in DEFAULT_FMT else ""


CALC_META = {}   # per-calc (datatype, role, type); default is real / measure / quantitative

# ---- dashboard parameters: Period and Channel drive every measure through the "In scope" test
PERIODS = ["All time", "FY 2025", "H1 2026", "Last 3 months"]
CHOICES = ["All channels"] + CHANNELS
P1, P2 = "[Parameter 1]", "[Parameter 2]"


def _param(name, caption, values):
    return (f"<column caption='{caption}' datatype='string' name='{name}' param-domain-type='list' role='measure' "
            f"type='nominal' value='{Q}{values[0]}{Q}'><calculation class='tableau' formula='{Q}{values[0]}{Q}' />"
            "<members>" + "".join(f"<member value='{Q}{v}{Q}' />" for v in values) + "</members></column>")


PARAM_COLUMNS = _param(P1, "Period", PERIODS) + _param(P2, "Channel", CHOICES)
PARAM_DS = ("  <datasource hasconnection='false' inline='true' name='Parameters' version='18.1'>\n"
            "    <aliases enabled='yes' />\n    " + PARAM_COLUMNS + "\n  </datasource>")
PARAM_REF = "<datasource hasconnection='false' name='Parameters' />"
PARAM_DEP = f"<datasource-dependencies datasource='Parameters'>{PARAM_COLUMNS}</datasource-dependencies>"
_p1, _p2 = "[Parameters].[Parameter 1]", "[Parameters].[Parameter 2]"
CALCS[SCOPE] = ("In scope",
                f'({_p1}="All time" OR ({_p1}="FY 2025" AND YEAR([month_start])=2025) '
                f'OR ({_p1}="H1 2026" AND [month_start]>=#2026-01-01#) '
                f'OR ({_p1}="Last 3 months" AND [month_start]>=#2026-04-01#)) '
                f'AND ({_p2}="All channels" OR [channel]={_p2})')
CALC_META[SCOPE] = ("boolean", "dimension", "nominal")

# base measures, already limited to the chosen scope
REV, SPEND, ORD, SESS = (f"[Calculation_600000000000000{i}]" for i in range(1, 5))
for _nm, _cap, _col in ((REV, "Revenue", "revenue"), (SPEND, "Spend", "spend"),
                        (ORD, "Orders", "orders"), (SESS, "Sessions", "sessions")):
    CALCS[_nm] = (_cap, f"SUM(IF {SCOPE} THEN [{_col}] END)")

# sparkline series: each month minus its average over the scope, so the line always straddles zero and
# Tableau's automatic axis fills the height in every period and channel
_S = SCOPE
_E = {  # per-month expression, and whether it is additive (per-month average) or a ratio (scope-wide ratio)
    "Revenue": (f"SUM(IF {_S} THEN [revenue] END)", True),
    "Spend": (f"SUM(IF {_S} THEN [spend] END)", True),
    "ROAS": (f"SUM(IF {_S} AND [spend_missing]=0 THEN [revenue] END) / SUM(IF {_S} THEN [spend] END)", False),
    "Orders": (f"SUM(IF {_S} THEN [orders] END)", True),
    "AOV": (f"SUM(IF {_S} THEN [revenue] END) / SUM(IF {_S} THEN [orders] END)", False),
    "Conv Rate": (f"SUM(IF {_S} THEN [orders] END) / SUM(IF {_S} THEN [sessions] END)", False),
}
_MONTHS = "MIN({FIXED : COUNTD(IF " + _S + " THEN [month_start] END)})"
SPK = {}
for _i, (_k, (_e, _additive)) in enumerate(_E.items(), start=1):
    _nm = f"[Calculation_800000000000000{_i}]"
    _base = "MIN({FIXED : " + _e + "})"
    CALCS[_nm] = (f"{_k} trend", f"({_e}) - ({_base} / {_MONTHS})" if _additive else f"({_e}) - ({_base})")
    SPK[_k] = _nm

# live label under the page title, e.g. "Jan 2026 – Jun 2026 · 1 of 5 channels"
SCOPE_LBL = "[Calculation_7000000000000001]"
CALCS[SCOPE_LBL] = ("Scope", f'MIN(CASE {_p1} WHEN "All time" THEN "Jan 2025 – Jun 2026" '
                    f'WHEN "FY 2025" THEN "Jan 2025 – Dec 2025" WHEN "H1 2026" THEN "Jan 2026 – Jun 2026" '
                    f'WHEN "Last 3 months" THEN "Apr 2026 – Jun 2026" END) + "  ·  " + '
                    f'IF {_p2}="All channels" THEN "5 of 5 channels" ELSE "1 of 5 channels" END')
CALC_META[SCOPE_LBL] = ("string", "measure", "nominal")
CH_CALC = {}
for _i, _ch in enumerate(CHANNELS):
    _nm = f"[Calculation_200000000000000{_i + 1}]"
    CALCS[_nm] = (_ch, f'SUM(IF {SCOPE} AND [channel]="{_ch}" THEN [revenue] END)')
    CH_CALC[_ch] = _nm

# Extract columns: (name, datatype, role, type, remote-type)
COLUMNS = [
    ("month_start", "date", "dimension", "ordinal", 133),
    ("month", "string", "dimension", "nominal", 129),
    ("channel", "string", "dimension", "nominal", 129),
    ("spend", "real", "measure", "quantitative", 5),
    ("spend_missing", "integer", "measure", "quantitative", 20),
    ("revenue", "real", "measure", "quantitative", 5),
    ("orders", "integer", "measure", "quantitative", 20),
    ("sessions", "integer", "measure", "quantitative", 20),
    ("bounced_sessions", "real", "measure", "quantitative", 5),
]
HYPER_TYPES = {"date": SqlType.date(), "string": SqlType.text(), "real": SqlType.double(),
               "integer": SqlType.big_int()}


def build_extract():
    """channel_performance.csv -> Tableau .hyper extract (table [Extract].[Extract])."""
    df = pd.read_csv(CSV)
    cols = [TableDefinition.Column(n, HYPER_TYPES[dt]) for n, dt, *_ in COLUMNS]
    td = TableDefinition(TableName("Extract", "Extract"), cols)
    with HyperProcess(Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hp:
        with Connection(hp.endpoint, HYPER, CreateMode.CREATE_AND_REPLACE) as conn:
            conn.catalog.create_schema("Extract")
            conn.catalog.create_table(td)
            with Inserter(conn, td) as ins:
                for r in df.itertuples(index=False):
                    y, m, d = (int(x) for x in r.month_start.split("-"))
                    ins.add_row([
                        Date(y, m, d), r.month, r.channel,
                        None if pd.isna(r.spend) else float(r.spend), int(r.spend_missing),
                        float(r.revenue), int(r.orders), int(r.sessions), float(r.bounced_sessions),
                    ])
                ins.execute()
    return len(df)


def datasource_xml():
    recs = []
    for i, (n, dt, role, ty, rt) in enumerate(COLUMNS):
        agg = "Year" if dt == "date" else ("Count" if dt == "string" else "Sum")
        coll = "        <collation flag='0' name='LEN_RUS' />\n" if dt == "string" else ""
        recs.append(
            f"      <metadata-record class='column'>\n"
            f"        <remote-name>{n}</remote-name>\n        <remote-type>{rt}</remote-type>\n"
            f"        <local-name>[{n}]</local-name>\n        <parent-name>[Extract]</parent-name>\n"
            f"        <remote-alias>{n}</remote-alias>\n        <ordinal>{i}</ordinal>\n"
            f"        <local-type>{dt}</local-type>\n        <aggregation>{agg}</aggregation>\n"
            f"        <contains-null>true</contains-null>\n{coll}"
            f"      </metadata-record>"
        )
    calcs = "\n".join(
        f"  <column caption='{cap}' datatype='{CALC_META.get(nm, ('real',))[0]}'{dfmt(nm)} name='{nm}' "
        f"role='{CALC_META.get(nm, ('', 'measure'))[1]}' type='{CALC_META.get(nm, ('', '', 'quantitative'))[2]}'>\n"
        f"    <calculation class='tableau' formula='{f}' />\n  </column>"
        for nm, (cap, f) in CALCS.items()
    )
    declared = "\n".join(
        f"  <column datatype='{dt}'{dfmt('[' + n + ']')} name='[{n}]' role='{role}' type='{ty}' />"
        for n, dt, role, ty, _ in COLUMNS
    )
    maps = "\n".join(f"        <map to='{hx}'><bucket>{Q}{nm}{Q}</bucket></map>" for nm, hx in CHANNEL_COLORS)
    # colour follows the entity: assign the channel palette under every field spelling Tableau might key on
    channel_enc = "".join(
        f"      <encoding attr='color' field='{fld}' type='palette'>\n{maps}\n      </encoding>\n"
        for fld in ("[none:channel:nk]", "[channel]", f"[{DS}].[none:channel:nk]")
    )
    return f"""<datasource caption='channel_performance' inline='true' name='{DS}' version='18.1'>
  <connection class='federated'>
    <named-connections>
      <named-connection caption='channel_performance' name='{CONN}'>
        <connection authentication='auth-none' author-locale='en_US' class='hyper' dbname='{HYPER_IN_TWBX}' default-settings='yes' sslmode='' username='tableau_internal_user' />
      </named-connection>
    </named-connections>
    <relation connection='{CONN}' name='Extract' table='[Extract].[Extract]' type='table' />
    <metadata-records>
{chr(10).join(recs)}
    </metadata-records>
  </connection>
  <column datatype='string' name='[:Measure Names]' role='dimension' type='nominal' />
{declared}
{calcs}
  <layout dim-ordering='alphabetic' dim-percentage='0.5' measure-ordering='alphabetic' measure-percentage='0.4' show-structure='true' />
  <style>
    <style-rule element='mark'>
{channel_enc}      <encoding attr='color' field='[:Measure Names]' type='palette'>
        <map to='{T["ink"]}'><bucket>{Q}[{DS}].[usr:{REV[1:-1]}:qk]{Q}</bucket></map>
        <map to='{T["slate"]}'><bucket>{Q}[{DS}].[usr:{SPEND[1:-1]}:qk]{Q}</bucket></map>
{chr(10).join(f"        <map to='{hx}'><bucket>{Q}[{DS}].[usr:{CH_CALC[nm].strip('[]')}:qk]{Q}</bucket></map>" for nm, hx in CHANNEL_COLORS)}
      </encoding>
    </style-rule>
  </style>
</datasource>"""


def field(inst):
    return f"[{DS}].[{inst}]"


def dep_column(name):
    if name in CALCS:
        cap, f = CALCS[name]
        dt, role, ty = CALC_META.get(name, ("real", "measure", "quantitative"))
        return (f"<column caption='{cap}' datatype='{dt}' name='{name}' role='{role}' type='{ty}'>"
                f"<calculation class='tableau' formula='{f}' /></column>")
    for n, dt, role, ty, _ in COLUMNS:
        if f"[{n}]" == name:
            return f"<column datatype='{dt}' name='{name}' role='{role}' type='{ty}' />"
    raise KeyError(name)


def instance(name, kind):
    """kind: sum | usr | month | none. Returns (instance-name, <column-instance/>)."""
    base = name.strip("[]")
    if kind == "sum":
        inst, deriv, typ = f"[sum:{base}:qk]", "Sum", "quantitative"
    elif kind == "usr":
        inst, deriv, typ = f"[usr:{base}:qk]", "User", "quantitative"
    elif kind == "usrn":
        inst, deriv, typ = f"[usr:{base}:nk]", "User", "nominal"
    elif kind == "month":
        inst, deriv, typ = f"[tmn:{base}:qk]", "Month-Trunc", "quantitative"
    else:
        inst, deriv, typ = f"[none:{base}:nk]", "None", "nominal"
    return inst, f"<column-instance column='{name}' derivation='{deriv}' name='{inst}' pivot='key' type='{typ}' />"


def fmt(code):
    return code.replace('"', Q)


# ---------------------------------------------------------------- worksheet formatting
def fmt_rule(element, *formats):
    """<style-rule element=..> from (attr, value[, extra-attrs]) tuples."""
    body = ""
    for f in formats:
        attr, value, *extra = f
        more = "".join(f" {k}='{v}'" for k, v in (extra[0].items() if extra else []))
        body += f"<format attr='{attr}'{more} value='{value}' />"
    return f"<style-rule element='{element}'>{body}</style-rule>"


def line_rules(grid_color, zero_color, visible=True):
    """Gridline + zeroline rules under every attribute spelling Tableau might honour."""
    scopes = ("rows", "cols")
    g, z = [], []
    for col, out in ((grid_color, g), (zero_color, z)):
        out += [("color", col), ("stroke-color", col)]
        out += [("color", col, {"class": "0", "scope": s}) for s in scopes]
        out += [("stroke-color", col, {"class": "0", "scope": s}) for s in scopes]
    if not visible:
        for out in (g, z):
            out += [("line-visibility", "off")]
            out += [("line-visibility", "off", {"class": "0", "scope": s}) for s in scopes]
    else:
        g += [("line-visibility", "off", {"class": "0", "scope": "cols"})]
    return fmt_rule("gridline", *g) + fmt_rule("zeroline", *z)


def sheet_style(kind, text_x=None, textfmt=None, size=None, axis_fields=(), axis_formats=(), fixed_range=None,
                num_formats=(), palette=None, color_x=None, bg_color=None):
    """Worksheet-level <style>: dark/light surface, hairline grid, mono axis labels."""
    bg = bg_color or T["surface"]
    r = [
        fmt_rule("worksheet", ("background-color", bg), ("font-family", SANS), ("font-size", "9"),
                 ("color", T["muted"])),
        fmt_rule("table", ("background-color", bg), ("border-style", "none")),
        fmt_rule("pane", ("background-color", bg), ("border-style", "none")),
    ]
    if kind in ("kpi", "meta", "button"):
        fam, col, align = {"kpi": (SANS_SB, T["ink"], "left"), "meta": (MONO, T["muted"], "right"),
                           "button": (SANS_SB, T["ink2"], "center")}[kind]
        parts = ([("text-format", textfmt, {"field": text_x})] if textfmt else []) + [
            ("font-family", fam), ("font-size", str(size)), ("color", col), ("text-align", align),
            ("background-color", bg)] + ([("width", "560")] if kind == "meta" else [])
        r.append(fmt_rule("cell", *parts))
    elif kind == "spark":
        hide = "".join(
            f"<format attr='display' class='0' field='{fld}' scope='{sc}' value='false' />" for fld, sc in axis_fields
        )
        space = ""
        if fixed_range:
            fld, (lo, hi) = fixed_range
            space = (f"<encoding attr='space' class='0' field='{fld}' field-type='quantitative' include-zero='false' "
                     f"range-type='auto' scope='rows' type='space' />")
        r += [
            f"<style-rule element='axis'>{hide}<format attr='color' value='{bg}' /><format attr='tick-color' value='{bg}' />"
            f"<format attr='line-visibility' value='off' /><format attr='stroke-color' value='{bg}' />{space}</style-rule>",
            fmt_rule("header", ("color", bg), ("background-color", bg)),
            line_rules(bg, bg, visible=False),
        ]
    else:  # line / bar chart
        titles = "".join(
            f"<format attr='title' class='0' field='{fld}' scope='{sc}' value='' />" for fld, sc in axis_fields
        )
        fmts = "".join(
            f"<format attr='text-format' class='0' field='{fld}' scope='{sc}' value='{nf}' />"
            for fld, sc, nf in axis_formats
        )
        also = "".join(
            f"<format attr='text-format' class='0' field='{fld}' scope='{sc}' value='{nf}' />"
            for fld, sc, nf in axis_formats
        )
        r += [
            f"<style-rule element='header'><format attr='font-family' value='{MONO}' /><format attr='font-size' value='8' />"
            f"<format attr='color' value='{T['muted']}' /><format attr='background-color' value='{bg}' />{also}</style-rule>",
            f"<style-rule element='cell'>{also}</style-rule>",
            f"<style-rule element='axis'><format attr='font-family' value='{MONO}' /><format attr='font-size' value='8' />"
            f"<format attr='color' value='{T['muted']}' /><format attr='tick-color' value='{T['axis']}' />"
            f"<format attr='stroke-color' value='{T['axis']}' />{fmts}{titles}</style-rule>",
            line_rules(T["grid"], T["axis"]),
        ]
        if num_formats:
            r.append("<style-rule element='cell'>" + "".join(
                f"<format attr='text-format' field='{fld}' value='{nf}' />" for fld, nf in num_formats) + "</style-rule>")
        if palette and color_x:
            maps = "".join(f"<map to='{hx}'><bucket>{Q}{nm}{Q}</bucket></map>" for nm, hx in palette)
            r.append(f"<style-rule element='mark'><encoding attr='color' field='{color_x}' type='palette'>"
                     f"{maps}</encoding></style-rule>")
    return "<style>" + "".join(r) + "</style>"


def worksheet(name, cols=None, rows=None, mark="Automatic", text=None, color=None, textfmt=None,
              measure_values=None, kind="chart", size=22, mark_color=None, label_last=False,
              stroke=False, palette=None, bar_size=None, stack_top_down=None,
              bg_color=None):
    """cols/rows: list of (column, kind). text/color: (column, kind). measure_values: list of (column, kind)."""
    deps, insts = {}, {}

    def use(col, k):
        inst, xml = instance(col, k)
        deps[col] = dep_column(col)
        insts[inst] = xml
        return field(inst.strip("[]"))

    cols_x = [use(c, k) for c, k in (cols or [])]
    rows_x = [use(c, k) for c, k in (rows or [])]
    text_x = use(*text) if text else None
    color_x = use(*color) if color else None

    view_extra, mv = "", []
    if measure_values:
        mv = [use(c, k) for c, k in measure_values]
        members = "".join(
            f"<groupfilter function='member' level='[:Measure Names]' member='{Q}{m}{Q}' />" for m in mv
        )
        order = ""
        if stack_top_down:
            buckets = "".join(f"<bucket>{Q}{field(instance(c, k)[0].strip('[]'))}{Q}</bucket>" for c, k in stack_top_down)
            order = (f"<sort class='manual' column='[{DS}].[:Measure Names]' direction='ASC'>"
                     f"<dictionary>{buckets}</dictionary></sort>")
        view_extra = (
            f"<filter class='categorical' column='[{DS}].[:Measure Names]'>"
            f"<groupfilter function='union' user:op='manual'>{members}</groupfilter></filter>"
            f"{order}<slices><column>[{DS}].[:Measure Names]</column></slices>"
        )
        deps[":mn"] = "<column datatype='string' name='[:Measure Names]' role='dimension' type='nominal' />"
        rows_x = [f"[{DS}].[Multiple Values]"]
        color_x = f"[{DS}].[:Measure Names]"

    axis_fields = [(x, "cols") for x in cols_x] + [(x, "rows") for x in rows_x]
    axis_formats = []
    for cx in cols_x:
        base = cx.split("].[", 1)[1].rstrip("]")               # tmn:month_start:qk
        for spelling in (cx, f"[{base}]", f"[{DS}].[month_start]", "[month_start]"):
            axis_formats.append((spelling, "cols", "*mmm yy"))
    num_formats = []
    for rx in rows_x:
        if "Calculation_1000000000000001" in rx:
            base = rx.split("].[", 1)[1].rstrip("]")
            for spelling in (rx, f"[{base}]"):
                axis_formats.append((spelling, "rows", fmt('n0"x"')))
            num_formats.append((rx, fmt('n0.00"x"')))
        elif "Multiple Values" not in rx:
            num_formats.append((rx, fmt('c"$"#,##0')))
    num_formats += [(m, fmt('c"$"#,##0')) for m in mv]
    style = sheet_style(kind, text_x=text_x, textfmt=textfmt, size=size, num_formats=num_formats,
                        axis_fields=axis_fields, axis_formats=axis_formats, palette=palette, color_x=color_x,
                        fixed_range=None,
                        bg_color=bg_color)

    mark_fmt = ""
    if mark_color:
        mark_fmt += f"<format attr='mark-color' value='{mark_color}' />"
    if stroke:
        mark_fmt += f"<format attr='has-stroke' value='true' /><format attr='stroke-color' value='{T['surface']}' />"
    if bar_size:
        mark_fmt += f"<format attr='size' value='{bar_size}' />"
    if label_last:
        mark_fmt += ("<format attr='mark-labels-show' value='true' /><format attr='mark-labels-cull' value='false' />"
                     "<format attr='mark-labels-mode' value='most-recent' />")
    pane_style = f"<style><style-rule element='mark'>{mark_fmt}</style-rule></style>" if mark_fmt else ""
    sizing = "<mark-sizing mark-sizing-setting='marks-scaling-off' />" if bar_size else ""

    enc = ""
    if text_x:
        enc += f"<text column='{text_x}' />"
    if color_x:
        enc += f"<color column='{color_x}' />"
    enc = f"<encodings>{enc}</encodings>" if enc else ""

    return f"""<worksheet name='{name}'>
  <table>
    <view>
      <datasources><datasource caption='channel_performance' name='{DS}' />{PARAM_REF}</datasources>
      <datasource-dependencies datasource='{DS}'>{''.join(deps.values())}{''.join(insts.values())}</datasource-dependencies>
      {PARAM_DEP}
      {view_extra}
      <aggregation value='true' />
    </view>
    {style}
    <panes>
      <pane selection-relaxation-option='selection-relaxation-allow'>
        <view><breakdown value='auto' /></view>
        <mark class='{mark}' />
        {sizing}
        {enc}
        {pane_style}
      </pane>
    </panes>
    <rows>{' / '.join(rows_x)}</rows>
    <cols>{' * '.join(cols_x)}</cols>
  </table>
  <simple-id uuid='{guid("sheet/" + name)}' />
</worksheet>"""


def table_sheet(name, measures):
    """Text table for the Tables view: one row per month, one column per measure."""
    mn = f"[{DS}].[:Measure Names]"
    deps = dep_column("[month]") + "".join(dep_column(m) for m in measures)
    insts = instance("[month]", "none")[1] + "".join(instance(m, "usr")[1] for m in measures)
    mv = [field(instance(m, "usr")[0].strip("[]")) for m in measures]
    members = "".join(f"<groupfilter function='member' level='[:Measure Names]' member='{Q}{x}{Q}' />" for x in mv)
    buckets = "".join(f"<bucket>{Q}{x}{Q}</bucket>" for x in mv)
    bg = T["surface"]
    cw = 620 if len(measures) == 1 else 380 if len(measures) == 2 else 300   # roomy columns so tables fill their card
    cell_fmts = "".join(
        f"<format attr='text-format' field='{x}' value='{fmt('n0.00\"x\"') if m == ROAS else fmt('c\"$\"#,##0')}' />"
        for x, m in zip(mv, measures))
    style = ("<style>"
             + fmt_rule("worksheet", ("background-color", bg), ("font-family", SANS), ("font-size", "9"),
                        ("color", T["muted"]))
             + fmt_rule("table", ("background-color", bg), ("border-style", "none"), ("band-color", bg),
                        ("band-color", bg, {"class": "0", "scope": "rows"}))
             + fmt_rule("pane", ("background-color", bg), ("border-style", "none"))
             + fmt_rule("header", ("font-family", MONO), ("font-size", "9"), ("color", T["muted"]),
                        ("background-color", bg))
             + f"<style-rule element='cell'><format attr='font-family' value='{SANS}' /><format attr='font-size' value='10' />"
               f"<format attr='color' value='{T['ink']}' /><format attr='text-align' value='right' />"
               f"<format attr='background-color' value='{bg}' /><format attr='width' value='{cw}' />{cell_fmts}</style-rule>"
             + line_rules(T["grid"], T["axis"]) + "</style>")
    return f"""<worksheet name='{name}'>
  <table>
    <view>
      <datasources><datasource caption='channel_performance' name='{DS}' />{PARAM_REF}</datasources>
      <datasource-dependencies datasource='{DS}'>{deps}<column datatype='string' name='[:Measure Names]' role='dimension' type='nominal' />{insts}</datasource-dependencies>
      {PARAM_DEP}
      <filter class='categorical' column='{mn}'><groupfilter function='union' user:op='manual'>{members}</groupfilter></filter>
      <sort class='manual' column='{mn}' direction='ASC'><dictionary>{buckets}</dictionary></sort>
      <slices><column>{mn}</column></slices>
      <aggregation value='true' />
    </view>
    {style}
    <panes>
      <pane selection-relaxation-option='selection-relaxation-allow'>
        <view><breakdown value='auto' /></view>
        <mark class='Text' />
        <encodings><text column='[{DS}].[Multiple Values]' /></encodings>
      </pane>
    </panes>
    <rows>{field('none:month:nk')}</rows>
    <cols>{mn}</cols>
  </table>
  <simple-id uuid='{guid("sheet/" + name)}' />
</worksheet>"""


# ---------------------------------------------------------------- sheets
KPI = [  # label, sheet, (column, kind), number format, sparkline colour
    ("Revenue", "Revenue", (REV, "usr"), fmt('c"$"#,##0')),
    ("Spend", "Spend", (SPEND, "usr"), fmt('c"$"#,##0')),
    ("ROAS", "ROAS", (ROAS, "usr"), fmt('n0.00"x"')),
    ("Orders", "Orders", (ORD, "usr"), fmt("n#,##0")),
    ("Avg order value", "AOV", (AOV, "usr"), fmt('c"$"#,##0')),
    ("Conversion rate", "Conv Rate", (CVR, "usr"), fmt("p0.0%")),
]
MONTH = ("[month_start]", "month")
sheets = []
for i, (label, sheet, meas, f) in enumerate(KPI):
    sheets.append(worksheet(sheet, mark="Text", text=meas, textfmt=f, kind="kpi", size=44 if i == 0 else 22))
    sheets.append(worksheet(sheet + " trend", cols=[MONTH], rows=[(SPK[sheet], "usr")], mark="Line", kind="spark",
                            mark_color=T["spark"]))
sheets.append(worksheet("Revenue and Spend by Month", cols=[MONTH], mark="Line", label_last=True,
                        measure_values=[(REV, "usr"), (SPEND, "usr")]))
sheets.append(worksheet("ROAS by Month", cols=[MONTH], rows=[(ROAS, "usr")], mark="Line",
                        mark_color=T["ink"], label_last=True))
sheets.append(worksheet("Revenue by Channel", cols=[MONTH], mark="Bar", stroke=True, bar_size="2.5",
                        measure_values=[(CH_CALC[c], "usr") for c in CHANNELS],
                        stack_top_down=[(CH_CALC[c], "usr") for c in reversed(CHANNELS)]))

sheets.append(worksheet("Scope", mark="Text", text=(SCOPE_LBL, "usrn"), kind="meta", size=9, bg_color=T["page"]))
TABLES = {
    "Revenue and Spend table": [REV, SPEND],
    "ROAS table": [ROAS],
    "Revenue by Channel table": [CH_CALC[c] for c in CHANNELS],
}
for _n, _m in TABLES.items():
    sheets.append(table_sheet(_n, _m))

# ---------------------------------------------------------------- dashboard layout (1400 x 900, absolute 0-100000 units)
DASH_W, DASH_H = 1400, 900
zid = [10]


def nid():
    zid[0] += 1
    return zid[0]


def zstyle(bg=None, border=None, margin=0, padding=0):
    f = ""
    if border:
        f += (f"<format attr='border-color' value='{border}' /><format attr='border-style' value='solid' />"
              "<format attr='border-width' value='1' />")
    else:
        f += "<format attr='border-style' value='none' /><format attr='border-width' value='0' />"
    f += f"<format attr='margin' value='{margin}' /><format attr='padding' value='{padding}' />"
    if bg:
        f += f"<format attr='background-color' value='{bg}' />"
    return f"<zone-style>{f}</zone-style>"


def run(text, size=10, color=None, bold=False, font=SANS, align=None):
    a = f" fontalignment='{align}'" if align is not None else ""
    b = " bold='true'" if bold else ""
    return (f"<run fontname='{font}' fontsize='{size}' fontcolor='{color or T['ink']}'{b}{a}>"
            f"{escape(text)}</run>")


NL = "<run>Æ&#10;</run>"


def fixed_attr(fixed):
    return f" fixed-size='{fixed}' is-fixed='true'" if fixed else ""


def sheet_zone(name, x, y, w, h, bg=None, border=None, margin=0, padding=0, title=False, fixed=None):
    st = "" if title else " show-title='false'"
    return (f"<zone{fixed_attr(fixed)} h='{h}' id='{nid()}' name='{name}'{st} w='{w}' x='{x}' y='{y}'>"
            f"{zstyle(bg, border, margin, padding)}</zone>")


def text_zone(runs, x, y, w, h, bg=None, border=None, margin=0, padding=0, fixed=None):
    return (f"<zone{fixed_attr(fixed)} h='{h}' id='{nid()}' type-v2='text' w='{w}' x='{x}' y='{y}'>"
            f"<formatted-text>{runs}</formatted-text>{zstyle(bg, border, margin, padding)}</zone>")


def container(param, x, y, w, h, inner, bg=None, border=None, margin=0, padding=0, fixed=None):
    return (f"<zone{fixed_attr(fixed)} h='{h}' id='{nid()}' param='{param}' type-v2='layout-flow' w='{w}' x='{x}' y='{y}'>"
            f"{inner}{zstyle(bg, border, margin, padding)}</zone>")


def card(x, y, w, h, title, subtitle, sheet_name, legend_runs=None):
    """Mockup card: surface panel, hairline border, title + muted subtitle, legend top-right, chart below."""
    head_h = int(h * 0.20)
    left_w = int(w * (0.66 if legend_runs else 1.0))
    head = text_zone(run(title, 11, T["ink"], True, SANS_SB) + NL + run(subtitle, 9, T["muted"]),
                     x, y, left_w, head_h, bg=T["surface"], padding=2)
    if legend_runs:
        head += text_zone(legend_runs, x + left_w, y, w - left_w, head_h, bg=T["surface"], padding=2)
    hdr = container("horz", x, y, w, head_h, head, bg=T["surface"], fixed=44)
    chart = sheet_zone(sheet_name, x, y + head_h, w, h - head_h, bg=T["surface"])
    return container("vert", x, y, w, h, hdr + chart, bg=T["surface"], border=T["hair"], margin=6, padding=8)


def kpi_tile(i, x, y, w, h, label, sheet):
    lab_h, val_h = int(h * 0.20), int(h * (0.46 if i == 0 else 0.40))
    spark_h = h - lab_h - val_h
    inner = (text_zone(run(label, 9, T["muted"], font=SANS), x, y, w, lab_h, bg=T["surface"], padding=4, fixed=24)
             + sheet_zone(sheet, x, y + lab_h, w, val_h, bg=T["surface"], fixed=86 if i == 0 else 44)
             + sheet_zone(sheet + " trend", x, y + lab_h + val_h, w, spark_h, bg=T["surface"]))
    return container("vert", x, y, w, h, inner, bg=T["surface"], border=T["hair"], margin=0, padding=8)


H_BRAND, H_FILT, H_HEAD, H_KPI, H_MID, H_BOT = 5000, 5000, 6000, 21000, 31500, 31500
DASH_CHARTS, DASH_TABLES = "Overview", "Overview - Tables"


def param_zone(param, x, w, y):
    return (f"<zone h='{H_FILT}' id='{nid()}' mode='compact' param='[Parameters].{param}' type-v2='paramctrl' "
            f"w='{w}' x='{x}' y='{y}'>{zstyle(T['surface'], padding=4)}</zone>")


def window_id(dash_name):
    """Stable GUID Tableau uses to identify a dashboard window (buttons point at it)."""
    return guid("window/" + dash_name)


def view_toggle(tables, y):
    """Charts | Tables switch: the current view is a static pill, the other is a navigation button."""
    def pill(label, x):
        return text_zone(run(label, 10, T["surface"], True, SANS_SB, align=1), x, y, 8000, H_FILT,
                         bg=T["ink"], padding=14, margin=8)

    def button(label, target, x):
        return (f"<zone h='{H_FILT}' id='{nid()}' type-v2='dashboard-object' w='8000' x='{x}' y='{y}'>"
                f"<button action='tabdoc:goto-sheet window-id=&quot;{window_id(target)}&quot;' button-type='text'>"
                f"<button-visual-state><caption>{label}</caption>"
                f"<button-caption-font-style fontcolor='{T['ink2']}' fontname='{SANS_SB}' fontsize='10' />"
                f"<format attr='background-color' value='{T['hair']}' /></button-visual-state></button>"
                f"{zstyle(None, margin=8)}</zone>")
    if tables:
        return button("Charts", DASH_CHARTS, 84000) + pill("Tables", 92000)
    return pill("Charts", 84000) + button("Tables", DASH_TABLES, 92000)


def build_dashboard(dash_name, tables):
    """The Overview dashboard; tables=True swaps each chart for its data table."""
    y = 0
    brand = text_zone(
        run("Channel Ledger", 13, T["ink"], True, SANS_SB) + run("     channel_performance.hyper", 8, T["muted"], font=MONO),
        0, y, 100000, H_BRAND, bg=T["surface"], padding=12, fixed=46)
    y += H_BRAND
    filter_bar = container(
        "horz", 0, y, 100000, H_FILT,
        param_zone(P1, 0, 15000, y) + param_zone(P2, 15000, 15000, y)
        + text_zone(run(" ", 8, T["muted"]), 30000, y, 54000, H_FILT, bg=T["surface"])
        + view_toggle(tables, y),
        bg=T["surface"], fixed=58)
    y += H_FILT
    page_head = container("horz", 0, y, 100000, H_HEAD,
                          text_zone(run("How are we doing?", 17, T["ink"], True, SANS_SB), 0, y, 60000, H_HEAD,
                                    bg=T["page"], padding=8)
                          + sheet_zone("Scope", 60000, y, 40000, H_HEAD, bg=T["page"]), bg=T["page"], fixed=48)
    y += H_HEAD
    hero_w = 26000
    tile_w = (100000 - hero_w) // 5
    tiles, x = "", 0
    for i, (label, sheet, *_r) in enumerate(KPI):
        w = hero_w if i == 0 else tile_w
        tiles += kpi_tile(i, x, y, w, H_KPI, label, sheet)
        x += w
    kpi_row = container("horz", 0, y, 100000, H_KPI, tiles, bg=T["page"], margin=6, fixed=236)
    y += H_KPI
    leg_lines = (run("▬ ", 9, T["ink"], align=2) + run("Revenue     ", 9, T["ink2"], align=2)
                 + run("▬ ", 9, T["slate"], align=2) + run("Spend", 9, T["ink2"], align=2))
    leg_ch = "".join(run("■ ", 9, c, align=2) + run(nm + "    ", 9, T["ink2"], align=2) for nm, c in CHANNEL_COLORS)
    s_rs, s_roas, s_ch = (("Revenue and Spend table", "ROAS table", "Revenue by Channel table") if tables else
                          ("Revenue and Spend by Month", "ROAS by Month", "Revenue by Channel"))
    mid_row = container("horz", 0, y, 100000, H_MID,
                        card(0, y, 58000, H_MID, "Revenue and spend by month",
                             "One dollar axis, so the gap between the lines is the margin over spend.",
                             s_rs, None if tables else leg_lines)
                        + card(58000, y, 42000, H_MID, "ROAS by month",
                               "Revenue per dollar of spend, all selected channels.", s_roas),
                        bg=T["page"])
    y += H_MID
    bot_row = container("horz", 0, y, 100000, H_BOT,
                        card(0, y, 100000, H_BOT, "Revenue by channel",
                             "Monthly revenue stacked by the channel on the order.", s_ch, None if tables else leg_ch),
                        bg=T["page"])
    stack = container("vert", 0, 0, 100000, 100000, brand + filter_bar + page_head + kpi_row + mid_row + bot_row,
                      bg=T["page"])
    return f"""<dashboard name='{dash_name}'>
  <style />
  <size sizing-mode='automatic' />
  <datasources><datasource name='Parameters' /></datasources>
  <datasource-dependencies datasource='Parameters'>{PARAM_COLUMNS}</datasource-dependencies>
  <zones>
    <zone h='100000' id='4' type-v2='layout-basic' w='100000' x='0' y='0'>{stack}{zstyle(T["page"])}</zone>
  </zones>
  <simple-id uuid='{guid("dashboard/" + dash_name)}' />
</dashboard>"""


dash_charts = build_dashboard(DASH_CHARTS, tables=False)
dash_tables = build_dashboard(DASH_TABLES, tables=True)

common = [s for _l, s, *_r in KPI] + [s + " trend" for _l, s, *_r in KPI] + ["Scope"]
names_charts = common + ["Revenue and Spend by Month", "ROAS by Month", "Revenue by Channel"]
names_tables = common + list(TABLES)
names = names_charts + list(TABLES)
win_sheets = "".join(
    f"<window class='worksheet' name='{n}'><cards>"
    "<edge name='left'><strip size='160'><card type='pages' /><card type='filters' /><card type='marks' /></strip></edge>"
    "<edge name='top'><strip size='2147483647'><card type='columns' /></strip><strip size='2147483647'><card type='rows' /></strip>"
    "<strip size='31'><card type='title' /></strip></edge></cards></window>"
    for n in names
)


def _vp(lst):
    return "".join(f"<viewpoint name='{n}' />" for n in lst)


windows = (f"<windows source-height='30'>"
           f"<window class='dashboard' maximized='true' name='{DASH_CHARTS}'><viewpoints>{_vp(names_charts)}</viewpoints>"
           f"<active id='-1' /><simple-id uuid='{window_id(DASH_CHARTS)}' /></window>"
           f"<window class='dashboard' name='{DASH_TABLES}'><viewpoints>{_vp(names_tables)}</viewpoints>"
           f"<active id='-1' /><simple-id uuid='{window_id(DASH_TABLES)}' /></window>{win_sheets}</windows>")

doc = f"""<?xml version='1.0' encoding='utf-8' ?>
<workbook original-version='18.1' source-build='2022.1.0 (20221.22.0324.1508)' source-platform='win' version='18.1' xmlns:user='http://www.tableausoftware.com/xml/user'>
  <document-format-change-manifest>
    <BasicButtonObject />
    <BasicButtonObjectTextSupport />
    <NavigationAction />
    <SheetIdentifierTracking />
    <WindowsPersistSimpleIdentifiers />
  </document-format-change-manifest>
  <preferences>
    <preference name='ui.encoding.shelf.height' value='24' />
    <preference name='ui.shelf.height' value='26' />
  </preferences>
  <datasources>
{PARAM_DS}
{datasource_xml()}
  </datasources>
  <worksheets>
{chr(10).join(sheets)}
  </worksheets>
  <dashboards>
{dash_charts}
{dash_tables}
  </dashboards>
  {windows}
</workbook>
"""

ET.fromstring(doc.encode("utf-8"))  # fail loudly if the XML is not well formed
n_rows = build_extract()
with zipfile.ZipFile(TWBX, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr(TWB_NAME, doc.encode("utf-8"))
    zf.write(HYPER, HYPER_IN_TWBX)
print(f"theme   : {THEME}")
print(f"extract : {HYPER} ({n_rows} rows)")
print(f"workbook: {TWBX} ({len(sheets)} sheets + 2 dashboards)")
