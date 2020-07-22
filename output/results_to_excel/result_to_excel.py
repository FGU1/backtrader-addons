###############################################################################
#
# Software program written by Neil Murphy in year 2020.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
#
# By using this software, the Disclaimer and Terms distributed with the
# software are deemed accepted, without limitation, by user.
#
# You should have received a copy of the Disclaimer and Terms document
# along with this program.  If not, see... https://bit.ly/2Tlr9ii
#
###############################################################################
import xlsxwriter
from pathlib import Path

"""
Module for converting to excel the standard and custom indicator dictionaries exported in the 
strategy object at the end of a Backtrader backtest. 
"""


def transactions(results, workbook=None, sheet_format=None):
    """
    Returns the transactions dataframe.
    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    workbook
    :return workbook: Excel workbook to be saved to disk.
    """
    trans_dict = results[0].analyzers.getbyname("transactions").get_analysis()

    columns = ["Date", "Units", "Price", "SID", "Ticker", "Value"]

    worksheet = workbook.add_worksheet("transaction")
    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:A", sheet_format["x_wide"], None)
    worksheet.set_column("C:C", sheet_format["medium"], sheet_format["float_2d"])
    worksheet.set_column("E:F", sheet_format["medium"], sheet_format["float_2d"])

    dates = []
    trans = []
    for d, v in trans_dict.items():
        for t in v:
            dates.append(d)
            trans.append(t)

    for n in range(len(dates)):
        date = [dates[n].strftime("%y-%m-%d %H:%M")]

        worksheet.write_row(n + 1, 0, date)
        worksheet.write_row(n + 1, 1, trans[n])

    return workbook


def unnest_trade_analysis(d, trade_analysis_dict, pk=""):
    """
    Recursive function that will create layered key names and attach the values.
    Used by the tradeanalyzer function.
    """
    for k, v in d.items():
        if isinstance(v, dict):
            unnest_trade_analysis(v, trade_analysis_dict, pk + "_" + k)
        else:
            trade_analysis_dict[(pk + "_" + k)[1:]] = v


def tradeanalyzer(results, workbook=None, sheet_format=None):
    """
    This is the trades analysis nested dictionary, converting to single row for insertion into table.
    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    workbook
    :return workbook: Excel workbook to be saved to disk.
    Return columns: There are many single point metrics available. Please refer to the following for reference.

    [
    'key',
    'total_total', 'total_open', 'total_closed',

    'streak_won_current', 'streak_won_longest', 'streak_lost_current', 'streak_lost_longest',

    'pnl_gross_total', 'pnl_gross_average', 'pnl_net_total', 'pnl_net_average',

    'won_total', 'won_pnl_total', 'won_pnl_average', 'won_pnl_max',

    'lost_total', 'lost_pnl_total', 'lost_pnl_average', 'lost_pnl_max',

    'long_total',
    'long_pnl_total', 'long_pnl_average',
    'long_pnl_won_total', 'long_pnl_won_average', 'long_pnl_won_max',
    'long_pnl_lost_total', 'long_pnl_lost_average', 'long_pnl_lost_max',
    'long_won', 'long_lost',

    'short_total',
    'short_pnl_total', 'short_pnl_average',
    'short_pnl_won_total', 'short_pnl_won_average', 'short_pnl_won_max',
    'short_pnl_lost_total', 'short_pnl_lost_average', 'short_pnl_lost_max',
    'short_won', 'short_lost',

    'len_total', 'len_average', 'len_max', 'len_min',
    'len_won_total', 'len_won_average', 'len_won_max', 'len_won_min',
    'len_lost_total', 'len_lost_average', 'len_lost_max', 'len_lost_min',
    'len_long_total', 'len_long_average', 'len_long_max', 'len_long_min',
    'len_long_won_total', 'len_long_won_average', 'len_long_won_max', 'len_long_won_min',
    'len_long_lost_total', 'len_long_lost_average', 'len_long_lost_max', 'len_long_lost_min',
    'len_short_total', 'len_short_average', 'len_short_max', 'len_short_min',
    'len_short_won_total', 'len_short_won_average', 'len_short_won_max', 'len_short_won_min',
    'len_short_lost_total', 'len_short_lost_average', 'len_short_lost_max', 'len_short_lost_min'
    ]
    """

    trades = results[0].analyzers.getbyname("trades").get_analysis()

    # Create empty dictionary for final keys:values, this will pass through recursive function unest_trade_analysis
    trade_analysis_dict = {}
    unnest_trade_analysis(trades, trade_analysis_dict, pk="")

    columns = [
        "Item",
        "Value",
    ]

    worksheet = workbook.add_worksheet("trade_analysis")

    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:B", sheet_format["x_wide"], sheet_format["align_left"])

    for i, (k, v) in enumerate(trade_analysis_dict.items()):
        worksheet.write_row(i + 1, 0, [k])
        worksheet.write_row(i + 1, 1, [v])

    return workbook


def drawdown_analysis(d, drawdown_analysis_dict, pk=""):
    """
    Recursive function that will create layered key names and attach the values.
    """
    for k, v in d.items():
        if isinstance(v, dict):
            drawdown_analysis(v, drawdown_analysis_dict, pk + "_" + k)
        else:
            drawdown_analysis_dict[(pk + "_" + k)[1:]] = v

    return drawdown_analysis_dict


def drawdown(results, workbook=None, sheet_format=None):
    """
    Calculates drawdown information.
    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    workbook
    :return workbook: Excel workbook to be saved to disk.
    """
    # Get the drawdowns auto ordered nested dictionary
    drawdown = results[0].analyzers.getbyname("drawdown").get_analysis()

    # Create empty dictionary for final keys:values
    drawdown_analysis_dict = {}
    drawdown_analysis_dict = drawdown_analysis(drawdown, drawdown_analysis_dict, pk="")

    columns = [
        "Item",
        "Value",
    ]

    worksheet = workbook.add_worksheet("drawdown")

    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:A", sheet_format["x_wide"], sheet_format["align_left"])
    worksheet.set_column("B:B", sheet_format["medium"], sheet_format["align_left"])

    for i, (k, v) in enumerate(drawdown_analysis_dict.items()):
        worksheet.write_row(i + 1, 0, [k])
        worksheet.write_row(i + 1, 1, [v])

    return workbook


def stats_analysis(d, stats_analysis_dict, pk=""):
    """
    Recursive function that will create layered key names and attach the values.
    """
    for k, v in d.items():
        if isinstance(v, dict):
            stats_analysis(v, stats_analysis_dict, pk + "_" + k)
        else:
            stats_analysis_dict[(pk + "_" + k)[1:]] = v

    return stats_analysis_dict


def periodstats(results, workbook=None, sheet_format=None):
    """
    Period Stats, basic stats for the period
    This is the basic stats analysis nested dictionary, converting to single row for insertion into table.

    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    :return workbook: Excel workbook to be saved to disk.
    """
    # Get the stats auto ordered nested dictionary
    stats = results[0].analyzers.getbyname("period_stats").get_analysis()

    # Create empty dictionary for final keys:values
    stats_analysis_dict = {}

    stats_analysis_dict = stats_analysis(stats, stats_analysis_dict, pk="")

    worksheet = workbook.add_worksheet("stats")

    columns = [
        "Item",
        "Value",
    ]
    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:B", sheet_format["x_wide"], sheet_format["align_left"])

    for i, (k, v) in enumerate(stats_analysis_dict.items()):
        worksheet.write_row(i + 1, 0, [k])
        worksheet.write_row(i + 1, 1, [v])

    return workbook


def returns(results, workbook=None, sheet_format=None):
    """
    Rate of return.

    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    :return workbook: Excel workbook to be saved to disk.
    """
    # Get the stats auto ordered nested dictionary
    rate_of_return = results[0].analyzers.getbyname("returns").get_analysis()

    worksheet = workbook.add_worksheet("return")

    columns = [
        "Item",
        "Value",
    ]
    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:B", sheet_format["x_wide"], sheet_format["float_5d"])

    for i, (k, v) in enumerate(rate_of_return.items()):
        worksheet.write_row(i + 1, 0, [k])
        worksheet.write_row(i + 1, 1, [v])

    return workbook


def vwr(results, workbook=None, sheet_format=None):
    """
    Value weighted return. A better sharpe ratio.
    https://www.crystalbull.com/sharpe-ratio-better-with-log-returns/

    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    :return workbook: Excel workbook to be saved to disk.
    """
    # Get the stats auto ordered nested dictionary
    vwr = results[0].analyzers.getbyname("VWR").get_analysis()

    worksheet = workbook.add_worksheet("VWR")

    columns = [
        "Item",
        "Value",
    ]
    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:A", sheet_format["medium"], sheet_format["align_left"])
    worksheet.set_column("B:B", sheet_format["medium"], sheet_format["float_5d"])

    for i, (k, v) in enumerate(vwr.items()):
        worksheet.write_row(i + 1, 0, [k])
        worksheet.write_row(i + 1, 1, [v])

    return workbook


def positionsvalue(results, workbook=None, sheet_format=None):
    """
    Tracks all positions over time.

    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    workbook
    :return workbook: Excel workbook to be saved to disk.
    """
    # Get the stats auto ordered nested dictionary
    position = results[0].analyzers.getbyname("position").get_analysis()

    worksheet = workbook.add_worksheet("positions")

    columns_date = ["Date"]
    columns_row = position["Datetime"]

    worksheet.write_row(0, 0, columns_date)
    worksheet.write_row(0, 1, columns_row)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:A", sheet_format["wide"], None)
    worksheet.set_column("B:B", sheet_format["wide"], sheet_format["float_2d"])
    worksheet.set_column("C:ZZ", sheet_format["medium"], sheet_format["float_2d"])

    for i, (k, v) in enumerate(position.items()):
        if i == 0:
            continue

        date = k.strftime("%Y-%m-%d %H:%M")
        worksheet.write_row(i, 0, [date])
        worksheet.write_row(i, 1, v)

    return workbook


def tradelist(results, workbook=None, sheet_format=None):
    """
    This analyzer prints a list of trades similar to amibroker, containing MFE and MAE

    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    :return workbook: Excel workbook to be saved to disk.
    """
    trade_list = results[0].analyzers.getbyname("trade_list").get_analysis()

    worksheet = workbook.add_worksheet("trade_list")
    columns = trade_list[0].keys()
    columns = [x.capitalize() for x in columns]
    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("D:D", sheet_format["x_wide"], None)
    worksheet.set_column("E:E", sheet_format["narrow"], sheet_format["float_2d"])
    worksheet.set_column("F:F", sheet_format["x_wide"], None)
    worksheet.set_column("G:G", sheet_format["narrow"], sheet_format["float_2d"])
    worksheet.set_column("H:H", sheet_format["narrow"], sheet_format["percent"])
    worksheet.set_column("I:I", sheet_format["narrow"], sheet_format["int_0d"])
    worksheet.set_column("J:J", sheet_format["narrow"], sheet_format["percent"])
    worksheet.set_column("L:M", sheet_format["narrow"], sheet_format["int_0d"])
    worksheet.set_column("O:O", sheet_format["narrow"], sheet_format["int_0d"])
    worksheet.set_column("P:P", sheet_format["narrow"], sheet_format["percent"])
    worksheet.set_column("Q:Q", sheet_format["narrow"], sheet_format["percent"])

    for i, d in enumerate(trade_list):
        d["datein"] = d["datein"].strftime("%Y-%m-%d %H:%M")
        d["dateout"] = d["dateout"].strftime("%Y-%m-%d %H:%M")

        worksheet.write_row(i + 1, 0, d.values())

    return workbook


def cashmarket(results, workbook=None, sheet_format=None):
    """
    Portfolio cash and total values.

    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    workbook
    :return workbook: Excel workbook to be saved to disk.
    """
    # Get the stats auto ordered nested dictionary
    value = results[0].analyzers.getbyname("cash_market").get_analysis()

    columns = [
        "Date",
        "Cash",
        "Value",
    ]

    worksheet = workbook.add_worksheet("value")

    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:B", sheet_format["medium"], sheet_format["float_2d"])

    for i, (k, v) in enumerate(value.items()):
        date = k.strftime("%y-%m-%d %H:%M")
        worksheet.write_row(i + 1, 0, [date])
        worksheet.write_row(i + 1, 1, v)

    return workbook


def ohlcv(results, workbook=None, sheet_format=None):
    """
    OHLCV

    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    workbook
    :return workbook: Excel workbook to be saved to disk.
    """
    # Get the stats auto ordered nested dictionary
    ohlcv = results[0].analyzers.getbyname("OHLCV").get_analysis()

    columns = ["Date", "Open", "High", "Low", "Close", "Volume"]

    worksheet = workbook.add_worksheet("ohlcv")

    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:A", sheet_format["wide"], None)
    worksheet.set_column("B:E", sheet_format["narrow"], sheet_format["float_2d"])
    worksheet.set_column("F:F", sheet_format["medium"], sheet_format["int_0d"])

    for i, (k, v) in enumerate(ohlcv.items()):
        if i == 0:
            continue

        date = k.strftime("%Y-%m-%d %H:%M")
        worksheet.write_row(i, 0, [date])
        worksheet.write_row(i, 1, v)

    return workbook


def tradeclosed(results, workbook=None, sheet_format=None):
    """
    Closed trades, pnl, commission, and duration.
    :param workbook: Excel workbook to be saved to disk.
    :param results: Backtest results.
    :param sheet_format: Dictionary holding formatting information such as col width, font etc.
    :return workbook: Excel workbook to be saved to disk.
    """
    trade_dict = results[0].analyzers.getbyname("trade_closed").get_analysis()

    columns = [
        "Date Closed",
        "Time",
        "Ticker",
        "PnL",
        "PnL Comm",
        "Commission",
        "Days Open",
    ]

    worksheet = workbook.add_worksheet("trade")

    worksheet.write_row(0, 0, columns)

    worksheet.set_row(0, None, sheet_format["header_format"])

    worksheet.set_column("A:A", sheet_format["wide"], None)
    worksheet.set_column("B:B", sheet_format["medium"], None)
    worksheet.set_column("D:E", sheet_format["narrow"], sheet_format["float_2d"])
    worksheet.set_column("F:F", sheet_format["medium"], sheet_format["float_2d"])
    worksheet.set_column("G:G", sheet_format["medium"], sheet_format["int_0d"])

    for i, value in enumerate(trade_dict.values()):
        worksheet.write_row(i + 1, 0, [value[0].strftime("%Y-%m-%d")])
        worksheet.write_row(i + 1, 1, [value[0].strftime("%H:%M")])
        worksheet.write_row(i + 1, 2, value[1:])

    return workbook


def save_to_excel(results, save_path, save_name):
    """ Extraction of analyzer lines in dictionary form and save to excel file. """

    # If there are no transactions, return None and test number.
    if (
        len(
            [an for an in results[0].analyzers if type(an).__name__ == "Transactions"][
                0
            ].get_analysis()
        )
        == 0
    ):
        print(f"Backtest has no transactions. No excel workbook saved.")
        return

    path = Path(save_path)
    path.mkdir(parents=True, exist_ok=True)
    filename = save_name
    filepath = path / filename

    # Create workbook.
    workbook = xlsxwriter.Workbook(filepath)

    # Add some cell formats.
    sheet_format = dict(
        narrow=8,
        medium=12,
        wide=16,
        x_wide=20,
        header_format=workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "align": "center",
                "font_color": "black",
                # "border": 1,
            }
        ),
        float_2d=workbook.add_format({"num_format": "#,##0.00"}),
        float_5d=workbook.add_format({"num_format": "#,##0.00000"}),
        int_0d=workbook.add_format({"num_format": "#,##0"}),
        percent=workbook.add_format({"num_format": "0%"}),
        align_left=workbook.add_format({"align": "left"}),
    )

    for analyzer in results[0].analyzers:
        try:
            workbook = eval(type(analyzer).__name__.lower())(
                results, workbook, sheet_format
            )
        except:
            pass

    workbook.close()

    return None
