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
from datetime import datetime

import backtrader as bt
import numpy as np


class TradeList(bt.analyzers.Analyzer):
    """
    Trade list similar to Amibroker output: Columns are:
    [ref, ticker, dir, datein, pricein, dateout, priceout, chng%,
    pnl, pnl%, size, value, cumpnl, nbars, pnl/bar, mfe%, mae%]
    """

    def get_analysis(self):

        return self.trades

    def __init__(self):

        self.trades = []
        self.cumprofit = 0.0

    def notify_trade(self, trade):

        if trade.isclosed:

            brokervalue = self.strategy.broker.getvalue()

            dir = "short"
            if trade.history[0].event.size > 0:
                dir = "long"

            pricein = trade.history[len(trade.history) - 1].status.price
            priceout = trade.history[len(trade.history) - 1].event.price
            datein = bt.num2date(trade.history[0].status.dt)
            dateout = bt.num2date(trade.history[len(trade.history) - 1].status.dt)

            if trade.data._timeframe >= bt.TimeFrame.Days:
                datein = datein.date()
                dateout = dateout.date()

            pcntchange = 100 * priceout / pricein - 100
            pnl = trade.history[len(trade.history) - 1].status.pnlcomm
            pnlpcnt = 100 * pnl / brokervalue
            barlen = trade.history[len(trade.history) - 1].status.barlen
            pbar = pnl / barlen
            self.cumprofit += pnl

            size = value = 0.0
            for record in trade.history:
                if abs(size) < abs(record.status.size):
                    size = record.status.size
                    value = record.status.value

            highest_in_trade = max(trade.data.high.get(ago=0, size=barlen + 1))
            lowest_in_trade = min(trade.data.low.get(ago=0, size=barlen + 1))
            hp = 100 * (highest_in_trade - pricein) / pricein
            lp = 100 * (lowest_in_trade - pricein) / pricein
            if dir == "long":
                mfe = hp
                mae = lp
            if dir == "short":
                mfe = -lp
                mae = -hp

            self.trades.append(
                {
                    "ref": trade.ref,
                    "ticker": trade.data._name,
                    "dir": dir,
                    "datein": datein,
                    "pricein": pricein,
                    "dateout": dateout,
                    "priceout": priceout,
                    "chng%": round(pcntchange, 2),
                    "pnl": pnl,
                    "pnl%": round(pnlpcnt, 2),
                    "size": size,
                    "value": value,
                    "cumpnl": self.cumprofit,
                    "nbars": barlen,
                    "pnl/bar": round(pbar, 2),
                    "mfe%": round(mfe, 2),
                    "mae%": round(mae, 2),
                }
            )


class CashMarket(bt.analyzers.Analyzer):
    """
    Analyzer returning cash and market values
    """

    def start(self):
        super(CashMarket, self).start()

    def create_analysis(self):
        self.rets = {}
        self.vals = 0.0

    def notify_cashvalue(self, cash, value):
        self.vals = (cash, value)
        self.rets[self.strategy.datetime.datetime()] = self.vals

    def get_analysis(self):
        return self.rets


class TradeClosed(bt.analyzers.Analyzer):
    """
    Analyzer returning closed trade information.
    """

    def start(self):
        super(TradeClosed, self).start()

    def create_analysis(self):
        self.rets = {}
        self.vals = tuple()

    def notify_trade(self, trade):
        """Receives trade notifications before each next cycle"""
        if trade.isclosed:
            self.vals = (
                self.strategy.datetime.datetime(),
                trade.data._name,
                round(trade.pnl, 2),
                round(trade.pnlcomm, 2),
                trade.commission,
                (trade.dtclose - trade.dtopen),
            )
            self.rets[trade.ref] = self.vals

    def get_analysis(self):
        return self.rets


class OHLCV(bt.analyzers.Analyzer):
    """This analyzer reports the OHLCV of each of datas.
    Params:
      - timeframe (default: ``None``)
        If ``None`` then the timeframe of the 1st data of the system will be
        used
      - compression (default: ``None``)
        Only used for sub-day timeframes to for example work on an hourly
        timeframe by specifying "TimeFrame.Minutes" and 60 as compression
        If ``None`` then the compression of the 1st data of the system will be
        used
    Methods:
      - get_analysis
        Returns a dictionary with returns as values and the datetime points for
        each return as keys
    """

    def start(self):
        tf = min(d._timeframe for d in self.datas)
        self._usedate = bt.TimeFrame.Minutes
        self.rets = {}

    def next(self):

        self.rets[self.data.datetime.datetime()] = [
            self.datas[0].open[0],
            self.datas[0].high[0],
            self.datas[0].low[0],
            self.datas[0].close[0],
            self.datas[0].volume[0],
        ]

    def get_analysis(self):
        return self.rets
