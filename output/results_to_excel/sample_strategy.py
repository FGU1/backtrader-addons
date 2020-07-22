import backtrader as bt
from result_to_excel import save_to_excel
from analyzer import TradeList, OHLCV, TradeClosed, CashMarket

class Strategy(bt.Strategy):

    params = (
        ("period", 14),
        ("lowerband", 30),
        ("upperband", 70),
    )

    def __init__(self):
        self.rsi = bt.ind.RSI(
            period=self.p.period, lowerband=self.p.lowerband, upperband=self.p.upperband,
        )

    def log(self, txt, dt=None):
        """ Logging function fot this strategy"""
        dt = dt or self.data.datetime[0]
        if isinstance(dt, float):
            dt = bt.num2date(dt)
        if self.datas[0]._timeframe >= bt.TimeFrame.Days:
            print("%s, %s" % (dt.date(), txt))
        else:
            print("%s, %s" % (dt, txt))

    def print_signal(self):
        self.log(
            "open {:7.2f}\thigh {:7.2f}\tlow {:7.2f}\tclose {:7.2f}\tvolume {:7.0f}\trsi {:5.0f}".format(
                self.datas[0].open[0],
                self.datas[0].high[0],
                self.datas[0].low[0],
                self.datas[0].close[0],
                self.datas[0].volume[0],
                self.rsi[0],
            )
        )

    def notify_order(self, order):
        if order.status in [order.Submitted, order.Accepted]:
            # Buy/Sell order submitted/accepted to/by broker - Nothing to do
            return

        # Check if an order has been completed
        # Attention: broker could reject order if not enougth cash
        if order.status in [order.Canceled, order.Margin]:
            if order.isbuy():
                self.log("BUY FAILED, Cancelled or Margin")
        if order.status in [order.Completed, order.Canceled, order.Margin]:
            if order.isbuy():
                self.log(
                    "BUY EXECUTED, Price: %.2f, Cost: %.2f, Comm %.2f"
                    % (order.executed.price, order.executed.value, order.executed.comm)
                )

                self.buyprice = order.executed.price
                self.buycomm = order.executed.comm
            else:  # Sell
                self.log(
                    "SELL EXECUTED, Price: %.2f, Cost: %.2f, Comm %.2f"
                    % (order.executed.price, order.executed.value, order.executed.comm)
                )

            self.bar_executed = len(self)

        # Write down: no pending order
        self.order = None

    def notify_trade(self, trade):
        if not trade.isclosed:
            return

        self.log("OPERATION PROFIT, GROSS %.2f, NET %.2f" % (trade.pnl, trade.pnlcomm))

    def next(self):
        self.print_signal()

        if not self.position:
            if self.rsi <= self.p.lowerband:
                self.buy(size=3)
        elif self.rsi >= self.p.upperband:
            self.close()

def main(tf):
    cerebro = bt.Cerebro()

    if tf == "day":
        data = bt.feeds.GenericCSVData(
            dataname="data/day_2year.csv",
            dtformat=("%Y-%m-%d"),
            timeframe=bt.TimeFrame.Days,
            compression=1,
        )
    else:
        # Minutes data.
        data = bt.feeds.GenericCSVData(
            dataname="data/minute_1week.csv",
            dtformat=("%Y-%m-%d %H:%M:%S"),
            timeframe=bt.TimeFrame.Minutes,
            compression=1,
        )


    cerebro.adddata(data, name="ES Mini")

    cerebro.addstrategy(Strategy)

    cerebro.addanalyzer(
        bt.analyzers.PositionsValue, headers=True, cash=True, _name="position"
    )
    cerebro.addanalyzer(bt.analyzers.Transactions, _name="transactions")
    cerebro.addanalyzer(bt.analyzers.DrawDown, _name="drawdown")
    cerebro.addanalyzer(bt.analyzers.PeriodStats, _name="period_stats")
    cerebro.addanalyzer(bt.analyzers.Returns, _name="returns", tann=252)
    cerebro.addanalyzer(bt.analyzers.TradeAnalyzer, _name="trades")
    cerebro.addanalyzer(bt.analyzers.VariabilityWeightedReturn, _name="VWR")

    # Custom Analyzers
    cerebro.addanalyzer(TradeList, _name="trade_list")
    cerebro.addanalyzer(OHLCV, _name="OHLCV")
    cerebro.addanalyzer(TradeClosed, _name="trade_closed")
    cerebro.addanalyzer(CashMarket, _name="cash_market")

    # Execute
    strat = cerebro.run(tradehistory=True)

    save_path = "results"
    save_name = "strategy_results.xlsx"
    save_to_excel(strat, save_path, save_name)



if __name__ == "__main__":

    # Timeframe minute or day
    tf = "minute"
    main(tf)