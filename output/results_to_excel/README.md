## Results to Excel
Bactrader addon that saves results to excel spreadsheets.
### Describe
This addon uses the standard built in analyzers in backtrader, as well as some custom analyzers and saves these to 
excel. You can save results for a single backtest at a time. Does not work for optimization. 
### What it does
Results to Excel takes the analyzer outputs from a strategy object, and re-formats the information into friendly columns 
and rows as well as some number and decimal formatting, and saves each analyzer to a sheet in an excel workbook.
### How it works
The first step is to add the file `result_to_excel.py` your project and import the function 'save_to_excel'.
```
from result_to_excel import save_to_excel
``` 
Import any custom analyzers you want to use. 
```
from analyzer import TradeList, OHLCV, TradeClosed, CashMarket
```
Then add to cerebro the analyzers you wish to use. You can name the analyzer anything you wish, but it is critical
that the class names remain unchanged, as this is how the result_to_excel addon identifies the analyzer.
```
cerebro.addanalyzer(bt.analyzers.PositionsValue, headers=True, cash=True, _name="position")
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
```
To run the strategy and save the spreadsheets, you must set cerebro to a variable. 
(```tradehistory=True``` if using the trade_list analyzer.)
And also you must set the directory and filename you wish to use to save the spreadsheet.
```
strat = cerebro.run(tradehistory=True)

save_path = "results"
save_name = "strategy_results.xlsx"
save_to_excel(strat, save_path, save_name)
```
There is a sample_strategy.py file included to try it out.

### Created by
Neil Murphy a.k.a. run-out

### License: 
Same license as current backtrader license: GNU GENERAL PUBLIC LICENSE, Version 3, 29 June 2007
