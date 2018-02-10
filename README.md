# batch-stock-scrape

## Description

This takes a list of tickers in a spreadsheet and populates that same spreadsheet with that comapny's data points (price, technical indicators, fundamental indicators, financials, financial ratios, earnings, etc). Data points are taken from a JSON api hosted at <a href="https://iextrading.com/">IEX Trading<a/>, parsed with VBA into the spreadsheet.
  
  ###### alpha.bas
  Is the current working version of the program, and the only thing you need from the this repo.
  
This script relies very heavily on a JSON parser. I would reccomend using <a>this</a> one.

Make sure you enable a Microsoft Scripting Runtime reference.

It varies in speed, the slowest I've see it go is around 3 companies per second, the fastest I've seen it go was a little over 25 companies per second.

I don't yet know why it varies so greatly in speed. But I'm working on that.

## Tickersets

The <a href="http://www.nasdaq.com/screening/company-list.aspx">NASDAQ website</a> does a great job of maintaining up-to-date CSVs containing all the tickers listed on the NYSE, NASDAQ, or AMEX.

Here's a list of direct download links by exchange:
<ul>
  <li><a href="https://www.nasdaq.com/screening/companies-by-name.aspx?letter=0&exchange=nyse&render=download">NYSE</a></li>
  <li><a href="https://www.nasdaq.com/screening/companies-by-name.aspx?letter=0&exchange=nasdaq&render=download">NASDAQ</a></li>
  <li><a href="https://www.nasdaq.com/screening/companies-by-name.aspx?letter=0&exchange=amex&render=download">AMEX</a></li>
</ul>

I also have a <a href="https://github.com/santarini/batch-stock-scrape/blob/master/sandp500.csv">S&P 500 csv</a> up here. It may or may not be up-to-date), I haven't spent a lot of time working on error handling, so if you run into a snag using this list it may because the ticker changed or doesn't exist or something.

## Errors
Some tickers have *fancy characters*, for example on the NYSE, (BAC) Bank of America has tickers like BAC^A and BAC.WS.A, in it's current state this program will likely fatal out upon encountering tickers with *fancy characters*.
