# Batch Stock Scrape

This script takes a list of tickers in a spreadsheet and populates that same spreadsheet with that company's data points (price, technical indicators, fundamental indicators, financials, financial ratios, earnings, etc). Data points are taken from the JSON API of <a href="https://iextrading.com/">IEX Trading<a/> and are then parsed with VBA into the spreadsheet.

It varies in speed, the slowest I've see it go is around 3 companies per second, the fastest I've seen it go was a little over 48 companies per second. On average I'd say it does 30 companies/sec.

I don't yet know why it varies so greatly in speed. But ... I'll spend some time on that later.

## Contents

#### alpha.bas
Is the current working version of the program and the only thing you really need from this repo.
  
This script is dependent on a JSON parser. I would reccomend using <a href="https://github.com/VBA-tools/VBA-JSON">Tim Hall's JSON Converter</a>, which I've included in this repo with his permission. See the License section below for more details.

Make sure you enable Microsoft Scripting Runtime and Microsoft Active X Data Object Library.

## Tickersets

The <a href="http://www.nasdaq.com/screening/company-list.aspx">NASDAQ website</a> does a great job of maintaining up-to-date CSVs containing all the tickers listed on the NYSE, NASDAQ, or AMEX.

Here's a list of direct download links by exchange:
<ul>
  <li><a href="https://www.nasdaq.com/screening/companies-by-name.aspx?letter=0&exchange=nyse&render=download">NYSE</a></li>
  <li><a href="https://www.nasdaq.com/screening/companies-by-name.aspx?letter=0&exchange=nasdaq&render=download">NASDAQ</a></li>
  <li><a href="https://www.nasdaq.com/screening/companies-by-name.aspx?letter=0&exchange=amex&render=download">AMEX</a></li>
  <li><a href="https://www.nasdaq.com/investing/etfs/etf-finder-results.aspx?download=Yes">ETFs</a></li>
</ul>

I also have a <a href="https://github.com/santarini/batch-stock-scrape/blob/master/sandp500.csv">S&P 500 csv</a> up here. It may or may not be up-to-date), I haven't spent a lot of time working on error handling, so if you run into a snag using this list it may because the ticker changed or doesn't exist or something.

## Errors
Some tickers have *fancy characters* -- for example on the NYSE, (BAC) Bank of America has tickers like BAC^A and BAC.WS.A -- in it's current state this program will likely fatal out upon encountering tickers with *fancy characters*. I'll put in some error handling for it when I get a chance.

## License

This project is licensed under the MIT License - see the LICENSE.md file for details

The <a href="https://github.com/santarini/batch-stock-scrape/blob/master/JsonConverter.bas">JsonConverter.bas<a/> was included in this repo with permission for the sake of simplicity. <b>JsonConverter.bas IS NOT INCLUDED UNDER MY MIT LICENSE</b>. At the top of the file I've maintained it's original license. It was taken directly from <a href="https://github.com/VBA-tools/VBA-JSON">Tim Hall</a>'s GitHub (<a href="https://github.com/VBA-tools/VBA-JSON/blob/master/JsonConverter.bas">here</a>).
