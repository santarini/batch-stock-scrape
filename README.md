# batch-stock-scrape

This takes a list of tickers in a spreadsheet and populates a spreadsheet with that companies data points.

Make sure you enable a Microsoft Scripting Runtime reference.

It varies in speed, the slowest I've see it go is around 3 companies per second, the fastest I've seen it go was a little 25 companies per second.

I don't yet know why it varies so greatly in speed. Working on that.

## Tickersets

The <a href="http://www.nasdaq.com/screening/company-list.aspx">NASDAQ website</a> does a great job of maintaining up-to-date CSVs containing all the tickers listed on the NYSE, NASDAQ, or AMEX.

Here's a list of direct download links by exchange:
<ul>
  <li><a href="http://www.nasdaq.com/screening/companies-by-industry.aspx?exchange=NYSE">NYSE</a></li>
  <li><a href="http://www.nasdaq.com/screening/companies-by-industry.aspx?exchange=NASDAQ">NASDAQ</a></li>
  <li><a href="http://www.nasdaq.com/screening/companies-by-industry.aspx?exchange=AMEX">AMEX</a></li>
</ul>

I also have an S & P 500 csv up here. (it may or may not be up-to-date).
