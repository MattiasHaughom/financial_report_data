# financial_report_data
Get financial data from pdf reports on different Norwegian company websites

## Data
The script relies on a list of tickers taken from Dagens Næringsliv's norwegian listed (OBX) stock overview. Othervise one needs to run the financial calendar script to get the calendar days for the different reports, which is used in the report download script.

## Scripts
The getLinksFromOsloBørs.py script accesses the reports released on that day, for public companies, downloads the report, and attempts to extract the income statement from the report. The resulting dataframe of data is saved in a dedicated company folder in your path location. You would need to change this.

## Improvements
I want to develop the accounting and fundamental analysis of the income statement. Maybe paired with other tables in the report (balance sheet, cash flow etc.) to create models predicting or forecasting price and growth. This is extremely difficult. Analysis and output depends on industry, company and special cases. The data differs in their names (profit for the period, net profit, operating profit etc...), their currency and abbreviation. The main benefit of doing this yourself is the possibility of getting data at speed and for free. However, even building a extraction model for each company won't work if reports change for each quarter. 

I also want to implement a sentiment analysis of the event text, which is released along with the report/presentation, using for example the loughran and McDonald dictionary or similar to estimate postive or negative sentiment. This might not be informative because even hard losses are communicated positively, but we will see.

The script obviously needs a lot more tweaks to work for all reports, some are causing issues for tabluar and one idea could be to implement alternatives in these cases. This is an iterative processes.
