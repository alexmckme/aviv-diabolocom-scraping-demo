# Diabolocom Scraping and simple-salesforce demo

This project is a showcase of a script that is currently used at AVIV Group to address the needs of some internal clients, mainly the ability to supervise the main KPIs of customer support agents in France.

While the main repo is located on another Github account, the code is shown here for demonstration purposes.

At the end, the following file is updated regularly:

![Results File Screenshot](/images/Results-File-Example.png/?raw=true)

## How It's Made:

**Tech used: Python**

I used the following packages:
- **Selenium** to automate manual extracts downloads from 2 different Diabolocom websites
- **simple-salesforce** to fetch data from 2 different Salesforce organizations
- **pandas, gspread, gspread-dataframe** to put the data to a Google Sheets file (for the internal clients own comfort)
- **Github Actions** to schedule the execution of the script with cron

## Optimizations

- The script takes about 3 minutes to run on average. Decreasing it will allow to save usage minutes, which are limited to 2,000 per month on Github free plan
- Make API requests instead of using Selenium to "scrape" Diabolocom. At the time of writing, I am not allowed to have an API key for this purpose. This script will be overhauled once AVIV security team allows distributing API keys.
- Improve error handling.

## Lessons Learned:

With this simple script, I successfully addressed a real life issue that some of my internal clients faced. They used to have to do manual exports and copy-paste to some badly formatted Google Sheets file every evening for years, without exception, with little to now way of automating this themselves.

With some knowledge of Python, and pandas for data analysis, I dwelved deeper into leveraging simple-salesforce to make requests to Salesforce API, Selenium to scrape websites as using their public API was not an option, and Github Actions to automate its execution.


