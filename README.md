# MacroDashBoard
## Workflow
### MacroDashboard-GitDeploy
+ Includes all working code from fetching data to updating the excel file
+ Divided into 4 main parts
    + Fetch data according to tickers from FRED database
    + Store data in csv files according to their frequency
    + Using data in csv files to update values in the MacroDashboard
### Raw Data
+ include csv files of the FRED data grouped by frequency
+ Values include value of the FRED data
+ Dates record the most recent date of the data
### MacroDashboard Versions
+ keep newest and past MacroDashboards
### Code Test
+ jupyter notebook on which I test my code
## Technical Problems
+ sometimes there will be missing date / nonconsecutive time, should I ignore it or add NaN to fill the gap? (likely the data is not released on weekands)
