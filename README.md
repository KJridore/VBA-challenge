# VBA-challenge
In this challenge, I processed stock data, summarized it, and presented the results in a neat, user-friendly way. Here's a detailed breakdown of how the code works:

Setting Worksheet References:

I started the macro by declaring and initializing several variables, such as worksheets (wsRawData and wsSummary), row counters (LastRow), and several stock-related variables (Ticker, TotalVolume, YearlyChange, PercentChange, OpenPrice, ClosePrice, i). I set a reference to the 2020 worksheet to retrieve the raw stock data. I then tried to set a reference to an existing Summary worksheet. In the event that this worksheet didn't already exist, I took steps to create it, ensuring that there was a dedicated place for the summarized data.

Header creation:
After securing the worksheet references, my next step was to ensure that the Summary worksheet was appropriately labeled. To this end, I added appropriate headers to identify the contents of each column, viz: "Ticker Symbol", "Total Stock Volume", "Yearly Change ($)", and "Percent Change". This was to ensure that any user going through the summary would have clarity about the nature of the data being presented.

Data Processing Loop:
To extract meaningful insights from the stock data, I initialized several variables that would accumulate data related to stock volume, yearly change, and percent change for each stock ticker. To determine the scope of the loop, I identified the last row of data in the "2020" worksheet. Starting with the second row (assuming the first row contained headers), I looped through each row, meticulously extracting data such as ticker symbol, opening and closing prices, and accumulating total stock volume. Calculations were made to determine the yearly change in stock prices and the percentage change. When a new ticker symbol appeared in the dataset, that was my cue to transfer all the aggregated and calculated data for the previous ticker to the Summary worksheet. Before continuing the loop, I reset the data accumulation variables to ensure accurate data aggregation for the next ticker.

Conditional Formatting:
Finally, after all the data processing was complete, I wanted to make the summarized data visually appealing and informative at a glance. I applied conditional formatting to the Yearly Change column in the Summary worksheet. Positive annual changes were highlighted in bright green, while negative changes were highlighted in red. In addition, for clarity and precision, I made sure that the "Percent Change" column displayed the data in a percentage format, rounded to two decimal places.
