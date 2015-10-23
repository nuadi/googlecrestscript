# Google CREST Script (GCS)
GCS is a Google Code script designed for use in Google Sheets. It will enable you to authenticate and access the Market endpoint inside EVE Online's CREST web service, and use custom functions to retrieve order data.

# Contents

1. [Setup and Configuration](#setup-and-configuration)
2. [Features](#features)
2. [Custom Function Use](#custom-function-use)
5. [Examples](#examples)
3. [Troubleshooting](#troubleshooting)
4. [Known Issues](#known-issues)
4. [Questions, Comments, Feedback](#questions-comments-feedback)

# Features

GCS contains the following custom functions

* **getMarketPrice**: This function will access the EVE CREST market endpoint to access the real-time market data for a given item in a region at a given station.
* **getMarketHistory**: This function returns a specific column value from the historical data (as seen in-game) for an item from a given region.
* **getMarketPriceList**: This function is volitile. Read [Known Issues](#known-issues) for more details. This function will accept a list of item IDs and then call getMarketPrice repeatedly to get prices for all items in the list. This is a convenience function only since CCP does not provide a multi-item endpoint for market prices at this time.

# Setup and Configuration

1. Create a Google Spreadsheet, if you don't have one already, and go to Tools > Script Editor.
2. Copy the contents of MarketScript.gs into the editor window.
3. At the top of the Script Editor window, select the initializeGetMarketPrice function from the drop-down list (where it says Select Function), and then hit the Run button (Play button) to the left.
4. Authorize the script to make contact with external services and manage data associated with the application. These are required for the UrlFetchApp.
5. It will say "Running initializeGetMarketPrice" at the top. Once this has disappeared, go to View > Logs. In it, you should see something like

  [15-02-05 09:31:33:751 EST] 1.73835E8

  [15-02-05 09:31:34:396 EST] 1.6889999995E8

6. These are market prices for the items in the initializeGetMarketPrice function. As long as you see these, setup is complete.

# Custom Function Use

The getMarketPrice() method is built as a Custom Function for use in cells of Google Sheets. The JavaDoc comment is setup so that you can see what each parameter is used for as you type in the function inside a cell. Here are some tips to keep in mind.

The return value is the price for that item at a given station based on the order type that you want. If the value is "sell" then you will be given the lowest sell price for the product. If it is "buy" then you will get the highest buy order price for the product. If no orders can be found at that station, then the return price is 0.

1. You can get the Item ID using VLookup and a sheet dedicated to hold all items and their ID. This comes from the invType table of the Static Data Export (SDE). You can get this data in XLS format from a really nice guy over at https://www.fuzzwork.co.uk/dump/latest/
2. You can get the Region and Station IDs using the Map endpoint in CREST, no authorization requried. I use my browser window to find what I need. For convenience, here are the four major hub regions and stations

  Domain ID	10000043

  Heimatar ID	10000030

  Sinq Laison ID	10000032

  The Forge ID	10000002

  Amarr station ID	60008494

  Dodixie station ID	60011866

  Jita station ID	60003760

  Rens ID	60004588

3. The orderType parameter must be "sell" or "buy". I leave this as an input so that you can feed in a cell value in case you want to flip the return price for an entire column or section of the sheet.
4. The refresh parameter is optional, and is in fact never used inside the function. Google caches all function returns based on the input parameter array, so if the array doesn't change then neither does the output. This is inconsistent with a function that references dynamic content, so if you need to force Google to refresh the price, put a value in this paramter. Make sure that it is different every time. Flipping back and forth from "1" to "2" and back again does not work, it must be unique. I setup a cell at the top of my market page that I increment every time I need to force a refresh. Do not use the '=now()' function, or something that changes dynamically or the script will never return a value.

# Examples

## getMarketPrice

Your formula should look something like this:

    =getMarketPrice(29668, 10000032, 60011866, "sell", 1)

* 29668 : whatever item ID of the product you want prices for, in this case it's PLEX
* 10000032 : the region ID for the market of interest, in this case it's Sinq Laison
* 60011866 : the station ID for the station with the market of interest, in this case it's Dodoxie
* "sell" : if you want sell orders, or "buy" if you want buy orders
* 1 : The last argument can be any value. Change it if you think Google isn't updating the price, which can happen sometimes.

# Troubleshooting

After authenticated endpoints were removed, there is presently nothing to really troubleshoot. If you hit something, contact me.

# Known Issues

The Google Sheets and Scripts platform comes with it's own set of limitations. These effect the GCS in different ways, which I've listed here with possible work arounds.

## UrlFetchApp calls per day

Google limits the number of UrlFetchApp service calls that you can make per day. You can view the limits on the [Services Dashboard](https://script.google.com/dashboard) by clicking on the Quota Limits tab. This limit is on your entire Google account, so using multiple spreadsheets is not a viable workaround.

As a result of this limitation, it's not currently possible to construct a spreadsheet that requests the price of every item in EVE Online. It's recommended that you use GCS to build specific spreadsheets for tasks which benefit from pricing, but don't aim to analyze the entire market. If you simply cannot avoid this limitation with your spreadsheet, then you will either need to build a program yourself, or find someone who has a program to perform the analysis you need.

## The Multi-Pull function and Google's limits

Google limits all custom functions to 30 seconds or less. If a custom function exceeds this time, an '#ERROR' type is returned indicating as such. Since the multi-pull function is only making individual calls for a list of items, you may run into this limit from time to time. How many items you can put into an item ID list will depend on a number of factors including: the CREST server response time to Google, Google's server performance, the amount of data to sort through for each item.

If you hit the limit, reduce your item ID list by half, and observe behavior. If it continues, repeat the cleave until you're stable. If it stabilizes, then now you have a min-max range for your particular case, and you can continue to tinker within that range where you feel comfortable.

# Questions, Comments, Feedback

If you have any questions on how to use the function, comments on what could be improved, or general feedback on the script, this README, or bugs - Contact me

Reddit (best method): /u/nuadi

Twitter: @nuadibantine

If you have a GitHub account, feel free to create issues, feature requests, etc.
