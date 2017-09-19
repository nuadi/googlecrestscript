# Google CREST/ESI Script (GCES)
GCES is a Google Code script designed for use in Google Sheets. It provides custom functions for accessing the endpoints within EVE Online's CREST/ESI web service.

This entire script is slated for migration to use CCP's ESI endpoints. If you see a `_beta.gs` script in the repo, please take a moment to copy your sheet and try it out. The beta scripts are being used to migrate functions and test them before finalizing the version for public consumption.

# Contents

1. [Features](#features)
1. [Pre-builts](#pre-builts)
1. [Setup and Configuration](#setup-and-configuration)
1. [Custom Function Use](#custom-function-use)
1. [Examples](#examples)
1. [FAQ](#faq)
1. [Troubleshooting](#troubleshooting)
1. [Known Issues](#known-issues)
1. [Questions, Comments, Feedback](#questions-comments-feedback)

# Features

GCES contains the following custom functions

## Market Price Functions

* **getRegionMarketPrice**: This function will access the EVE CREST market endpoint to access the real-time market data for a given item in a region.
* **getStationMarketPrice**: This is like `getRegionMarketPrice` but will filter down to a specific station.
* ~~**getMarketPriceList**: This function is volitile. Read [Known Issues](#known-issues) for more details. This function will accept a list of item IDs and then call getMarketPrice repeatedly to get prices for all items in the list. This is a convenience function only since CCP does not provide a multi-item endpoint for market prices at this time.~~ This function is being removed with the 12 series of scripts. If you insist on keeping it, you can copy it from a previous version into the 12 series, or beyond. However, I will no longer support or maintain this function until CCP gives us a proper list-based endpoint.
* **getOrders**: This function will return all market order data (Date Issued, Volume, Price, Location) for a given item in a given region.
* **getOrdersAdv**: This function behaves as `getOrders` does, but accepts a single 2D array of options. See [Examples](#examples) below for details.
* **countStationOrders**: This function returns the number of available orders for a given item at a station.
* **countStationVolume**: This function sums up the total number of available units for a given item at a station.

## Market History Functions

* **getMarketHistory**: This function returns a specific column value from the historical data (as seen in-game) for an item from a given region.
* **getHistoryAdv**: This function returns all of the historical data for a given item in a given region. See [Examples](#examples) for more detail.
* **getAverageDailyOrders**: This is a convenience function that will calculate the average daily number or orders observed in a given number of days in the historical data.
* **getAverageDailyVolume**: This is a convenience function that will calculate the average daily trade volume observed in a given number of days in the historical data.

## Market Utility Functions

* **getMarketGroups**: This function will return a lit of all top-level market groups (what you see on the left of the in-game market), or a list of child groups if you specify a group ID.
* **getMarketGroupItems**: This function will return a list of all items found in a given market group, including child groups.
* **getMarketItems**: This function will return a list of all items found on the open market along with their corresponding item ID. This function has an optional refresh argument. **Warning:** This function has no direct migration path in the current form of ESI, so it will die with CREST.

## NPC Corporations and LP Stores
* **getNPCCorporations**: This function will return a list of all NPC corporations and their corp ID for use in the LP store function.
* **getNPCLoyaltyStore**: This function will return a matrix with all available NPC Loyalty Store items, their costs, and required items for a given corp ID.

## Industry Functions
* **getAdjustedPrice**: Returns the adjusted price for an item used in industrial calculations.
* **getCostIndex**: Returns the cost index for a given activity in a given solar system.

## Universe Functions

* **getRegions**: This function will return a list of all regions in the universe along with their corresponding region ID. There are no arguments to this function.

## Utility Functions

* **getItemVolume**: Returns the volume for a given item. Note: This is the assembled volume. There is no way to return the packaged volume.

## Depricated Functions

These functions will be removed in a future version of GCES. Please update your sheets to use the replacement functions listed.

* **getMarketPrice**: This function will access the EVE CREST/ESI market endpoint to access the real-time market data for a given item in a region at a given station. **REPLACEMENT:** Replace all uses of this function with `getStationMarketPrice` to maintain functionality.

# Pre-builts

* **Station Trading Sheet**: [[link](https://docs.google.com/spreadsheets/d/17Sl_YkR7J8AcYX05BrwsBNDP_AjGJTrESjfzVFUtfWo)] This sheet is built to support a station trader who buys/sells products in a single station. Be sure to read the README sheet for instructions on how to use and/or modify the sheet to fit your character. Let me know if you have any questions or need help.
* **LP Store**: [[link](https://docs.google.com/spreadsheets/d/1kNuNNus_j1fwIkSm62ADkRdoWmntmGhW5_ndpSJsqBI)] This sheet is build to output an LP Store and then calculate the ISK/LP ratio for all items. Blueprint prices cannot be pulled from CREST, so these prices must be updated manually. Otherwise it handles everything else. Duplicate the LP Store sheet to change the store and then update the Market sheet with the items and prices needed.

# Setup and Configuration

## Example Spreadsheet

You can find an example spreadsheet here: [GCES Example](https://docs.google.com/spreadsheets/d/12QlphSOb-5xkukTeUlmM9I2pAFi48dXFLS5OIMA7DyY).

Once you have it open, got to `File -> Make a Copy...` and then name the new file what ever you like. This spreadsheet contains examples of how to use all of the available functions, along with the script already copied into the editor.

## Initialization

1. Create a Google Spreadsheet, or copy the example above, and then go to Tools > Script Editor.
2. Copy the contents of MarketScript.gs into the editor window and save the script. You may need to name the project to save it, so use any name.
3. At the top of the Script Editor window, select the initializeGetMarketPrice function from the drop-down list (where it says Select Function), and then hit the Run button (Play button) to the left.
4. Authorize the script to make contact with external services. This are required for the UrlFetchApp.
5. It will say "Running initializeGetMarketPrice" at the top. Once this has disappeared, go to View > Logs. In it, you should see something like

  [15-02-05 09:31:33:751 EST] 1.73835E8

  [15-02-05 09:31:34:396 EST] 1.6889999995E8

6. These are market prices for the items in the initializeGetMarketPrice function. As long as you see these, setup is complete.

# Custom Function Use

The getStationMarketPrice() method is built as a Custom Function for use in cells of Google Sheets. The JavaDoc comment is setup so that you can see what each parameter is used for as you type in the function inside a cell. Here are some tips to keep in mind.

The return value is the price for that item at a given station based on the order type that you want. If the value is "sell" then you will be given the lowest sell price for the product. If it is "buy" then you will get the highest buy order price for the product. If no orders can be found at that station, then the return price is 0.

1. You can get the Item ID using VLookup and a sheet dedicated to hold all items and their ID. This comes from the invType table of the Static Data Export (SDE). You can get this data in XLS format from a really nice guy over at https://www.fuzzwork.co.uk/dump/latest/
2. You can get the Region and Station IDs using the Map endpoint in CREST/ESI, no authorization requried. I use my browser window to find what I need. For convenience, here are the four major hub regions and stations

  Domain ID	10000043

  Heimatar ID	10000030

  Sinq Laison ID	10000032

  The Forge ID	10000002

  Amarr station ID	60008494

  Dodixie station ID	60011866

  Jita station ID	60003760

  Rens ID	60004588

3. The orderType parameter must be `"sell"` or `"buy"` - quotes required. I leave this as an input so that you can feed in a cell value in case you want to flip the return price for an entire column or section of the sheet.
4. The refresh parameter is optional, and is in fact never used inside the function. Google caches all function returns based on the input parameter array, so if the array doesn't change then neither does the output. This is inconsistent with a function that references dynamic content, so if you need to force Google to refresh the price, put a value in this paramter. Make sure that it is different every time. Flipping back and forth from "1" to "2" and back again does not work, it must be unique. I setup a cell at the top of my market page that I increment every time I need to force a refresh. Do not use the '=now()' function, or something that changes dynamically or the script will never return a value. See [Troubleshooting](#troubleshooting) for more details.

# Examples

## getStationMarketPrice

Your formula should look something like this:

    =getStationMarketPrice(29668, 10000032, 60011866, "sell", 1)

* 29668 : whatever item ID of the product you want prices for, in this case it's PLEX
* 10000032 : the region ID for the market of interest, in this case it's Sinq Laison
* 60011866 : the station ID for the station with the market of interest, in this case it's Dodoxie
* "sell" : if you want sell orders, or "buy" if you want buy orders
* 1 : The last argument can be any value, and is used to reload the function. Read [Troubleshooting](#troubleshooting) for more details.

## getMarketPriceList

**WARNING:** This function is being removed in future versions of GCES. You are welcome to keep the function in your scripts, but I will no longer support or maintain compatability after the 12 series script is final.

Your formula should look something like this:

    =getMarketPriceList(A1:A10, 10000032, 60011866, "sell", 1)

* A1:A10 : the range of item IDs you would like to pull prices for
* All remaining arguments are the same as the example above
 
## getOrders

Your formula should look something like this:

    =getOrders(29668, 10000032, "sell", 1)

See the getStationMarketPrice for a description of these parameters.

This function will return a 2D array 4 columns wide and an unknown number of rows high which will populate a sheet starting with the cell you called the function in, and then proceed to the right and down the sheet. The four columns have headers, and are titled Issued, Price, Volume, and Location. The rows are automatically sorted by price, lowest to highest.

If any cell would be overwritten, the function will fail with a #REF! error, so be sure the formula has the space it needs. If there are not enough columns to the right or rows beneath the function cell, Google Sheets will expand your sheet for you. However, the number of columns and rows added may far exceed what the function needs, so if you care about sheet size please keep this in mind.

## getOrdersAdv

This function requires a 2D array of options that is 2 columns wide and at least 3 rows high. The options can be in any order. Setup a range of cells like the following

| | A | B
|---|:--|--:
|1| itemId | 29668
|2| regionId | 10000002
|3| orderType | sell
|4| sortIndex | 1
|5| sortOrder | 1
|6| Refresh | y

and your function call will look like this

    =getOrdersAdv(A1:B6)

The available options are shown in the table below. All option keys are case sensitive.

| Option Key | Required? | Description
|:--|---|:--
| itemId | Yes | The ID of the item to look for. Must be a number value.
| regionId | Yes | The region ID to look in. Must be a number value.
| orderType | Yes | The order types to return. Must be "sell" or "buy".
| headers | no | Set to FALSE to hide the headers. Default is TRUE.
| showOrderId | no | Set to TRUE, or 1, to add a new column named "Order ID" that will contain each order's ID.
| showStationId | no | Set to TRUE, or 1, to add a new column named "Station ID" that will contain the ID of the station the order is placed within.
| sortIndex | no | Numver value for the column to sort. 0 = Location, 1 = Price (default), 2 = Volume, 3 = Location, and so on.
| sortOrder | no | Number value for sort order. 1 = Normal order (default for sell), -1 = Reverse order (default for buy)
| stationId | no | Specify a specific station ID to show only orders from that station. Only 1 value supported at this time.
| refresh | no | Same as all refresh parameters. Never used. Can be any deterministic value.

## getHistoryAdv

This function requires a 2D array of options that is 2 columns wide and at least 2 rows high. The options can be in any order. Setup a range of cells like the following

| | A | B
|---|:--|--:
|1| itemId | 29668
|2| regionId | 10000002
|3| sortIndex | 0
|4| sortOrder | -1

and your function call will look like this

    =getHistoryAdv(A1:B4)

The available options are shown in the table below. All option keys are case sensitive.

| Option Key | Required? | Description
|:--|---|:--
| itemId | Yes | The ID of the item to look for. Must be a number value.
| regionId | Yes | The region ID to look in. Must be a number value.
| days | no | Only return a given number of days. Default will return all historical data.
| headers | no | Set to FALSE to hide the headers. Default is TRUE.
| sortIndex | no | Number value for the column to sort. 0 = Location, 1 = Price (default), 2 = Volume, 3 = Location, and so on.
| sortOrder | no | Number value for sort order. 1 = Normal order (default for sell), -1 = Reverse order (default for buy)
| refresh | no | Same as all refresh parameters. Never used. Can be any deterministic value.

## Filters

Many of the market functions support an optional filters parameter. This parameter accepts a 2D array of key-value pairs in the same style as getOrdersAdv above. The filters can be passed directly into getOrdersAdv. The available options are

| Filter Key | Description
|---|:--
|maxPrice| This will filter all orders above a given price
|minPrice| This will filter all orders below a given price
|maxVolume| This will filter all orders above a given volume
|minVolume| This will filter all orders below a given volume

# FAQ

I get a lot of questions about this script, so I felt it was time to build this section. If your question does not appear here then by all means [contact me](#questions-comments-feedback).

## Do the functions automatically update over time?

No, and yes. First off, nothing in this script automatically updates. However, Google *does* update all cells in a workbook periodically. I am not aware of the exact timing, but I believe it is somewhere in the neighborhood of 15-30 minutes as long as you keep the workbook open. I do not think it updates automatically when closed, but it Google will try to call the custom functions when you open the sheet again.

Therefore, if you need the scripts automatically updated you have a choice of leaving the workbook open all the time (as I do), or periodically opening the workbook yourself to force a workbook-wide update. Be warned: You can inadvertantly hit your URL Fetch quota in both cases.

## Can you get the history for a market item at a station?

No. CCP does not provide this high of a resolution for historical data. What you see in-game in the historical data is exactly what you get through the CREST/ESI API.

If you truly need station-specific historical data, then you will have to scrape market data yourself, store it, and then figure out a way to determine the historical data through your own heuristics.

# Troubleshooting

## Function not updating and the reload parameter

Google will automatically reload the function every few minutes. I'm not aware of any official timer, but it seems to be between 3-5 minutes (YMMV). It is for this reason that I provide the refresh parameter in the functions. One very important note: the argument value must be deterministic. Here is what Google has to say on the matter (from their [Custom Functions documentation](https://developers.google.com/apps-script/guides/sheets/functions#arguments))

>Custom function arguments must be deterministic. That is, built-in spreadsheet functions that return a different result each time they calculate — such as NOW() or RAND() — are not allowed as arguments to a custom function. If a custom function tries to return a value based on one of these volatile built-in function, it will display Loading... indefinitely.

Some users have tried to use non-deterministic values for this argument. Keep in mind that it is there only so that you can force a reload faster than what Google already does. If you are stuck with a `Loading...` message for more than 1-2 minutes, then you have a dynamic argument in the function parameters and it must be removed.

# Known Issues

The Google Sheets and Scripts platform comes with it's own set of limitations. These effect the GCES in different ways, which I've listed here with possible work arounds.

## UrlFetchApp calls per day

Google limits the number of UrlFetchApp service calls that you can make per day. You can view the limits on the [Services Dashboard](https://script.google.com/dashboard) by clicking on the Quota Limits tab. This limit is on your entire Google account, so using multiple spreadsheets is not a viable workaround.

As a result of this limitation, it's not currently possible to construct a spreadsheet that requests the price of every item in EVE Online. It's recommended that you use GCES to build specific spreadsheets for tasks which benefit from pricing, but don't aim to analyze the entire market. If you simply cannot avoid this limitation with your spreadsheet, then you will either need to build a program yourself, or find someone who has a program to perform the analysis you need.

## The Multi-Pull function and Google's limits

~~Google limits all custom functions to 30 seconds or less. If a custom function exceeds this time, an '#ERROR' type is returned indicating as such. Since the multi-pull function is only making individual calls for a list of items, you may run into this limit from time to time. How many items you can put into an item ID list will depend on a number of factors including: the CREST server response time to Google, Google's server performance, the amount of data to sort through for each item.

If you hit the limit, reduce your item ID list by half, and observe behavior. If it continues, repeat the cleave until you're stable. If it stabilizes, then now you have a min-max range for your particular case, and you can continue to tinker within that range where you feel comfortable.~~

This section describes a function that has been removed. This section will remain until the 13-series is released, at which point it will be removed.

# Questions, Comments, Feedback

If you have any questions on how to use the functions, comments on what could be improved, general feedback on the script, this README, or bugs - do not hesitate to contact me. I also help in custom implementations if your corp has some workflow that doesn't quite fit into the script's feature set.

Email: nuadi.bantine@gmail.com

If you have a GitHub account, feel free to create issues, feature requests, etc.
