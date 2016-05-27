# Google CREST Script (GCS)
GCS is a Google Code script designed for use in Google Sheets. It provides custom functions for accessing the endpoints within EVE Online's CREST web service.

# Contents

1. [Features](#features)
1. [Setup and Configuration](#setup-and-configuration)
1. [Custom Function Use](#custom-function-use)
1. [Examples](#examples)
1. [Troubleshooting](#troubleshooting)
1. [Known Issues](#known-issues)
1. [Questions, Comments, Feedback](#questions-comments-feedback)

# Features

GCS contains the following custom functions

## Market Functions

* **getMarketPrice**: This function will access the EVE CREST market endpoint to access the real-time market data for a given item in a region at a given station.
* **getMarketHistory**: This function returns a specific column value from the historical data (as seen in-game) for an item from a given region.
* **getHistoryAdv**: This function returns all of the historical data for a given item in a given region. See [Examples](#examples) for more detail.
* **getMarketPriceList**: This function is volitile. Read [Known Issues](#known-issues) for more details. This function will accept a list of item IDs and then call getMarketPrice repeatedly to get prices for all items in the list. This is a convenience function only since CCP does not provide a multi-item endpoint for market prices at this time.
* **getOrders**: This function will return all market order data (Date Issued, Volume, Price, Location) for a given item in a given region.
* **getOrdersAdv**: This function behaves as `getOrders` does, but accepts a single 2D array of options. See [Examples](#examples) below for details.

## Universe Functions

* **getRegions**: This function will return a list of all regions in the universe along with their corresponding region ID. There are no arguments to this function.

# Setup and Configuration

## Example Spreadsheet

You can find an example spreadsheet here: [GCS Example](https://docs.google.com/spreadsheets/d/12QlphSOb-5xkukTeUlmM9I2pAFi48dXFLS5OIMA7DyY).

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

3. The orderType parameter must be `"sell"` or `"buy"` - quotes required. I leave this as an input so that you can feed in a cell value in case you want to flip the return price for an entire column or section of the sheet.
4. The refresh parameter is optional, and is in fact never used inside the function. Google caches all function returns based on the input parameter array, so if the array doesn't change then neither does the output. This is inconsistent with a function that references dynamic content, so if you need to force Google to refresh the price, put a value in this paramter. Make sure that it is different every time. Flipping back and forth from "1" to "2" and back again does not work, it must be unique. I setup a cell at the top of my market page that I increment every time I need to force a refresh. Do not use the '=now()' function, or something that changes dynamically or the script will never return a value. See [Troubleshooting](#troubleshooting) for more details.

# Examples

## getMarketPrice

Your formula should look something like this:

    =getMarketPrice(29668, 10000032, 60011866, "sell", 1)

* 29668 : whatever item ID of the product you want prices for, in this case it's PLEX
* 10000032 : the region ID for the market of interest, in this case it's Sinq Laison
* 60011866 : the station ID for the station with the market of interest, in this case it's Dodoxie
* "sell" : if you want sell orders, or "buy" if you want buy orders
* 1 : The last argument can be any value, and is used to reload the function. Read [Troubleshooting](#troubleshooting) for more details.

## getMarketPriceList

Your formula should look something like this:

    =getMarketPriceList(A1:A10, 10000032, 60011866, "sell", 1)

* A1:A10 : the range of item IDs you would like to pull prices for
* All remaining arguments are the same as the example above
 
## getOrders

Your formula should look something like this:

    =getOrders(29668, 10000032, "sell", 1)

See the getMarketPrice for a description of these parameters.

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

# Troubleshooting

## Function not updating and the reload parameter

Google will automatically reload the function every few minutes. I'm not aware of any official timer, but it seems to be between 3-5 minutes (YMMV). It is for this reason that I provide the refresh parameter in the functions. One very important note: the argument value must be deterministic. Here is what Google has to say on the matter (from their [Custom Functions documentation](https://developers.google.com/apps-script/guides/sheets/functions#arguments))

>Custom function arguments must be deterministic. That is, built-in spreadsheet functions that return a different result each time they calculate — such as NOW() or RAND() — are not allowed as arguments to a custom function. If a custom function tries to return a value based on one of these volatile built-in function, it will display Loading... indefinitely.

Some users have tried to use non-deterministic values for this argument. Keep in mind that it is there only so that you can force a reload faster than what Google already does. If you are stuck with a `Loading...` message for more than 1-2 minutes, then you have a dynamic argument in the function parameters and it must be removed.

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

In-game: Message Nuadi (CEO of Fem Bot Industries), or join the channel "Fembot Pub"

If you have a GitHub account, feel free to create issues, feature requests, etc.
