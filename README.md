# Google CREST Script (GCS)
GCS is a Google Code script designed for use in Google Sheets. It will enable you to authenticate and access the Market endpoint inside EVE Online's CREST web service, and use custom functions to retrieve order data.

# BE ADVISED THIS IS CHANGING

See here: https://www.reddit.com/r/Eve/comments/3ptkw0/googlecrestscript_now_using_public_market_data/

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

* **getMarketPrice**: This function authenticates with EVE CREST to access the real-time Market price for a given item in a region at a given station.
* **getMarketHistory**: This function returns a specific column value from the historical data (as seen in-game) for an item from a given region.

# Setup and Configuration

1. First, you need to create an application through https://developers.eveonline.com. Specify the URL of your Google Sheet as the Callback URL. Make sure to select CREST Access and the publicData scope.
  - **Important:** Make sure that the Google Sheet URL that you use for the Callback URL does not contain additional parameters such as `#gid=0` or `usp=sharing`. These can interfere with the callback process in CREST. See [Troubleshooting](#troubleshooting) for details.
2. Create a Google Spreadsheet, if you don't have one already, and go to Tools > Script Editor.
3. Copy the contents of MarketScript.gs into the editor window.
4. From your application management page, copy the Client ID and Client Secret into the script. These are located on, or about, line 146 of the code in the getAuthToken() method.
5. Next you need to log into EVE SSO using a link constructed like the following, replacing YOUR_CALLBACK_URL and YOUR_CLIENT_ID with the appropriate values

  https://login-tq.eveonline.com/oauth/authorize/?response_type=code&redirect_uri=YOUR_CALLBACK_URL&client_id=YOUR_CLIENT_ID&scope=publicData&state=yarp

6. After you successfully log in, you will be redirected back to your sheet. Inside the URL look for the code parameter, which will look something like

  https://docs.google.com/spreadsheets/d/YOUR_SHEET_KEY/edit?code=LONG_AUTH_CODE&state=yarp

7. Copy the portion of LONG_AUTH_CODE into the initializeGetMarketPrice() method on line 6.
8. At the top of the Script Editor window, select the initializeGetMarketPrice function from the drop-down list (where it says Select Function), and then hit the Run button (Play button) to the left.
9. Authorize the script to make contact with external services and manage data associated with the application. These are required for the UrlFetchApp and PropertiesService, respectively.
10. It will say "Running initializeGetMarketPrice" at the top. Once this has disappeared, go to View > Logs. In it, you should see something like

  [15-02-05 09:31:32:889 EST] {expires_in=1200, token_type=Bearer, refresh_token=TOKEN_STUFF, access_token=LONG_ACCSS_TOKEN_STUFF}  

  [15-02-05 09:31:33:751 EST] 1.73835E8

  [15-02-05 09:31:34:396 EST] 1.6889999995E8

11. Copy the auth code from earlier into a cell of your spreadsheet. If you ever have to log in again, simply copy the authorization code into that cell. You should not need to run initializeGetMarketPrice after this.
11. Configuration complete.

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
4. The authCode parameter is not optional at this time, but my next effort will be to make it so that it can be omitted. For now, be sure to include a reference to the cell from step 10 of Configuration.
5. The refresh parameter is optional, and is in fact never used inside the function. Google caches all function returns based on the input parameter array, so if the array doesn't change then neither does the output. This is inconsistent with a function that references dynamic content, so if you need to force Google to refresh the price, put a value in this paramter. Make sure that it is different every time. Flipping back and forth from "1" to "2" and back again does not work, it must be unique. I setup a cell at the top of my market page that I increment every time I need to force a refresh.

# Examples

## getMarketPrice

Your formula should look something like this:

    =getMarketPrice(29668, 10000032, 60011866, "sell", A1, 1)

* 29668 : whatever item ID of the product you want prices for, in this case it's PLEX
* 10000032 : the region ID for the market of interest, in this case it's Sinq Laison
* 60011866 : the station ID for the station with the market of interest, in this case it's Dodoxie
* "sell" : if you want sell orders, or "buy" if you want buy orders
* A1 : This references a cell with your auth code in it. I do not recommend having the auth code copied into every function call, just put it in a cell and reference that.
* 1 : The last argument can be any number. Change it if you think Google isn't updating the price, which can happen sometimes.

# Troubleshooting

## Confirm that you are using publicData scope

>Request failed for https://crest-tq.eveonline.com/market/10000032/orders/sell/?type=http://crest-tq.eveonline.com/types/4247/ returned code 401. Truncated server response: {"message": "Authentication scope needed", "key": "authNeeded", "exceptionType": "UnauthorizedError"} (use muteHttpExceptions option to examine full response). (line 68, file "MarketScript")

Make sure that you have the publicData scope selected for your application in the EVE Developers web site. If you see no scopes available, then select the CREST endpoint type for the application.

## Malfunctioning callback URL

>{"error":"invalid_request", "error_description":"Some parameters are either missing or invalid"}

If your SSO URL fails on callback, it may be due to the Callback URL you provided in your application details. Make sure that the Google Sheet URL you provide does not contain any additional paramaters such as `#gid=0` or `usp=sharing`.

Here are an examples of bad URLs

>https://docs.google.com/spreadsheet/ccc?key=SHEET_KEY#gid=0

>https://docs.google.com/spreadsheet/d/SHEET_KEY/edit#gid=0

Here is an example of a good URL

>https://docs.google.com/spreadsheet/d/SHEET_KEY

# Known Issues

The Google Sheets and Scripts platform comes with it's own set of limitations. These effect the GCS in different ways, which I've listed here with possible work arounds.

## UrlFetchApp calls per day

Google limits the number of UrlFetchApp service calls that you can make per day. You can view the limits on the [Services Dashboard](https://script.google.com/dashboard) by clicking on the Quota Limits tab. This limit is on your entire Google account, so using multiple spreadsheets is not a viable workaround.

As a result of this limitation, it's not currently possible to construct a spreadsheet that requests the price of every item in EVE Online. It's recommended that you use GCS to build specific spreadsheets for tasks which benefit from pricing, but don't aim to analyze the entire market. I will work to develop methods to get around this limitation.

# Questions, Comments, Feedback

If you have any questions on how to use the function, comments on what could be improved, or general feedback on the script, this README, or bugs - Contact me

Reddit (best method): /u/nuadi

In-game: Nuadi, CEO of Fem Bot Industries

Twitter: @nuadibantine

If you have a GitHub account, feel free to create issues, feature requests, etc.
