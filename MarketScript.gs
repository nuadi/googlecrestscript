// Google Crest Script (GCS)
var version = '7b'
// /u/nuadi @ Reddit
// @nuadibantine (Twitter)
//
// LICENSE: Use at your own risk, and fly safe.

// Global variable needed to track number of retries attempted
var retries = 0;
// Global variables used in order comparison function and set by
// Advanced Orders function. Default is the Price column in ascending order.
var sortIndex = 1;
var sortOrder = 1;


/**
 * Private helper function that is used to initialize the refresh token
 */
function initializeGetMarketPrice()
{
  // Test PLEX in Dodixie
  var itemId = 29668;
  var regionId = 10000032;
  var stationId = 60011866;
  var orderType = "SELL";

  var price = getMarketPrice(itemId, regionId, stationId, orderType);
  Logger.log(price);

  // Now in Jita
  regionId = 10000002;
  stationId = 60003760;

  price = getMarketPrice(itemId, regionId, stationId, orderType);
  Logger.log(price);
}


/**
 * Private helper method that performs a basic comparison of two objects.
 */
function basicCompare(object1, object2)
{
  var comparison = 0;
  if (object1[sortIndex] != null && object2[sortIndex] != null)
  {
    if (object1[sortIndex] < object2[sortIndex])
    {
      comparison = -1;
    }
    else if (object1[sortIndex] > object2[sortIndex])
    {
      comparison = 1;
    }
  }
  return comparison * sortOrder;
}


/**
 * Private helper method that wraps the UrlFetchApp in a semaphore
 * to prevent service overload.
 *
 * @param {url} url The URL to contact
 * @param {options} options The fetch options to utilize in the request
 */
function fetchUrl(url)
{
  if (gcsGetLock())
  {
    // Make the service call
    headers = {"User-Agent": "Google Crest Script version " + version + " (/u/nuadi @Reddit.com)"}
    params = {"headers": headers}
    httpResponse = UrlFetchApp.fetch(url, params);
  }
    
  return httpResponse;
}


/**
 * Custom implementation of a semaphore after LockService failed to support GCS properly.
 * Hopefully this works a bit longer...
 *
 * This function searches through N semaphores, until it finds one that is not defined.
 * Once it finds one, that n-th semaphore is set to TRUE and the function returns.
 * If no semaphore is open, the function sleeps 0.1 seconds before trying again.
 */
function gcsGetLock()
{
  var NLocks = 150;
  var lock = false;
  while (!lock)
  {
    for (var nLock = 0; nLock < NLocks; nLock++)
    {
      if (CacheService.getDocumentCache().get('GCSLock' + nLock) == null)
      {
        CacheService.getDocumentCache().put('GCSLock' + nLock, true, 1)
        lock = true;
        break;
      }
    }
  }
  return lock;
}


/**
 * Returns the average daily volume traded for a given item
 * in a given region in a given number of days.
 * @param {regionId} regionId the region to look at historical data
 * @param {itemId} itemId the item traded
 * @param {days} days the number of days to consider
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value
 * @customfunction
 */
function getAverageDailyOrders(regionId, itemId, days, refresh)
{
  var options = [
    ['days', days],
    ['headers', false],
    ['itemId', itemId],
    ['refresh', refresh],
    ['regionId', regionId]
  ];
  var historicalData = getHistoryAdv(options);

  var totalOrders = 0;
  for (var rowNumber = 0; rowNumber < historicalData.length; rowNumber++)
  {
    totalOrders += historicalData[rowNumber][2];
  }
  return totalOrders / days;
}


/**
 * Returns the average daily volume traded for a given item
 * in a given region in a given number of days.
 * @param {regionId} regionId the region to look at historical data
 * @param {itemId} itemId the item traded
 * @param {days} days the number of days to consider
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value
 * @customfunction
 */
function getAverageDailyVolume(regionId, itemId, days, refresh)
{
  var options = [
    ['days', days],
    ['headers', false],
    ['itemId', itemId],
    ['refresh', refresh],
    ['regionId', regionId]
  ];
  var historicalData = getHistoryAdv(options);

  var totalVolume = 0;
  for (var rowNumber = 0; rowNumber < historicalData.length; rowNumber++)
  {
    totalVolume += historicalData[rowNumber][1];
  }
  return totalVolume / days;
}


/**
 * Return all historical data using the options provided.
 *
 * @param {options} options see README in GitHub repo.
 * @customfunction
 */
function getHistoryAdv(options)
{
  var days = null;
  var itemId = null;
  var regionId = null;
  var showHeaders = true;

  sortIndex = 0;
  sortOrder = -1;

  if (options.length <= 0)
  {
    throw new Error("No options found");
  }
  else if (options[0] == null || options[0].length < 2)
  {
    throw new Error("Options must have 2 columns");
  }

  for (var row = 0; row < options.length; row++)
  {
    for (var col = 0; col < options[row].length; col++)
    {
      var optionKey = options[row][col];
      var optionValue = options[row][++col];
      if (optionValue === '')
      {
        continue;
      }
      else if (optionKey == 'days')
      {
        days = optionValue;
      }
      else if (optionKey == 'headers')
      {
        showHeaders = optionValue;
      }
      else if (optionKey == 'itemId')
      {
        itemId = optionValue;
      }
      else if (optionKey == 'regionId')
      {
        regionId = optionValue;
      }
      else if (optionKey == 'sortIndex')
      {
        sortIndex = optionValue;
      }
      else if (optionKey == 'sortOrder')
      {
        sortOrder = optionValue;
      }
    }
  }

  if (itemId == null)
  {
    throw new Error('No "itemId" option found');
  }
  else if (regionId == null)
  {
    throw new Error('No "regionId" option found');
  }

  var historyReturn = [];

  var headers = ['Date', 'Volume', 'Orders', 'Low', 'High', 'Average'];

  if (showHeaders == true)
  {
    historyReturn.push(headers);
  }

  var cuttoffDate = null;
  if (days != null)
  {
    var rightNow = new Date();
    var utcTimestamp = Date.UTC(rightNow.getUTCFullYear(), rightNow.getUTCMonth(), rightNow.getUTCDate());
    cuttoffDate = utcTimestamp - days*24*60*60*1000;
  }

  var historyEndpoint = "https://crest-tq.eveonline.com/market/" + regionId + "/types/" + itemId + "/history/";
  var historyJson = JSON.parse(fetchUrl(historyEndpoint));

  var historicalData = historyJson['items'];

  var history = [];
  for (var itemIndex in historicalData)
  {
    var historicalItem = historicalData[itemIndex];
    var newRow = [];
    var saveRow = true;
    for (var name in historicalItem)
    {
      var historyValue = historicalItem[name];
      if (name == 'date')
      {
        var dateValues = historyValue.split(/[T-]/);
        var dateString = dateValues.slice(0,3).join('/') + ' ' + dateValues[3];
        Logger.log('Converting date (' + historyValue + ') string: ' + dateString);
        newRow[0] = new Date(dateString);

        if (cuttoffDate != null && cuttoffDate > newRow[0].getTime())
        {
          saveRow = false;
          break;
        }
      }
      else if (name == 'volume')
      {
        newRow[1] = historyValue;
      }
      else if (name == 'orderCount')
      {
        newRow[2] = historyValue;
      }
      else if (name == 'lowPrice')
      {
        newRow[3] = historyValue;
      }
      else if (name == 'highPrice')
      {
        newRow[4] = historyValue;
      }
      else if (name == 'avgPrice')
      {
        newRow[5] = historyValue;
      }
    }

    if (saveRow)
    {
      history.push(newRow);
    }
  }

  history.sort(basicCompare);
  historyReturn = historyReturn.concat(history);
  return historyReturn;
}


/**
 * Returns yesterdays market history for a given item.
 *
 * @param {itemId} itemId the item ID of the product to look up
 * @param {regionId} regionId the region ID for the market to look up
 * @param {column} column the name of the historical column you are trying to access; "orderCount", "lowPrice", "highPrice", "avgPrice", "volume"
 * @customfunction
 */
function getMarketHistory(itemId, regionId, column)
{
  var propIndices = {
    'volume': 1,
    'orderCount': 2,
    'lowPrice': 3,
    'highPrice': 4,
    'avgPrice': 5
  };

  if (!propIndices[column] > 0)
  {
    throw new Error('Invalid column name.');
  }

  var options = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['property', propIndices[column]],
    ['headers', false]
  ];
  return getHistoryAdv(options)[0][propIndices[column]];
}


/**
 * Returns a list of all items found on the market along with their corresponding item ID.
 *
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getMarketItems(refresh)
{
  var marketItemsEndpoint = 'https://crest-tq.eveonline.com/market/types/';
  var marketItemsResponse = JSON.parse(fetchUrl(marketItemsEndpoint));

  var totalPages = marketItemsResponse['pageCount'];

  var itemList = [];
  var headers = ['Item Name', 'ID'];
  itemList.push(headers);

  for (var currentPage = 1; currentPage <= totalPages; currentPage++)
  {
    Logger.log('Processing page ' + currentPage);
    var marketItems = marketItemsResponse['items'];
    for (var itemReference in marketItems)
    {
      var item = marketItems[itemReference];
      itemList.push([item['type']['name'], item['id']]);
    }

    if (currentPage < totalPages)
    {
      var nextEndpoint = marketItemsResponse['next']['href'];
      marketItemsResponse = JSON.parse(fetchUrl(nextEndpoint));
    }
  }

  return itemList;
}


/**
 * Private helper function that will return the JSON provided for
 * a given CREST market query.
 *
 * @param {itemId} itemId the item ID of the product to look up
 * @param {regionId} regionId the region ID for the market to look up
 * @param {orderType} orderType this should be set to "sell" or "buy" orders
 */
function getMarketJson(itemId, regionId, orderType)
{
  var marketData = null;

  // Validate incoming arguments
  if (itemId == null)
  {
    throw new Error("Item ID cannot be NULL.");
  }
  else if (typeof(itemId) != "number")
  {
    throw new Error('Item ID must be a number. Instead found a(n) ' + typeof(itemId) + '.');
  }
  else if (regionId == null || typeof(regionId) != "number")
  {
    throw new Error("Invalid Region ID");
  }
  else if (orderType == null)
  {
    throw new Error("Order type cannot be NULL.");
  }
  else if (typeof(orderType) != "string")
  {
    throw new Error('Order type must be a STRING value.');
  }
  else if (orderType.toLowerCase() != 'sell' && orderType.toLowerCase() != 'buy')
  {
    throw new Error('Order type must be set to "sell" or "buy".');
  }
  else
  {
    orderType = orderType.toLowerCase();
    
    // Setup variables for the market endpoint we want
    var marketUrl = "https://crest-tq.eveonline.com/market/" + regionId + "/orders/" + orderType + "/";
    var typeUrl = "?type=https://crest-tq.eveonline.com/inventory/types/" + itemId + "/";
    Logger.log("Pulling market orders from url: " + marketUrl + typeUrl)
    
    try
    {
      // Make the call to get some market data
      marketData = JSON.parse(fetchUrl(marketUrl + typeUrl));
    }
    catch (unknownError)
    {
      Logger.log(unknownError);
      var addressError = "Address unavailable:";
      if (unknownError.message.slice(0, addressError.length) == addressError)
      {
        var maxRetries = 3;
        
        // See if we can try again
        if (retries <= maxRetries)
        {
          retries++;
          marketData = getMarketJson(itemId, regionId, orderType); 
        }
        else
        {
          marketData = "";
          for (i in unknownError)
          {
            marketData += i + ": " + unknownError[i] + "\n";
          }
        }
      }
      else
      {
        marketData = "";
        for (var i in unknownError)
        {
          marketData += i + ": " + unknownError[i] + "\n";
        }
      }
    }
  }

  return marketData;
}


/**
 * Returns the market price for a given item.
 *
 * @param {itemId} itemId the item ID of the product to look up
 * @param {regionId} regionId the region ID for the market to look up
 * @param {stationId} stationId the station ID for the market to focus on
 * @param {orderType} orderType this should be set to "sell" or "buy" orders
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value
 */
function getMarketPrice(itemId, regionId, stationId, orderType, refresh)
{
  var orderOptions = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['stationId', stationId],
    ['orderType', orderType.toLowerCase()],
    ['refresh', refresh],
    ['headers', false]
  ];

  if (orderType.toLowerCase() == 'buy')
  {
    orderOptions.push(['sortOrder', -1]);
  }
  
  var orderData = getOrdersAdv(orderOptions);

  if (orderData == null)
  {
    throw new Error('Order data came back NULL');
  }
  else if (orderData.length <= 0 || orderData[0] == null)
  {
    orderData.push(['', 0]);
  }
  
  SpreadsheetApp.flush();
  return orderData[0][1];
}


/**
 * Custom function that returns an array of prices for a given array
 * of items.
 *
 * @param {itemIdList} itemIdList the list of item IDs of the products to look up
 * @param {regionId} regionId the region ID for the market to look up
 * @param {stationId} stationId the station ID for the market to focus on
 * @param {orderType} orderType this should be set to "sell" or "buy" orders
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value
 * @customfunction
 */
function getMarketPriceList(itemIdList, regionId, stationId, orderType, refresh)
{
  var returnValues = [];

  // Only validate arguments within the context of this function
  // Further validation will occur inside getMarketPrice
  if (itemIdList == null || typeof(itemIdList) != "object")
  {
    throw new Error("Invalid Item list");
  }
  else
  {
    for (var itemIndex = 0; itemIndex < itemIdList.length; itemIndex++)
    {
      var itemId = itemIdList[itemIndex];
      if (typeof(itemId) == "object")
      {
        // This needs to be fixed before passing to getMarketPrice() function
        if (itemId.length == 1)
        {
          // This is only a number
          itemId = Number(itemId);
        }
      }
      
      // Make sure to handle blank cells accordingly
      if (itemId > 0)
      {
        returnValues[itemIndex] = getMarketPrice(itemId, regionId, stationId, orderType.toLowerCase())
      }
      else
      {
        returnValues[itemIndex] = "";
      }
    }
  }

  return returnValues;
}


/**
 * Return all market orders for an item from a region.
 *
 * @param {itemId} itemId the item ID of the product to look up
 * @param {regionId} regionId the region ID for the market to look up
 * @param {orderType} orderType this should be set to "sell" or "buy" orders
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value
 * @customfunction
 */
function getOrders(itemId, regionId, orderType, refresh)
{
  var orderOptions = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['orderType', orderType.toLowerCase()],
    ['refresh', refresh]
  ];

  if (orderType.toLowerCase() == 'buy')
  {
    orderOptions.push(['sortOrder', -1]);
  }
  
  return getOrdersAdv(orderOptions);
}


/**
 * Advanced version of getOrders.
 *
 * @param {options} options See README in GitHub repo.
 * @customfunction
 */
function getOrdersAdv(options)
{
  var itemId = null;
  var regionId = null;
  var orderType = null;
  var refresh = null;
  var showHeaders = true;
  var showOrderId = false;
  var showStationId = false;
  var stationId = null;

  var sortOrderSet = false;

  if (options.length <= 0)
  {
    throw new Error("No options found");
  }
  else if (options[0] == null || options[0].length < 2)
  {
    throw new Error("Options must have 2 columns");
  }

  for (var row = 0; row < options.length; row++)
  {
    for (var col = 0; col < options[row].length; col++)
    {
      var optionKey = options[row][col];
      var optionValue = options[row][++col];
      if (optionValue === '')
      {
        continue;
      }
      else if (optionKey == 'headers')
      {
        showHeaders = optionValue;
      }
      else if (optionKey == 'itemId')
      {
        itemId = optionValue;
      }
      else if (optionKey == 'regionId')
      {
        regionId = optionValue;
      }
      else if (optionKey == 'orderType')
      {
        orderType = optionValue.toLowerCase();
      }
      else if (optionKey == 'showOrderId')
      {
        showOrderId = optionValue;
      }
      else if (optionKey == 'showStationId')
      {
        showStationId = optionValue;
      }
      else if (optionKey == 'sortIndex')
      {
        sortIndex = optionValue;
      }
      else if (optionKey == 'sortOrder')
      {
        sortOrder = optionValue;
        sortOrderSet = true;
      }
      else if (optionKey == 'stationId')
      {
        stationId = optionValue;
      }
    }
  }

  if (itemId == null)
  {
    throw new Error('No "itemId" option found');
  }
  else if (regionId == null)
  {
    throw new Error('No "regionId" option found');
  }
  else if (orderType == null)
  {
    throw new Error('No "orderType" option found');
  }

  if (sortOrderSet == false && orderType == 'buy')
  {
    sortOrder = -1;
  }

  var headers = ['Issued', 'Price', 'Volume'];
  var locationColumn, minVolumeColumn, orderIdColumn, rangeColumn, stationIdColumn;
  if (orderType == 'buy')
  {
    minVolumeColumn = headers.length;
    headers.push('Min Volume');
    rangeColumn = headers.length;
    headers.push('Range');
  }
  locationColumn = headers.length;
  headers.push('Location');
  if (showStationId == true)
  {
    stationIdColumn = headers.length;
    headers.push('Station ID');
  }
  if (showOrderId == true)
  {
    orderIdColumn = headers.length;
    headers.push('Order ID');
  }

  var marketReturn = [];

  if (showHeaders == true)
  {
    marketReturn.push(headers);
  }
  
  // Make the call to get all market data for this item
  var jsonMarket = getMarketJson(itemId, regionId, orderType);
  var marketItems = jsonMarket['items'];
  
  // Convert all data to an array for proper output
  var outputArray = [];
  for (var rowKey in marketItems)
  {
    var saveRow = true;
    var rowData = marketItems[rowKey];
    var newRow = [];
    for (var colKey in rowData)
    {
      if (colKey == 'id' && showOrderId == true)
      {
        newRow[orderIdColumn] = rowData[colKey];
      }
      else if (colKey == 'issued')
      {
        var dateValues = rowData[colKey].split(/[T-]/);
        var dateString = dateValues.slice(0,3).join('/') + ' ' + dateValues[3];
        //Logger.log("Converting date string: " + dateString);
        newRow[0] = new Date(dateString);
      }
      else if (colKey == 'minVolume' && orderType == 'buy')
      {
        newRow[minVolumeColumn] = rowData[colKey];
      }
      else if (colKey == 'price')
      {
        newRow[1] = rowData[colKey];
      }
      else if (colKey == 'range' && orderType == 'buy')
      {
        newRow[rangeColumn] = rowData[colKey];
      }
      else if (colKey == 'volume')
      {
        newRow[2] = rowData[colKey];
      }
      else if (colKey == 'location')
      {
        var locationData = rowData[colKey];
        newRow[locationColumn] = locationData['name'];

        if (stationId != null && stationId != locationData['id'])
        {
          saveRow = false;
          break;
        }

        if (showStationId == true)
        {
          newRow[stationIdColumn] = locationData['id'];
        }
      }
    }

    if (saveRow)
    {
      outputArray.push(newRow);
    }
  }
  outputArray.sort(basicCompare);
  marketReturn = marketReturn.concat(outputArray);

  return marketReturn;
}


/**
 * Returns the market price for a given item over an entire region.
 *
 * @param {itemId} itemId the item ID of the product to look up.
 * @param {regionId} regionId the region ID for the market to look up.
 * @param {orderType} orderType this should be set to "sell" or "buy" orders.
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getRegionMarketPrice(itemId, regionId, orderType, refresh)
{
  var orderOptions = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['orderType', orderType.toLowerCase()],
    ['refresh', refresh],
    ['headers', false]
  ];

  if (orderType.toLowerCase() == 'buy')
  {
    orderOptions.push(['sortOrder', -1]);
  }
  
  var orderData = getOrdersAdv(orderOptions);

  if (orderData == null)
  {
    throw new Error('Order data came back NULL');
  }
  else if (orderData.length <= 0 || orderData[0] == null)
  {
    orderData.push(['', 0]);
  }
  
  SpreadsheetApp.flush();
  return orderData[0][1];
}


/**
 * Returns a list of all regions and their IDs.
 *
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getRegions(refresh)
{
  var regionsEndpoint = 'https://crest-tq.eveonline.com/regions/';
  var regionsData = JSON.parse(fetchUrl(regionsEndpoint));
  var regionItems = regionsData['items'];
  var regionList = [];
  var headers = ['Name', 'ID'];
  regionList.push(headers);
  for (var rowKey in regionItems)
  {
    var region = regionItems[rowKey];
    regionList.push([region['name'], region['id']]);
  }
  return regionList;
}


/**
 * Returns the market price for a given item over an entire region.
 *
 * @param {itemId} itemId the item ID of the product to look up.
 * @param {regionId} regionId the region ID for the market to look up.
 * @param {stationId} stationId the station ID for the market to look up.
 * @param {orderType} orderType this should be set to "sell" or "buy" orders.
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getStationMarketPrice(itemId, regionId, stationId, orderType, refresh)
{
  var orderOptions = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['stationId', stationId],
    ['orderType', orderType.toLowerCase()],
    ['refresh', refresh],
    ['headers', false]
  ];

  if (orderType.toLowerCase() == 'buy')
  {
    orderOptions.push(['sortOrder', -1]);
  }
  
  var orderData = getOrdersAdv(orderOptions);

  if (orderData == null)
  {
    throw new Error('Order data came back NULL');
  }
  else if (orderData.length <= 0 || orderData[0] == null)
  {
    orderData.push(['', 0]);
  }
  
  SpreadsheetApp.flush();
  return orderData[0][1];
}


/**
 * This function will run when the spreadsheet loads and perform the following:
 * 1) Request the current version from Github and inform the user of new versions
 */
function onOpen()
{
  var versionEndpoint = 'https://raw.githubusercontent.com/nuadi/googlecrestscript/master/version';
  var newVersion = fetchUrl(versionEndpoint);

  if (newVersion != null)
  {
    newVersion = newVersion.getContentText().trim();
    Logger.log('Current version from Github: ' + newVersion);

    if (newVersion > version)
    {
      var message = 'A new version of GCS is available at https://github.com/nuadi/googlecrestscript';
      var title = 'GCS version ' + newVersion + ' available!';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 120);
    }
  }
}
