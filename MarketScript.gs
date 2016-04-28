// Google Crest Script (GCS)
// version 4h
// /u/nuadi @ Reddit
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
 * Private helper method that will compare two market orders.
 */
function compareOrders(order1, order2)
{
  var comparison = 0;
  if (order1[sortIndex] < order2[sortIndex])
  {
    comparison = -1;
  }
  else if (order1[sortIndex] > order2[sortIndex])
  {
    comparison = 1;
  }
  return comparison * sortOrder;
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
  if (itemId == null || typeof(itemId) != "number")
  {
    throw new Error("Invalid Item ID");
  }
  else if (regionId == null || typeof(regionId) != "number")
  {
    throw new Error("Invalid Region ID");
  }
  else if (orderType == null || typeof(orderType) != "string" || orderType.toLowerCase() != 'sell' && orderType.toLowerCase() != 'buy')
  {
    throw new Error("Invalid order type");
  }
  else
  {
    orderType = orderType.toLowerCase();
    
    // Setup variables for the market endpoint we want
    var marketUrl = "https://crest-tq.eveonline.com/market/" + regionId + "/orders/" + orderType + "/";
    var typeUrl = "?type=https://crest-tq.eveonline.com/types/" + itemId + "/";
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
 * @customfunction
 */
function getMarketPrice(itemId, regionId, stationId, orderType, refresh)
{
  var returnPrice = 0;

  if (stationId == null || typeof(stationId) != "number")
  {
    throw new Error("Invalid Station ID");
  }
  else
  {
    var jsonMarket = getMarketJson(itemId, regionId, orderType);
    if (typeof(jsonMarket) == "string")
    {
      returnPrice = jsonMarket;
    }
    else if (jsonMarket != null)
    {
      returnPrice = getPrice(jsonMarket, stationId, orderType)
    }
  }

  SpreadsheetApp.flush();
  return returnPrice;
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
        returnValues[itemIndex] = getMarketPrice(itemId, regionId, stationId, orderType)
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
    ['orderType', orderType],
    ['refresh', refresh]
  ];

  if (orderType == 'buy')
  {
    orderOptions.push(['sortOrder', -1]);
  }
  
  return getOrdersAdv(orderOptions);
}

/**
 * Advanced version of getOrders.
 *
 * @param {options} See README in GitHub repo.
 * @customfunction
 */
function getOrdersAdv(options)
{
  var itemId = null;
  var regionId = null;
  var orderType = null;
  var refresh = null;
  var showOrderId = false;
  var showStationId = false;
  var stationId = null;

  var sortOrderSet = false;

  if (options.length <= 0)
  {
    throw new Error("No options found");
  }
  else if (options[0].length < 2)
  {
    throw new Error("Options must have 2 columns");
  }

  for (var row = 0; row < options.length; row++)
  {
    for (var col = 0; col < options[row].length; col++)
    {
      var optionKey = options[row][col];
      var optionValue = options[row][++col];
      if (optionValue == '')
      {
        continue;
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
        orderType = optionValue;
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
  marketReturn.push(headers);
  
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
        Logger.log("Converting date string: " + dateString);
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
  outputArray.sort(compareOrders);
  marketReturn = marketReturn.concat(outputArray);

  return marketReturn;
}

/**
 * Private helper method that will determine the best price for a given item from the
 * market data provided.
 *
 * @param {jsonMarket} jsonMarket the market data in JSON format
 * @param {stationId} stationId the station ID to focus the search on
 * @param {orderType} orderType the type of order is either "sell" or "buy"
 */
function getPrice(jsonMarket, stationId, orderType)
{
  var bestPrice = 0;

  // Pull all orders found and start iteration
  var orders = jsonMarket['items'];
  for (var orderIndex = 0; orderIndex < orders.length; orderIndex++)
  {
    var order = orders[orderIndex];
    if (stationId == order['location']['id'])
    {
      // This is the station market we want
      var price = order['price'];

      if (bestPrice > 0)
      {
        // We have a price from a previous iteration
        if (orderType == "sell" && price < bestPrice)
        {
          bestPrice = price;
        }
        else if (orderType == "buy" && price > bestPrice)
        {
          bestPrice = price;
        }
      }
      else
      {
        // This is the first price found, take it
        bestPrice = price;
      }
    }
  }

  return bestPrice;
}

/**
 * Returns yesterdays market history for a given item.
 *
 * @param {itemId} itemId the item ID of the product to look up
 * @param {regionId} regionId the region ID for the market to look up
 * @param {property} property the property you are trying to access; "orderCount", "lowPrice", "highPrice", "avgPrice", "volume"
 * @customfunction
 */
function getMarketHistory(itemId, regionId, property)
{
  var history = 0;
  
  // Validate incoming arguments
  if (itemId == null || typeof(itemId) != "number")
  {
    throw new Error("Invalid Item ID");
  }
  else if (regionId == null || typeof(regionId) != "number")
  {
    throw new Error("Invalid Region ID");
  }
  else if (property != "orderCount" && property != "lowPrice" && property != "highPrice" && property != "avgPrice" && property != "volume")
  {
    throw new Error("Property must be one of: 'orderCount', 'lowPrice', 'highPrice', 'avgPrice', 'volume'");
  }
  else
  { 
    // Setup variables for the market endpoint we want
    var marketUrl = "https://crest-tq.eveonline.com/market/" + regionId + "/types/" + itemId + "/history/"
  
    // Make the call to get some market data
    var jsonMarket = JSON.parse(fetchUrl(marketUrl));
  
    // Get the desired property
    var items = jsonMarket['items'];
    var history = items[items.length - 1][property]
  }
  return history;
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
  var lock = LockService.getUserLock().tryLock(1000/150);
  // Make the service call
  headers = {"User-Agent": "Google Crest Script version 4b (/u/nuadi @Reddit.com)"}
  params = {"headers": headers}
  httpResponse = UrlFetchApp.fetch(url, params);
  
  if (lock)
    LockService.getUserLock().releaseLock();

  return httpResponse;
}

/**
 * Private helper method that is used to examine function execution
 * in an effort to optimize performance.
 */
function profileGetMarketPrice(itemId, regionId, stationId, orderType, refresh)
{
  var startTime = new Date().getTime();
  var price = getMarketPrice(itemId, regionId, stationId, orderType, refresh);
  var endTime = new Date().getTime();

  if (typeof(price) == 'number')
  {
    return endTime - startTime;
  }
  else
  {
    return price;
  }
}
