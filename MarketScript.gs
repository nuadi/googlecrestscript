// Google Crest Script (GCS)
// version 4a
// /u/nuadi @ Reddit
//
// LICENSE: Use at your own risk, and fly safe.

// Global variable needed to track number of retries attempted
var retries = 0;

/**
 * Private helper function that is used to initialize the refresh token
 */
function initializeGetMarketPrice()
{
  // Test Amarr Fuel Block in Dodixie
  var itemId = 20059;
  var regionId = 10000032;
  var stationId = 60011866;
  var orderType = "SELL";

  var price = getMarketPrice(itemId, regionId, stationId, orderType);
  Logger.log(price);

  itemId = 20059;
  regionId = 10000002;
  stationId = 60003760;

  price = getMarketPrice(itemId, regionId, stationId, orderType);
  Logger.log(price);
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
  try
  {
    var returnPrice = 0;

    // Validate incoming arguments
    if (itemId == null || typeof(itemId) != "number")
    {
      returnPrice = "Invalid Item ID";
    }
    else if (regionId == null || typeof(regionId) != "number")
    {
      returnPrice = "Invalid Region ID";
    }
    else if (stationId == null || typeof(stationId) != "number")
    {
      returnPrice = "Invalid Station ID";
    }
    else if (orderType == null || typeof(orderType) != "string")
    {
      returnPrice = "Invalid order type";
    }
    else
    {
      orderType = orderType.toLowerCase();

      // Setup variables for the market endpoint we want
      var marketUrl = "https://public-crest.eveonline.com/market/" + regionId + "/orders/" + orderType + "/";
      var typeUrl = "?type=https://public-crest.eveonline.com/types/" + itemId + "/";

      // Make the call to get some market data
      var jsonMarket = JSON.parse(fetchUrl(marketUrl + typeUrl));
      returnPrice = getPrice(jsonMarket, stationId, orderType)
    }
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
        returnPrice = getMarketPrice(itemId, regionId, stationId, orderType, refresh); 
      }
      else
      {
        returnPrice = "";
        for (i in unknownError)
        {
          returnPrice += i + ": " + unknownError[i] + "\n";
        }
      }
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

  try
  {
    // Only validate arguments within the context of this function
    // Further validation will occur inside getMarketPrice
    if (itemIdList == null || typeof(itemIdList) != "object")
    {
      returnValues = "Invalid Item list";
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
  }
  catch (error)
  {
    returnValues = "Error received = " + error.message;
  }

  return returnValues;
}

/**
 * Return all market orders for an item from a region.
 *
 * @param {itemId}
 * @param {regionId}
 * @param {orderType}
 * @param {refresh}
 * @customfunction
 */
function getOrders(itemId, regionId, orderType, refresh)
{
  try
  {
    var marketReturn = [];

    // Validate incoming arguments
    if (itemId == null || typeof(itemId) != "number")
    {
      marketReturn = "Invalid Item ID";
    }
    else if (regionId == null || typeof(regionId) != "number")
    {
      marketReturn = "Invalid Region ID";
    }
    else if (orderType == null || typeof(orderType) != "string")
    {
      marketReturn = "Invalid order type";
    }
    else
    {
      orderType = orderType.toLowerCase();

      if (orderType != 'sell')
      {
        marketReturn = "Invalid order type";
      }
      else
      {
        // Setup variables for the market endpoint we want
        var marketUrl = "https://public-crest.eveonline.com/market/" + regionId + "/orders/" + orderType + "/";
        var typeUrl = "?type=https://public-crest.eveonline.com/types/" + itemId + "/";
        Logger.log('Calling URL: ' + marketUrl + typeUrl);

        // Make the call to get some market data
        var jsonMarket = JSON.parse(fetchUrl(marketUrl + typeUrl));
        var marketItems = jsonMarket['items'];

        // Convert to an array for proper output
        marketReturn.push(['Issued', 'Price', 'Volume', 'Location']);
        for (var rowKey in marketItems)
        {
          var rowData = marketItems[rowKey];
          var newRow = [];
          for (var colKey in rowData)
          {
            if (colKey == 'issued')
            {
              var dateValues = rowData[colKey].split(/[T-]/);
              var dateString = dateValues.slice(0,3).join('/') + ' ' + dateValues[3];
              Logger.log("Converting date string: " + dateString);
              newRow.push(new Date(dateString));
            }
            else if (colKey == 'price' || colKey == 'volume')
            {
              newRow.push(rowData[colKey]);
            }
            else if (colKey == 'location')
            {
              var locationData = rowData[colKey];
              newRow.push(locationData['name']);
            }
          }
          marketReturn.push(newRow);
        }
      }
    }
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
        marketReturn = getOrders(itemId, regionId, orderType, refresh); 
      }
      else
      {
        marketReturn = "";
        for (i in unknownError)
        {
          marketReturn += i + ": " + unknownError[i] + "\n";
        }
      }
    }
  }

  SpreadsheetApp.flush();
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
    history = "Invalid Item ID";
  }
  else if (regionId == null || typeof(regionId) != "number")
  {
    history = "Invalid Region ID";
  }
  else if (property != "orderCount" && property != "lowPrice" && property != "highPrice" && property != "avgPrice" && property != "volume")
  {
    history = "Property must be one of: 'orderCount', 'lowPrice', 'highPrice', 'avgPrice', 'volume'";
  }
  else
  { 
    // Setup variables for the market endpoint we want
    var marketUrl = "https://public-crest.eveonline.com/market/" + regionId + "/types/" + itemId + "/history/"
  
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
  headers = {"User-Agent": "Google Crest Script version 4a (/u/nuadi @Reddit.com)"}
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
