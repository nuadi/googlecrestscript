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
 * version 1.1
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
      var typeUrl = "?type=http://public-crest.eveonline.com/types/" + itemId + "/";

      // Make the call to get some market data
      var marketResponse = fetchUrl(marketUrl + typeUrl);

      var jsonMarket = JSON.parse(marketResponse);

      returnPrice = getPrice(jsonMarket, stationId, orderType)
    }
  }
  catch (unknownError)
  {
    returnPrice = unknownError.message;
  }

  return returnPrice;
}

/**
 * Private helper method that will determine the best price for a given item from the
 * market data provided.
 *
 * version 1.1
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
    var location = order['location'];

    if (stationId == location['id'])
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
 * version 1.2
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
    var marketResponse = fetchUrl(marketUrl);
  
    var jsonMarket = JSON.parse(marketResponse);
  
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
 * version 1.2
 *
 * @param {url} url The URL to contact
 * @param {options} options The fetch options to utilize in the request
 */
function fetchUrl(url, options)
{
  var semaphore = "lock";
  var semaphoreLife = 2;

  if (options == null)
  {
    httpResponse = UrlFetchApp.fetch(url);
  }
  else
  {
    httpResponse = UrlFetchApp.fetch(url, options);
  }
  
  // Wait based on rate limit
  var requestsPerSecond = 150;
  Utilities.sleep(1000/requestsPerSecond);

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
