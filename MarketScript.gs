// Google CREST/ESI Script (GCS)
var version = '12b'
// nuadi.bantine@gmail.com
//
// LICENSE: Use at your own risk, and fly safe.


// Global variables used in order comparison function and set by
// Advanced Orders function. Default is the Price column in ascending order.
var sortIndex = 1;
var sortOrder = 1;


/**
 * >>START HERE<<
 * Use this function to initialize, or test, the script and permissions.
 */
function initializeGetMarketPrice()
{
  // Test PLEX in Dodixie
  var itemId = 44992;
  var regionId = 10000032;
  var stationId = 60011866;
  var orderType = 'SELL';

  var price = getStationMarketPrice(itemId, regionId, stationId, orderType);
  Logger.log(price);

  // Now in Jita
  regionId = 10000002;
  stationId = 60003760;

  price = getStationMarketPrice(itemId, regionId, stationId, orderType);
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
 * Count the number of available orders for a given item in a station for a given order type.
 * @param {itemId} itemId the item ID of the product to look up.
 * @param {regionId} regionId the region ID for the market to look up.
 * @param {stationId} stationId the station ID for the market to look up.
 * @param {orderType} orderType this should be set to "sell" or "buy" orders.
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function countStationOrders(itemId, regionId, stationId, orderType, refresh)
{
  var ordersNumber = 0;

  var options = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['stationId', stationId],
    ['headers', false],
    ['orderType', orderType],
    ['refresh', refresh]
  ];
  
  var orders = getOrdersAdv(options);
  return orders.length;
}


/**
 * Count the number of available units for a given item in a station for a given order type.
 * @param {itemId} itemId the item ID of the product to look up.
 * @param {regionId} regionId the region ID for the market to look up.
 * @param {stationId} stationId the station ID for the market to look up.
 * @param {orderType} orderType this should be set to "sell" or "buy" orders.
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function countStationVolume(itemId, regionId, stationId, orderType, refresh)
{
  var ordersNumber = 0;

  var options = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['stationId', stationId],
    ['headers', false],
    ['orderType', orderType],
    ['refresh', refresh]
  ];
  
  var availableVolume = 0;
  var orders = getOrdersAdv(options);
  for (var row = 0; row < orders.length; row++)
  {
    availableVolume += orders[row][2];
  }
  return availableVolume;
}


/**
 * Private helper method that wraps the UrlFetchApp in a semaphore
 * to prevent service overload.
 *
 * @param {url} url The URL to contact
 * @param {esi} esi Set to TRUE to override semaphore subsystem and skip locking
 * @param {options} options The fetch options to utilize in the request
 */
function fetchUrl(url, esi=false)
{
  if (gcsGetLock() || esi)
  {
    // Make the service call
    headers = {'User-Agent': 'Google CREST/ESI Script version ' + version + ' (nuadi.bantine@gmail.com)'}
    params = {'headers': headers}
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
 * If no semaphore is open, the function will start over with no wait.
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
 * Returns the adjusted price used for industrial calculations.
 * @param {itemId} itemId the item ID of the product to look up
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value
 * @customfunction
 */
function getAdjustedPrice(itemId, refresh)
{
  var dataKey = itemId + 'adjustedPrice' + refresh;
  var adjustedPrice = CacheService.getDocumentCache().get(dataKey);
  if (adjustedPrice == null)
  {
    var sixHours = 60 * 60 * 6; // in seconds
    var marketPriceEndpoint = 'https://crest-tq.eveonline.com/market/prices/';
    var marketPriceData = JSON.parse(fetchUrl(marketPriceEndpoint));
    var itemPrices = marketPriceData['items'];
    for (var row = 0; row < itemPrices.length; row++)
    {
      var itemPriceData = itemPrices[row];
      if (itemId == itemPriceData['type']['id'])
      {
        adjustedPrice = itemPriceData['adjustedPrice'];
        CacheService.getDocumentCache().put(dataKey, adjustedPrice, sixHours);
        break;
      }
    }
  }
  else
  {
    adjustedPrice = Number(adjustedPrice);
  }

  return adjustedPrice;
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
  if (historicalData.length != days)
  {
    days = historicalData.length;
  }
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
  if (historicalData.length != days)
  {
    days = historicalData.length;
  }
  for (var rowNumber = 0; rowNumber < historicalData.length; rowNumber++)
  {
    totalVolume += historicalData[rowNumber][1];
  }
  return totalVolume / days;
}


/**
 * Returns the cost index for a given activity in a given solar system.
 * @param {systemName} systemName the name of the solar system
 * @param {activity} activity the name of the activity to look up. Can be "Invention", "Manufacturing", "Researching Time Efficiency", "Researching Material Efficiency", or "Copying".
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getCostIndex(systemName, activity, refresh)
{
  var dataKey = systemName + activity + 'costIndex' + refresh;
  var costIndex = CacheService.getDocumentCache().get(dataKey);
  if (costIndex == null)
  {
    var sixHours = 60 * 60 * 6; // in seconds
    var industrySystemsEndpoint = 'https://crest-tq.eveonline.com/industry/systems/';
    var industrySystemsData = JSON.parse(fetchUrl(industrySystemsEndpoint));
    var costIndices = industrySystemsData['items'];
    for (var row = 0; row < costIndices.length; row++)
    {
      var costIndexData = costIndices[row];
      if (systemName == costIndexData['solarSystem']['name'])
      {
        //return 'name found';
        var systemIndices = costIndexData['systemCostIndices'];
        for (var arow = 0; arow < systemIndices.length; arow++)
        {
          var activityData = systemIndices[arow];
          if (activity == activityData['activityName'])
          {
            costIndex = activityData['costIndex'];
            CacheService.getDocumentCache().put(dataKey, costIndex, sixHours);
            break;
          }
        }
        if (costIndex != null)
        {
          break;
        }
      }
    }
  }
  else
  {
    costIndex = Number(costIndex);
  }

  return costIndex;
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
    throw new Error('No options found');
  }
  else if (options[0] == null || options[0].length < 2)
  {
    throw new Error('Options must have 2 columns');
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

  var historyEndpoint = 'https://crest-tq.eveonline.com/market/' + regionId + '/history/';
  var typeUrl = '?type=https://crest-tq.eveonline.com/inventory/types/' + itemId + '/';
  var historyJson = JSON.parse(fetchUrl(historyEndpoint + typeUrl));

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
 * Private helper function that will return all information for a given item.
 * @param {itemId} itemId the item ID of the product to look up
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 */
function getItemInfo(itemId, refresh)
{
  var itemDataArray = [];
  var itemData = getItemJson(itemId, refresh);
  for (var key in itemData)
  {
    var newRow = [key, itemData[key]];
    itemDataArray.push(newRow);
  }
  return itemDataArray;
}


/**
 * Private helper function that will pull all information for a given item from CREST.
 * @param {itemId} itemId the item ID of the product to look up
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 */
function getItemJson(itemId, refresh)
{
  var itemEndpoint = 'https://crest-tq.eveonline.com/inventory/types/' + itemId + '/';
  var itemData = JSON.parse(fetchUrl(itemEndpoint));
  return itemData;
}


/**
 * Returns the volume (assembled only) for a given item ID.
 * @param {itemId} itemId the item ID of the product to look up
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getItemVolume(itemId, refresh)
{
  var itemData = getItemJson(itemId, refresh);
  return itemData['volume'];
}


/**
 * Returns a list of all market items found in a given group ID.
 *
 * @param {groupId} groupId Defines the parent group to retrieve subgroups. Default is base market group.
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getMarketGroupItems(groupId, refresh)
{
  var groupTypesEndpoint = 'https://crest-tq.eveonline.com/market/types/?group=https://crest-tq.eveonline.com/market/groups/' + groupId + '/';

  var groupItemList = [];

  var paging = true;
  while (paging)
  {
    var groupTypesJson = JSON.parse(fetchUrl(groupTypesEndpoint));
    var groupTypes = groupTypesJson['items'];

    for (var itemHandle in groupTypes)
    {
      var marketItem = groupTypes[itemHandle];
      groupItemList.push([marketItem['type']['name'], marketItem['type']['id']]);
    }

    if (groupTypesJson['next'] != null)
    {
      groupTypesEndpoint = groupTypesJson['next']['href'];
    }
    else
    {
      paging = false;
    }
  }

  sortIndex = 0;
  groupItemList.sort(basicCompare);

  var returnArray = [];
  returnArray.push(['Name', 'ID']);

  return returnArray.concat(groupItemList);
}


/**
 * Returns a list of all top-level market groups, or child groups of a given ID.
 *
 * @param {groupId} groupId Defines the parent group to retrieve subgroups. Default is base market group.
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getMarketGroups(groupId, refresh)
{
  var marketGroupEndpoint = 'https://crest-tq.eveonline.com/market/groups/';
  var marketGroupJson = JSON.parse(fetchUrl(marketGroupEndpoint));
  var marketGroups = marketGroupJson['items'];

  var marketGroupList = [];

  for (var groupHandle in marketGroups)
  {
    var marketGroup = marketGroups[groupHandle];

    var pushGroup = false;
    if (groupId == null || groupId == '')
    {
      // Only push base groups
      pushGroup = marketGroup['parentGroup'] == null;
    }
    else
    {
      // Only push children of this group
      pushGroup = marketGroup['parentGroup'] != null && marketGroup['parentGroup']['id'] == groupId
    }
    
    if (pushGroup)
    {
      marketGroupList.push([marketGroup['name'], marketGroup['id']]);
    }
  }

  sortIndex = 0;
  marketGroupList.sort(basicCompare);

  var returnArray = [];
  returnArray.push(['Name', 'ID']);

  return returnArray.concat(marketGroupList);
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
 * a given market query.
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
    throw new Error('Item ID cannot be NULL.');
  }
  else if (typeof(itemId) != 'number')
  {
    throw new Error('Item ID must be a number. Instead found a(n) ' + typeof(itemId) + '.');
  }
  else if (regionId == null)
  {
    throw new Error('Region ID cannot be NULL.');
  }
  else if (typeof(regionId) != 'number')
  {
    throw new Error('Region ID must be a number. Instead found a(n) ' + typeof(regionId) + '.');
  }
  else if (orderType == null)
  {
    throw new Error('Order type cannot be NULL.');
  }
  else if (typeof(orderType) != 'string')
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
    var marketUrl = 'https://esi.tech.ccp.is/latest/markets/' + regionId + '/orders/';
    var parameterUrl = '?order_type=' + orderType + '&type_id=' + itemId;
    Logger.log('Pulling market orders from url: ' + marketUrl + parameterUrl)
    
    var fullUrl = marketUrl + parameterUrl;
    
    // TODO: Insert a cachinator check here
    
    try
    {
      // Make the call to get some market data
      marketData = JSON.parse(fetchUrl(marketUrl + parameterUrl));
      
      // TODO: Insert a cachinator.save() call here
    }
    catch (unknownError)
    {
      Logger.log(unknownError);
      throw unknownError;
    }
  }

  return marketData;
}


/**
 * Returns the market price for a given item. (Deprecated)
 *
 * @param {itemId} itemId the item ID of the product to look up
 * @param {regionId} regionId the region ID for the market to look up
 * @param {stationId} stationId the station ID for the market to focus on
 * @param {orderType} orderType this should be set to "sell" or "buy" orders
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value
 */
function getMarketPrice(itemId, regionId, stationId, orderType, refresh)
{
  return getStationMarketPrice(itemId, regionId, stationId, orderType, refresh);
}


/**
 * Returns a list of NPC corporations and their IDs.
 *
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getNPCCorporations(refresh)
{
  var npcCorpEndpoint = 'https://crest-tq.eveonline.com/corporations/npccorps/';
  var npcCorpJson = JSON.parse(fetchUrl(npcCorpEndpoint));

  var npcCorpList = npcCorpJson['items'];

  var outputList = [];

  for (var npcCorpHandle in npcCorpList)
  {
    var npcCorp = npcCorpList[npcCorpHandle];
    outputList.push([npcCorp['name'], npcCorp['id']]);
  }

  sortIndex = 0;
  outputList.sort(basicCompare);

  var returnArray = [];
  returnArray.push(['Name', 'ID']);

  return returnArray.concat(outputList);
}


/**
 * Returns the loyalty store items for a given corporation ID.
 *
 * @param {corpId} corpId The ID of the NPC Corporation.
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value.
 * @customfunction
 */
function getNPCLoyaltyStore(corpId, refresh)
{
  var loyaltyStoreEndpoint = 'https://crest-tq.eveonline.com/corporations/' + corpId + '/loyaltystore/';
  var loyaltyStoreJson = JSON.parse(fetchUrl(loyaltyStoreEndpoint));

  var loyaltyStoreListings = loyaltyStoreJson['items'];

  var tempList = [];

  for (var entryHandle in loyaltyStoreListings)
  {
    var loyaltyStoreListing = loyaltyStoreListings[entryHandle];

    var loyaltyItem = loyaltyStoreListing['item'];
    var itemName = loyaltyItem['name'];
    var itemId = loyaltyItem['id'];

    var qty = loyaltyStoreListing['quantity'];

    var lpCost = loyaltyStoreListing['lpCost'];
    var iskCost = loyaltyStoreListing['iskCost'];

    var reqItemLimit = 5;
    var reqItemColumns = 3;

    var reqTempList = [];
    for (var reqItems = 0; reqItems < reqItemLimit; reqItems++)
    {
      for (var columns = 0; columns < reqItemColumns; columns++)
      {
        reqTempList.push('');
      }
    }

    var requiredItemList = loyaltyStoreListing['requiredItems'];
    for (var itemNumber = 0; itemNumber < requiredItemList.length; itemNumber++)
    {
      if (itemNumber >= reqTempList.length / 1)
      {
        break;
      }

      var requiredEntry = requiredItemList[itemNumber];
      var reqItem = requiredEntry['item'];

      var reqItemName = reqItem['name'];
      var reqItemId = reqItem['id'];
      var reqItemQty = requiredEntry['quantity']

      var colPosition = itemNumber * reqItemColumns;
      reqTempList[colPosition] = reqItemName;
      reqTempList[colPosition + 1] = reqItemId;
      reqTempList[colPosition + 2] = reqItemQty;
    }

    var newItemEntry = [itemName, itemId, qty, lpCost, iskCost];
    newItemEntry = newItemEntry.concat(reqTempList);
    tempList.push(newItemEntry);
  }

  sortIndex = 0;
  tempList.sort(basicCompare);

  var returnArray = [];
  var headers = ['Item Name', 'ID', 'Qty', 'LP Cost', 'ISK Cost'];
  var reqHeaders = [];
  for (var reqItems = 0; reqItems < reqItemLimit; reqItems++)
  {
    for (var columns = 0; columns < reqItemColumns; columns++)
    {
      switch (columns)
      {
        case 0:
          reqHeaders.push('Required Item ' + (reqItems + 1));
          break;
        case 1:
          reqHeaders.push('Required Item ' + (reqItems + 1) + ' ID');
          break;
        case 2:
          reqHeaders.push('Required Item ' + (reqItems + 1) + ' Qty');
          break;
      }
    }
  }
  headers = headers.concat(reqHeaders)
  returnArray.push(headers);

  return returnArray.concat(tempList);
}


/**
 * Return all market orders for an item from a region.
 *
 * @param {itemId} itemId the item ID of the product to look up
 * @param {regionId} regionId the region ID for the market to look up
 * @param {orderType} orderType this should be set to "sell" or "buy" orders
 * @param {refresh} refresh (Optional) Change this value to force Google to refresh return value
 * @param {filters} filters (Optional) An array of filters to place on orders during processing.
 * @customfunction
 */
function getOrders(itemId, regionId, orderType, refresh, filters)
{
  var orderOptions = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['orderType', orderType.toLowerCase()],
    ['refresh', refresh]
  ];
  
  if (filters != null)
  {
    for (var row = 0; row < filters.length; row++)
    {
      if (filters[row].length > 1)
      {
        orderOptions.push(filters[row]);
      }
      else
      {
        orderOptions.push(filters);
      }
    }
  }

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
  var minPrice = null;
  var maxPrice = null;
  var minVolume = null;
  var maxVolume = null;

  var sortOrderSet = false;

  if (options.length <= 0)
  {
    throw new Error('No options found');
  }
  else if (options[0] == null || options[0].length < 2)
  {
    throw new Error('Options must have 2 columns');
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
      else if (optionKey == 'minPrice')
      {
        minPrice = optionValue;
      }
      else if (optionKey == 'maxPrice')
      {
        maxPrice = optionValue;
      }
      else if (optionKey == 'minVolume')
      {
        minVolume = optionValue;
      }
      else if (optionKey == 'maxVolume')
      {
        maxVolume = optionValue;
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
  var marketItems = jsonMarket;
  
  // Convert all data to an array for proper output
  var outputArray = [];
  for (var rowKey in marketItems)
  {
    var saveRow = true;
    var rowData = marketItems[rowKey];
    var newRow = [];
    for (var colKey in rowData)
    {
      if (colKey == 'order_id' && showOrderId == true)
      {
        newRow[orderIdColumn] = rowData[colKey];
      }
      else if (colKey == 'issued')
      {
        var dateValues = rowData[colKey].split(/[T-]/);
        var dateString = dateValues.slice(0,3).join('/') + ' ' + dateValues[3];
        dateString = dateString.replace('Z', '');
        Logger.log('Converting date string: ' + dateString);
        newRow[0] = new Date(dateString);
      }
      else if (colKey == 'min_volume' && orderType == 'buy')
      {
        newRow[minVolumeColumn] = rowData[colKey];
      }
      else if (colKey == 'price')
      {
        var orderPrice = rowData[colKey];
        if (minPrice != null && orderPrice < minPrice)
        {
          saveRow = false;
        }
        if (maxPrice != null && orderPrice > maxPrice)
        {
          saveRow = false;
        }
        newRow[1] = orderPrice;
      }
      else if (colKey == 'range' && orderType == 'buy')
      {
        newRow[rangeColumn] = rowData[colKey];
      }
      else if (colKey == 'volume_remain')
      {
        var orderVolume = rowData[colKey];
        if (minVolume != null && orderVolume < minVolume)
        {
          saveRow = false;
        }
        if (maxVolume != null && orderVolume > maxVolume)
        {
          saveRow = false;
        }
        
        newRow[2] = orderVolume;
      }
      else if (colKey == 'location_id')
      {
        var locationId = rowData[colKey];
        newRow[locationColumn] = locationId;

        if (stationId != null && stationId != locationId)
        {
          saveRow = false;
          break;
        }

        if (showStationId == true)
        {
          newRow[stationIdColumn] = locationId;
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
 * @param {filters} filters (Optional) An array of filters to place on orders during processing.
 * @customfunction
 */
function getRegionMarketPrice(itemId, regionId, orderType, refresh, filters)
{
  var orderOptions = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['orderType', orderType.toLowerCase()],
    ['refresh', refresh],
    ['headers', false]
  ];
  
  if (filters != null)
  {
    for (var row = 0; row < filters.length; row++)
    {
      if (filters[row].length > 1)
      {
        orderOptions.push(filters[row]);
      }
      else
      {
        orderOptions.push(filters);
      }
    }
  }

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
 * @param {filters} filters (Optional) An array of filters to place on orders during processing.
 * @customfunction
 */
function getStationMarketPrice(itemId, regionId, stationId, orderType, refresh, filters)
{
  var orderOptions = [
    ['itemId', itemId],
    ['regionId', regionId],
    ['stationId', stationId],
    ['orderType', orderType.toLowerCase()],
    ['refresh', refresh],
    ['headers', false]
  ];

  if (filters != null)
  {
    for (var row = 0; row < filters.length; row++)
    {
      if (filters[row].length > 1)
      {
        orderOptions.push(filters[row]);
      }
      else
      {
        orderOptions.push(filters);
      }
    }
  }

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
  var submenus = [{name: 'Check for updates', functionName: 'versionCheck'}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu('GCS/ESI', submenus);
}


/**
 * Private helper function that will check for a new version of GCS/ESI.
 */
function versionCheck()
{
  var versionEndpoint = 'https://raw.githubusercontent.com/nuadi/googlecrestscript/master/version';
  var newVersion = fetchUrl(versionEndpoint);

  if (newVersion != null)
  {
    newVersion = newVersion.getContentText().trim();
    Logger.log('Current version from Github: ' + newVersion);

    var message = 'You are using the latest version of GCS/ESI. Fly safe. o7';
    var title = 'No updates found';
    if (newVersion > version)
    {
      message = 'A new version of GCS/ESI is available on GitHub.';
      title = 'GCS/ESI version ' + newVersion + ' available!';
    }
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 120);
  }
}
