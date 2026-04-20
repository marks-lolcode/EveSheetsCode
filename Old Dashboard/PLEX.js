function plexPrice(orderType) {
  orderType = (orderType || "sell").toLowerCase(); // Default to "sell"

  if (orderType !== "sell" && orderType !== "buy") {
    return "Invalid order type: use 'buy' or 'sell'";
  }

  const regionID = 19000001; // Global market region (NEW)
  const typeID = 44992; // PLEX
  const url = `https://esi.evetech.net/latest/markets/${regionID}/orders/?type_id=${typeID}&order_type=${orderType}`;

  try {
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    const data = JSON.parse(response.getContentText());

    if (!Array.isArray(data) || data.length === 0) return "No market data";

    const prices = data.map(order => order.price).filter(p => typeof p === "number");

    if (prices.length === 0) return "No valid prices";

    return orderType === "sell" ? Math.min(...prices) : Math.max(...prices);

  } catch (e) {
    return `Error: ${e.message}`;
  }
}