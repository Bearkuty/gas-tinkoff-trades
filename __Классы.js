class DropdownList {
  constructor(target='a1',source){
    var cell = SpreadsheetApp.getActive().getRange(target);
    var rule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInRange(source).build();
    cell.setDataValidation(rule);
  }
}

class TinkoffClient {
  constructor(token){
    this.token = token
    this.baseUrl = 'https://api-invest.tinkoff.ru/openapi/'
  }
  _makeApiCall(methodUrl){
    const url = this.baseUrl + methodUrl
    Logger.log(`[API Call] ${url}`)
    const params = {'escaping': false, 'headers': {'accept': 'application/json', "Authorization": `Bearer ${this.token}`}}
    const response = UrlFetchApp.fetch(url, params)
    if (response.getResponseCode() == 200)
      return JSON.parse(response.getContentText()) 
  }
  getFIGIbyTicker(ticker){
    const url = `market/search/by-ticker?ticker=${ticker}`
    const data = this._makeApiCall(url)
    return data.payload.instruments[0].figi
  }
  getInstrumentByFigi(figi){
    const url = `market/search/by-figi?figi=${figi}`
    const data = this._makeApiCall(url)
    return data.payload
  }
  getTickerByFigi(figi){
    const url = `market/search/by-figi?figi=${figi}`
    const data = this._makeApiCall(url)
    return data.payload.ticker
  }
  getOrderbookByFigi(figi){
    const url = `market/orderbook?depth=1&figi=${figi}`
    const data = this._makeApiCall(url)
    return data.payload
  }
  getOperations(from, to, figi){
    // Аргументы `from` и `to` должны быть в ISO 8601 формате
    const url = `operations?from=${from}&to=${to}&figi=${figi}`
    const data = this._makeApiCall(url)
    return data.payload.operations
  }
  getAll (from, to) {
    const url = `operations?from=${from}&to=${to}`
    const data = this._makeApiCall(url)
    return data.payload.operations
  }
  getAllIIS (from, to, IISid) {
    const url = `operations?from=${from}&to=${to}&brokerAccountId=${IISid}`
    const data = this._makeApiCall(url)
    return data.payload.operations
  }
  getPort(){
    const url = `portfolio`
    const data = this._makeApiCall(url)
    return data.payload.positions
  }
  getCur(){
    const url = `portfolio/currencies`
    const data = this._makeApiCall(url)
    return data.payload.currencies
  }
  getIIS(IISid){
    const url = `portfolio?brokerAccountId=${IISid}`
    const data = this._makeApiCall(url)
    return data.payload.positions
  }
  getIISid(){
    const url = `user/accounts`
    const data = this._makeApiCall(url)
    return data.payload.accounts
  }
  usdval(){
    const url = `market/orderbook?figi=BBG0013HGFT4&depth=1`
    const data = this._makeApiCall(url)
    return data.payload.lastPrice
  }
}

