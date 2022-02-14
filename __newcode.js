// TODO:
// 
// Изменить  createTickersDropdownList
// чтобы список тикеров кэшился на листе summary -- готово
// и обновлялся раз в сутки или через меню
// 
// 
// 
// 
// 

const OPENAPI_TOKEN = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("API_Token").getValue()
const TRADING_START_AT = new Date('Jan 01, 2019 10:00:00')
const MILLIS_PER_DAY = 1000 * 60 * 60 * 24
const CACHE = CacheService.getScriptCache()

function getArrayFromRange(range){
  let values=range.getValues()
  for(let[key,val]of values.entries())
    if(val[0])
      values[key]=val[0]
    else
      values.splice(key)
  return values

}

function buildRangeWithAllTradedTickers(){
  const ss=SpreadsheetApp.getActive()
  let values=[]
  for(let item of getAllTradedTickers())
    values.push([item])
  let sheet=ss.getSheetByName('summary')
  if(!sheet)
    sheet=ss.insertSheet().setName('summary')
  const range=sheet.getRange(1,1,values.length)
  range.setValues(values)
  ss.setNamedRange('tickers',range)
}

function getActiveSheetName(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getName()
}

function getAllTradedTickers(){
  // возвращает массив со всеми торговавшимися тикерами
  from = TRADING_START_AT.toISOString()
  to = new Date(new Date() + MILLIS_PER_DAY).toISOString()
  const operations = tinkoffClient.getAll (from, to)
  let figivalues = []
  for (let i=operations.length-1; i>=0; i--) {
    const {figi} = operations[i]
    figivalues.push(figi)
  }

  figivalues = [...new Set(figivalues)];
  // let tickervalues=figivalues.map(tinkoffClient.getTickerByFigi)
  const tickervalues=[]
  for(let item of figivalues){
    if (item){
      var ticker = tinkoffClient.getTickerByFigi(item)
      tickervalues.push(ticker)
      Utilities.sleep(200)
    }
  }
  return(tickervalues)
}

const tinkoffClient = new TinkoffClient(OPENAPI_TOKEN)

function createSheetsForTickers(){
  const portfolio = tinkoffClient.getPort()
  for (let i=portfolio.length-1; i>=0; i--) {
    const {ticker, name, balance, averagePositionPrice, expectedYield} = portfolio [i]
    if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ticker)===null){
      const rangevalues=[
        [ticker,'тикер',],
        ['=-sumproduct(c:c;D:D)','-расход +доход'], // баланс расходов-доходов
        ['=sum(f:f)','комиссия'], // комса
        ['=sum(c:c)','кол-во'], // кол-во бумаг сейчас
        ['=(A2+A3)/A4','средняя'], // средняя с учетом всего (докупок, продаж (в т. ч. частичных), комиссии)
        [null,null],
        ['=gettrades(a1)',''],
      ]
      SpreadsheetApp.getActiveSpreadsheet().insertSheet().setName(ticker)
      .getRange(1,1,rangevalues.length,2).setValues(rangevalues)
      createTickersDropdownList()
    }
  }
}

function createTickersDropdownList(){
  new DropdownList('a1',SpreadsheetApp.getActive().getRange('tickers'))
}

function isoToDate(dateStr){
  const str = dateStr.replace(/-/,'/').replace(/-/,'/').replace(/T/,' ').replace(/\+/,' \+').replace(/Z/,' +00')
  return new Date(str)
}
    
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  const entries = [
    {name:'buildRangeWithAllTradedTickers',functionName:'buildRangeWithAllTradedTickers'},
    {name:'createTickersDropdownList',functionName:'createTickersDropdownList'},
    {name:'createSheetsForTickers',functionName:'createSheetsForTickers'},
    null,
    {name : "Обновить", functionName : "refresh"}
  ]
  sheet.addMenu("TI", entries)
};

function refresh() {
  SpreadsheetApp.getActiveSpreadsheet().getRange('Z1').setValue(new Date().toTimeString());
}

function getPrice(ticker, refresh){
  const figi = tinkoffClient.getFIGIbyTicker(ticker)
  var {lastPrice} = tinkoffClient.getOrderbookByFigi(figi)
  return lastPrice
}

function getAllTickers(figi){
  var ticker
  if (!figi){
    ticker = ""
  }else{
    var {ticker} = tinkoffClient.getInstrumentByFigi(figi)
  }
  return ticker
}

function _calculateTrades(trades) {
  let totalSum = 0
  let totalQuantity = 0
  for (let j in trades) {
    const {quantity, price} = trades[j]
    totalQuantity += quantity
    totalSum += quantity * price
  }
  const weigthedPrice = totalSum / totalQuantity
  return [totalQuantity, totalSum, weigthedPrice]
}

function getTrades(ticker, from, to) {
  const figi = tinkoffClient.getFIGIbyTicker(ticker)
  if (!from) {
    from = TRADING_START_AT.toISOString()
  }
  if (!to) {
    const now = new Date()
    to = new Date(now + MILLIS_PER_DAY)
    to = to.toISOString()
  }
  const operations = tinkoffClient.getOperations(from, to, figi)
  const values = []
  var com_val
  for (let i=operations.length-1; i>=0; i--) {
    const {operationType, status, trades, date, currency, figi, commission} = operations[i]
    if (operationType == "BrokerCommission" || status == "Decline")
      continue
    let [totalQuantity, totalSum, weigthedPrice] = _calculateTrades(trades)
    if (commission){
      com_val = commission.value
    }else{
      com_val = "-"
    }
    if (operationType == "Sell") {
      totalQuantity = -totalQuantity
      totalSum = -totalSum
      commission.value = -commission.value
    }
    values.push([isoToDate(date), operationType, totalQuantity, weigthedPrice, currency, com_val])
  }
  return values
}

function getAllTrades(from, to, refresh){
  if (!from){
    from = TRADING_START_AT.toISOString()
  }else{
    from = from.toISOString()
  }
  if (!to){
    to = new Date(new Date() + MILLIS_PER_DAY).toISOString()
  }else{
    to = to.toISOString()
  }
  const operations = tinkoffClient.getAll (from, to)
  const values = []
  var com_val
  values.push(["Дата","Тикер","Тип","Кол-во","Цена за 1","Комиссия","Итого","Валюта"])
  for (let i=operations.length-1; i>=0; i--) {
    const {operationType, status, trades, date, currency, figi, commission, payment} = operations[i]
    if (operationType == "BrokerCommission" || operationType == "PayIn" || operationType == "PayOut" || status == "Decline")
      continue
    // если нужно отобразить комиссию брокера (BrokerCommission), пополнение (PayIn) или вывод (PayOut) средств со счета, удалите ненужный вариант. Например, если Вы хотите видеть отображение вывода средств со счёта, удалите " operationType == "PayIn" ||" из строки выше
    let [totalQuantity, totalSum, weigthedPrice] = _calculateTrades(trades)
    if (commission){
      com_val = -commission.value
    }else{
      com_val = 0
    }
    if (operationType == "Tax" || operationType == "TaxDividend" || operationType == "Dividend" || operationType == "PartRepayment" || operationType == "Coupon" || operationType == "TaxCoupon"){
      totalQuantity = '-'
      weigthedPrice = '-'
    }
    var ticker
    if (!figi){
      ticker = ""
    } else {
      var ticker = tinkoffClient.getTickerByFigi(figi)
    }
    values.push([isoToDate(date), ticker, operationType, totalQuantity, weigthedPrice, com_val, payment-com_val, currency])
  }
  return values
}

function getID(){
  const users = tinkoffClient.getIISid()
  for (let i=users.length-1; i>=0; i--) {
    const {brokerAccountId, brokerAccountType} = users [i]
    if (brokerAccountType == "TinkoffIis")
      IISid = brokerAccountId
  }
  return IISid
}

function getAllTradesIIS(from,to){
  IISid = getID()
  if (!from){
    from = TRADING_START_AT.toISOString()
  }else{
    from = from.toISOString()
  }
  if (!to){
    to = new Date(new Date() + MILLIS_PER_DAY).toISOString()
  }else{
    to = to.toISOString()
  }
  const operations = tinkoffClient.getAllIIS (from, to, IISid)
  const values = []
  values.push(["Дата","Тикер","Тип","Кол-во","Цена за 1","Итого","Валюта"])
  for (let i=operations.length-1; i>=0; i--) {
    const {operationType, status, trades, date, currency, figi, commission, payment} = operations[i]
    if (operationType == "BrokerCommission" || operationType == "PayIn" || operationType == "PayOut" || status == "Decline")
      continue
    // если нужно отобразить комиссию брокера (BrokerCommission), пополнение (PayIn) или вывод (PayOut) средств со счета, удалите ненужный вариант. Например, если Вы хотите видеть отображение вывода средств со счёта, удалите " operationType == "PayIn" ||" из строки выше
    let [totalQuantity, totalSum, weigthedPrice] = _calculateTrades(trades)
    if (operationType == "Sell") {
      totalQuantity = -totalQuantity
      totalSum = -totalSum
      commission.value = -commission.value
    }
    if (operationType == "Tax" || operationType == "TaxDividend" || operationType == "Dividend"){
      totalQuantity = '-'
      weigthedPrice = '-'
    }
    var ticker
    if (!figi){
      ticker = ""
    } else {
      var ticker = tinkoffClient.getTickerByFigi(figi)
    }
    values.push([isoToDate(date), ticker, operationType, totalQuantity, weigthedPrice, payment, currency])
  }
  return values
}

function getPortfolio(refresh){
  const portfolio = tinkoffClient.getPort()
  const values = []
  values.push(["Тикер", "Название", "Кол-во", "Покупка", "Текущая", "Валюта"])
  for (let i=portfolio.length-1; i>=0; i--) {
    const {ticker, name, balance, averagePositionPrice, expectedYield} = portfolio [i]
    buy_price = averagePositionPrice.value * balance
    values.push([
      ticker, name, balance, buy_price, buy_price + expectedYield.value, averagePositionPrice.currency
    ])
  }
  return values
}

function getCurrencies(refresh){
  const values = []
  const portcur = tinkoffClient.getCur()
  for (let i=portcur.length-1; i>=0; i--) {
    const {currency, balance} = portcur [i]
    values.push([currency,balance])
    }
  return values
}

function getUSDval(refresh){
  return tinkoffClient.usdval()
}

function getIISPort(refresh){
  const users = tinkoffClient.getIISid()
  for (let i=users.length-1; i>=0; i--) {
    const {brokerAccountId, brokerAccountType} = users [i]
    if (brokerAccountType == "TinkoffIis")
      IISid = brokerAccountId
  }
  const portfolio = tinkoffClient.getIIS(IISid)
  const values = []
  values.push(["Тикер", "Название", "Кол-во", "Покупка", "Текущая", "Валюта"])
  for (let i=portfolio.length-1; i>=0; i--) {
    const {ticker, name, balance, averagePositionPrice, expectedYield} = portfolio [i]
    buy_price = averagePositionPrice.value * balance
    values.push([
      ticker, name, balance, buy_price, buy_price + expectedYield.value, averagePositionPrice.currency
    ])
  }
  return values
}
