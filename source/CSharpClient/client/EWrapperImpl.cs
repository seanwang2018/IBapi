/* Copyright (C) 2018 Interactive Brokers LLC. All rights reserved. This code is subject to the terms
 * and conditions of the IB API Non-Commercial License or the IB API Commercial License, as applicable. */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IBApi

{
    //! [ewrapperimpl]
    public class EWrapperImpl : EWrapper 
    {
    //! [ewrapperimpl]
        private int nextOrderId;
        //! [socket_declare]
        EClientSocket clientSocket;
        public readonly EReaderSignal Signal;
        //! [socket_declare]

        //! [socket_init]
        public EWrapperImpl()
        {
            Signal = new EReaderMonitorSignal();
            clientSocket = new EClientSocket(this, Signal);
        }
        //! [socket_init]

        public EClientSocket ClientSocket
        {
            get { return clientSocket; }
            set { clientSocket = value; }
        }

        public int NextOrderId
        {
            get { return nextOrderId; }
            set { nextOrderId = value; }
        }

        public string BboExchange { get; private set; }

        public virtual void error(Exception e)
        {
            Console.WriteLine("Exception thrown: "+e);
            throw e;
        }
        
        public virtual void error(string str)
        {
            Console.WriteLine("Error: "+str+"\n");
        }
        
        //! [error]
        public virtual void error(int id, int errorCode, string errorMsg)
        {
            Console.WriteLine("Error. Id: " + id + ", Code: " + errorCode + ", Msg: " + errorMsg + "\n");
            ErrorCode = errorCode; 
        }
        //! [error]

        public virtual void connectionClosed()
        {
            Console.WriteLine("Connection closed.\n");
        }
        
        public virtual void currentTime(long time) 
        {
            Console.WriteLine("Current Timeeeeee: "+ time+"\n");
        }

        //! [tickprice]
        public virtual void tickPrice(int tickerId, int field, double price, TickAttrib attribs) 
        {
            //  Console.WriteLine("Tick Priceeee. Ticker Id:"+tickerId+", Field: "+field+", Price: "+price+", CanAutoExecute: "+attribs.CanAutoExecute + 
            //       ", PastLimit: " + attribs.PastLimit + ", PreOpen: " + attribs.PreOpen);

            if ((field == 4) || (field==68))   mkt_price = price; // tickType 4 is last, ticktype 68 is delayed last           
            if (field == 2) ask_price = price;  // check TickType.cs;  type 2 is ask; type 1 is bid
            if (field == 1) bid_price = price;
        }
        //! [tickprice]
        
        //! [ticksize]
        public virtual void tickSize(int tickerId, int field, int size)
        {
         //   Console.WriteLine("Tick Size. Ticker Id:" + tickerId + ", Field: " + field + ", Size: " + size);
        }
        //! [ticksize]
        
        //! [tickstring]
        public virtual void tickString(int tickerId, int tickType, string value)
        {
          //  Console.WriteLine("Tick string. Ticker Id:" + tickerId + ", Type: " + tickType + ", Value: " + value);
        }
        //! [tickstring]

        //! [tickgeneric]
        public virtual void tickGeneric(int tickerId, int field, double value)
        {
          //  Console.WriteLine("Tick Generic. Ticker Id:" + tickerId + ", Field: " + field + ", Value: " + value);
        }
        //! [tickgeneric]

        public virtual void tickEFP(int tickerId, int tickType, double basisPoints, string formattedBasisPoints, double impliedFuture, int holdDays, string futureLastTradeDate, double dividendImpact, double dividendsToLastTradeDate)
        {
            Console.WriteLine("TickEFP. "+tickerId+", Type: "+tickType+", BasisPoints: "+basisPoints+", FormattedBasisPoints: "+formattedBasisPoints+", ImpliedFuture: "+impliedFuture+", HoldDays: "+holdDays+", FutureLastTradeDate: "+futureLastTradeDate+", DividendImpact: "+dividendImpact+", DividendsToLastTradeDate: "+dividendsToLastTradeDate);
        }
        
        //! [ticksnapshotend]
        public virtual void tickSnapshotEnd(int tickerId)
        {
            Console.WriteLine("TickSnapshotEnd: "+tickerId);
        }
        //! [ticksnapshotend]

        //! [nextvalidid]
        public virtual void nextValidId(int orderId) 
        {
            Console.WriteLine("Next Valid Id: "+orderId);
            NextOrderId = orderId;
        }
        //! [nextvalidid]

        //! [deltaneutralvalidation]
        public virtual void deltaNeutralValidation(int reqId, DeltaNeutralContract deltaNeutralContract)
        {
            Console.WriteLine("DeltaNeutralValidation. " + reqId + ", ConId: " + deltaNeutralContract.ConId + ", Delta: " + deltaNeutralContract.Delta + ", Price: " + deltaNeutralContract.Price);
        }
        //! [deltaneutralvalidation]

        //! [managedaccounts]
        public virtual void managedAccounts(string accountsList) 
        {
          //  Console.WriteLine("Account lists: "+accountsList);
            account_ID = accountsList;
        }
        //! [managedaccounts]

        //! [tickoptioncomputation]
        public virtual void tickOptionComputation(int tickerId, int field, double impliedVolatility, double delta, double optPrice, double pvDividend, double gamma, double vega, double theta, double undPrice)
        {
          //  Console.WriteLine("TickOptionComputation. TickerId: "+tickerId+", field: "+field+", ImpliedVolatility: "+impliedVolatility+", Delta: "+delta
          //      +", OptionPrice: "+optPrice+", pvDividend: "+pvDividend+", Gamma: "+gamma+", Vega: "+vega+", Theta: "+theta+", UnderlyingPrice: "+undPrice);
        }
        //! [tickoptioncomputation]

        //! [accountsummary]
        public virtual void accountSummary(int reqId, string account, string tag, string value, string currency)
        {
            // Console.WriteLine("Acct Summary. ReqId: " + reqId + ", Acct: " + account + ", Tag: " + tag + ", Value: " + value + ", Currency: " + currency);

            if (tag == "TotalCashValue") account_value = value;
            if (tag == "BuyingPower") account_BuyingPower = value;
            if (tag == "InitMarginReq") account_InitMarginReq = value;
            if (tag == "MaintMarginReq") account_MaintMarginReq = value;
            if (tag == "ExcessLiquidity") account_ExcessLiquidity = value;
            if (tag == "AvailableFunds") account_AvailableFunds = value;

    }
        //! [accountsummary]

        //! [accountsummaryend]
        public virtual void accountSummaryEnd(int reqId)
        {
            Console.WriteLine("AccountSummaryEnd. Req Id: "+reqId+"\n");
        }
        //! [accountsummaryend]

        //! [updateaccountvalue]
        public virtual void updateAccountValue(string key, string value, string currency, string accountName)
        {
            Console.WriteLine("UpdateAccountValue. Key: " + key + ", Value: " + value + ", Currency: " + currency + ", AccountName: " + accountName);
        }
        //! [updateaccountvalue]

        //! [updateportfolio]
        public virtual void updatePortfolio(Contract contract, double position, double marketPrice, double marketValue, double averageCost, double unrealizedPNL, double realizedPNL, string accountName)
        {
            Console.WriteLine("UpdatePortfolio. "+contract.Symbol+", "+contract.SecType+" @ "+contract.Exchange
                +": Position: "+position+", MarketPrice: "+marketPrice+", MarketValue: "+marketValue+", AverageCost: "+averageCost
                +", UnrealizedPNL: "+unrealizedPNL+", RealizedPNL: "+realizedPNL+", AccountName: "+accountName);
        }
        //! [updateportfolio]

        //! [updateaccounttime]
        public virtual void updateAccountTime(string timestamp)
        {
            Console.WriteLine("UpdateAccountTime. Time: " + timestamp+"\n");
        }
        //! [updateaccounttime]

        //! [accountdownloadend]
        public virtual void accountDownloadEnd(string account)
        {
            Console.WriteLine("Account download finished: "+account+"\n");
        }
        //! [accountdownloadend]

        //! [orderstatus]
        public virtual void orderStatus(int orderId, string status, double filled, double remaining, double avgFillPrice, int permId, int parentId, double lastFillPrice, int clientId, string whyHeld, double mktCapPrice)
        {
            Console.WriteLine("OrderStatus. Id: "+orderId+", Status: "+status+", Filled"+filled+", Remaining: "+remaining
                + ", AvgFillPrice: " + avgFillPrice + ", PermId: " + permId + ", LastFillPrice: " + lastFillPrice + ", ClientId: " + clientId + ", MktCapPrice: " + mktCapPrice);
            if (currOrderId == orderId)
                remainingOrderSize = (int)remaining;
        }
        //! [orderstatus]

        //! [openorder]
        public virtual void openOrder(int orderId, Contract contract, Order order, OrderState orderState)
        {
          //  Console.WriteLine("OpenOrder. ID: "+orderId+", "+contract.Symbol+", "+contract.SecType + " @ "+contract.Strike+": "+order.Action+", "+order.OrderType+" "+order.TotalQuantity+", "+orderState.Status);
           // if (contract.Strike == 0)
         //  Console.WriteLine("Order Details. " + orderId + " Symbol:  " + contract.Symbol + ", Leg: " + contract.ComboLegsDescription + ", Action: " + order.Action + ", order type: " + order.OrderType + ", total:  " + order.TotalQuantity + ", status: " + orderState.Status);
          

               if (order.WhatIf)
            //   if (currOrderId == orderId)
                 {
                Console.WriteLine("What-If. ID: " + orderId + 
                    ", InitMarginBefore: " + Util.formatDoubleString(orderState.InitMarginBefore) + ", MaintMarginBefore: " + Util.formatDoubleString(orderState.MaintMarginBefore) + " EquityWithLoanBefore: " + Util.formatDoubleString(orderState.EquityWithLoanBefore) +
                    ", InitMarginChange: " + Util.formatDoubleString(orderState.InitMarginChange) + ", MaintMarginChange: " + Util.formatDoubleString(orderState.MaintMarginChange) + " EquityWithLoanChange: " + Util.formatDoubleString(orderState.EquityWithLoanChange) +
                    ", InitMarginAfter: " + Util.formatDoubleString(orderState.InitMarginAfter) + ", MaintMarginAfter: " + Util.formatDoubleString(orderState.MaintMarginAfter) + " EquityWithLoanAfter: " + Util.formatDoubleString(orderState.EquityWithLoanAfter));
                InitMarginChange = Convert.ToDouble(orderState.InitMarginChange);
                MaintMarginChange = Convert.ToDouble(orderState.MaintMarginChange);
                 } 

        }
        //! [openorder]

        //! [openorderend]

        public virtual void openOrderEnd()
        {
            Console.WriteLine("OpenOrderEnd");
            checkOrderEnd = true;
        }
        //! [openorderend]

        //! [contractdetails]
        public virtual void contractDetails(int reqId, ContractDetails contractDetails)
        {
            Console.WriteLine("ContractDetails begin. ReqId: " + reqId);
             //  printContractMsg(contractDetails.Contract);
            //   printContractDetailsMsg(contractDetails);
            contract_ID = contractDetails.Contract.ConId;
           // Console.WriteLine("ContractDetails end. ReqId: " + reqId);
        }
        //! [contractdetails]

        public void printContractMsg(Contract contract)
        {
            Console.WriteLine("\tConId: " + contract.ConId);
            Console.WriteLine("\tSymbol: " + contract.Symbol);
            Console.WriteLine("\tSecType: " + contract.SecType);
            Console.WriteLine("\tLastTradeDateOrContractMonth: " + contract.LastTradeDateOrContractMonth);
            Console.WriteLine("\tStrike: " + contract.Strike);
            Console.WriteLine("\tRight: " + contract.Right);
            Console.WriteLine("\tMultiplier: " + contract.Multiplier);
            Console.WriteLine("\tExchange: " + contract.Exchange);
            Console.WriteLine("\tPrimaryExchange: " + contract.PrimaryExch);
            Console.WriteLine("\tCurrency: " + contract.Currency);
            Console.WriteLine("\tLocalSymbol: " + contract.LocalSymbol);
            Console.WriteLine("\tTradingClass: " + contract.TradingClass);
        }

        public void printContractDetailsMsg(ContractDetails contractDetails)
        {
            Console.WriteLine("\tMarketName: " + contractDetails.MarketName);
            Console.WriteLine("\tMinTick: " + contractDetails.MinTick);
            Console.WriteLine("\tPriceMagnifier: " + contractDetails.PriceMagnifier);
            Console.WriteLine("\tOrderTypes: " + contractDetails.OrderTypes);
            Console.WriteLine("\tValidExchanges: " + contractDetails.ValidExchanges);
            Console.WriteLine("\tUnderConId: " + contractDetails.UnderConId);
            Console.WriteLine("\tLongName: " + contractDetails.LongName);
            Console.WriteLine("\tContractMonth: " + contractDetails.ContractMonth);
            Console.WriteLine("\tIndystry: " + contractDetails.Industry);
            Console.WriteLine("\tCategory: " + contractDetails.Category);
            Console.WriteLine("\tSubCategory: " + contractDetails.Subcategory);
            Console.WriteLine("\tTimeZoneId: " + contractDetails.TimeZoneId);
            Console.WriteLine("\tTradingHours: " + contractDetails.TradingHours);
            Console.WriteLine("\tLiquidHours: " + contractDetails.LiquidHours);
            Console.WriteLine("\tEvRule: " + contractDetails.EvRule);
            Console.WriteLine("\tEvMultiplier: " + contractDetails.EvMultiplier);
            Console.WriteLine("\tMdSizeMultiplier: " + contractDetails.MdSizeMultiplier);
            Console.WriteLine("\tAggGroup: " + contractDetails.AggGroup);
            Console.WriteLine("\tUnderSymbol: " + contractDetails.UnderSymbol);
            Console.WriteLine("\tUnderSecType: " + contractDetails.UnderSecType);
            Console.WriteLine("\tMarketRuleIds: " + contractDetails.MarketRuleIds);
            Console.WriteLine("\tRealExpirationDate: " + contractDetails.RealExpirationDate);
            Console.WriteLine("\tLastTradeTime: " + contractDetails.LastTradeTime);
            printContractDetailsSecIdList(contractDetails.SecIdList);
        }

        public void printContractDetailsSecIdList(List<TagValue> secIdList)
        {
            if (secIdList != null && secIdList.Count > 0) {
                Console.Write("\tSecIdList: {");
                foreach (TagValue tagValue in secIdList)
                {
                    Console.Write(tagValue.Tag + "=" + tagValue.Value + ";");
                }
                Console.WriteLine("}");
            }
        }

        public void printBondContractDetailsMsg(ContractDetails contractDetails)
        {
            Console.WriteLine("\tSymbol: " + contractDetails.Contract.Symbol);
            Console.WriteLine("\tSecType: " + contractDetails.Contract.SecType);
            Console.WriteLine("\tCusip: " + contractDetails.Cusip);
            Console.WriteLine("\tCoupon: " + contractDetails.Coupon);
            Console.WriteLine("\tMaturity: " + contractDetails.Maturity);
            Console.WriteLine("\tIssueDate: " + contractDetails.IssueDate);
            Console.WriteLine("\tRatings: " + contractDetails.Ratings);
            Console.WriteLine("\tBondType: " + contractDetails.BondType);
            Console.WriteLine("\tCouponType: " + contractDetails.CouponType);
            Console.WriteLine("\tConvertible: " + contractDetails.Convertible);
            Console.WriteLine("\tCallable: " + contractDetails.Callable);
            Console.WriteLine("\tPutable: " + contractDetails.Putable);
            Console.WriteLine("\tDescAppend: " + contractDetails.DescAppend);
            Console.WriteLine("\tExchange: " + contractDetails.Contract.Exchange);
            Console.WriteLine("\tCurrency: " + contractDetails.Contract.Currency);
            Console.WriteLine("\tMarketName: " + contractDetails.MarketName);
            Console.WriteLine("\tTradingClass: " + contractDetails.Contract.TradingClass);
            Console.WriteLine("\tConId: " + contractDetails.Contract.ConId);
            Console.WriteLine("\tMinTick: " + contractDetails.MinTick);
            Console.WriteLine("\tMdSizeMultiplier: " + contractDetails.MdSizeMultiplier);
            Console.WriteLine("\tOrderTypes: " + contractDetails.OrderTypes);
            Console.WriteLine("\tValidExchanges: " + contractDetails.ValidExchanges);
            Console.WriteLine("\tNextOptionDate: " + contractDetails.NextOptionDate);
            Console.WriteLine("\tNextOptionType: " + contractDetails.NextOptionType);
            Console.WriteLine("\tNextOptionPartial: " + contractDetails.NextOptionPartial);
            Console.WriteLine("\tNotes: " + contractDetails.Notes);
            Console.WriteLine("\tLong Name: " + contractDetails.LongName);
            Console.WriteLine("\tEvRule: " + contractDetails.EvRule);
            Console.WriteLine("\tEvMultiplier: " + contractDetails.EvMultiplier);
            Console.WriteLine("\tAggGroup: " + contractDetails.AggGroup);
            Console.WriteLine("\tMarketRuleIds: " + contractDetails.MarketRuleIds);
            Console.WriteLine("\tLastTradeTime: " + contractDetails.LastTradeTime);
            Console.WriteLine("\tTimeZoneId: " + contractDetails.TimeZoneId);
            printContractDetailsSecIdList(contractDetails.SecIdList);
        }


        //! [contractdetailsend]
        public virtual void contractDetailsEnd(int reqId)
        {
            Console.WriteLine("ContractDetailsEnd. "+reqId+"\n");
            checkContractEnd = true;
        }
        //! [contractdetailsend]

        //! [execdetails]
        public virtual void execDetails(int reqId, Contract contract, Execution execution)
        {
             // Console.WriteLine("ExecDetails. " + reqId + " - " + contract.Symbol + " - " + execution.LastLiquidity);
             if(contract.Strike == 0)
                Console.WriteLine("ExecDetails. "+reqId+" - "+contract.LastTradeDateOrContractMonth+", " + contract.ComboLegsDescription + ", " + contract.Symbol +", "+contract.Strike+ ", orderID: "+execution.OrderId+", time: "+execution.Time+", " + execution.Side + ", ExecId: " + execution.ExecId+", " + execution.Price + ", PermId: "+execution.PermId+", Orderref: "+execution.OrderRef+", " + execution.Shares);
        }
        //! [execdetails]

        //! [execdetailsend]
        public virtual void execDetailsEnd(int reqId)
        {
            Console.WriteLine("ExecDetailsEnd. "+reqId+"\n");
        }
        //! [execdetailsend]

        //! [commissionreport]
        public virtual void commissionReport(CommissionReport commissionReport)
        {
            Console.WriteLine("CommissionReport. "+commissionReport.ExecId+" - "+commissionReport.Commission+" "+commissionReport.Currency+" RPNL "+commissionReport.RealizedPNL);
        }
        //! [commissionreport]

        //! [fundamentaldata]
        public virtual void fundamentalData(int reqId, string data)
        {
            Console.WriteLine("FundamentalData. " + reqId + "" + data+"\n");
        }
        //! [fundamentaldata]

        //! [historicaldata]
        public virtual void historicalData(int reqId, Bar bar)
        {
            Console.WriteLine("HistoricalData. " + reqId + " - Time: " + bar.Time + ", Open: " + bar.Open + ", High: " + bar.High + ", Low: " + bar.Low + ", Close: " + bar.Close + ", Volume: " + bar.Volume + ", Count: " + bar.Count + ", WAP: " + bar.WAP);
            mkt_price = bar.Close;
        }
        //! [historicaldata]

        //! [marketdatatype]
        public virtual void marketDataType(int reqId, int marketDataType)
        {
            Console.WriteLine("MarketDataType. "+reqId+", Type: "+marketDataType+"\n");
        }
        //! [marketdatatype]

        //! [updatemktdepth]
        public virtual void updateMktDepth(int tickerId, int position, int operation, int side, double price, int size)
        {
            Console.WriteLine("UpdateMarketDepth. " + tickerId + " - Position: " + position + ", Operation: " + operation + ", Side: " + side + ", Price: " + price + ", Size" + size);
        }
        //! [updatemktdepth]

        //! [updatemktdepthl2]
        public virtual void updateMktDepthL2(int tickerId, int position, string marketMaker, int operation, int side, double price, int size)
        {
            Console.WriteLine("UpdateMarketDepthL2. " + tickerId + " - Position: " + position + ", Operation: " + operation + ", Side: " + side + ", Price: " + price + ", Size" + size);
        }
        //! [updatemktdepthl2]

        //! [updatenewsbulletin]
        public virtual void updateNewsBulletin(int msgId, int msgType, String message, String origExchange)
        {
            Console.WriteLine("News Bulletins. "+msgId+" - Type: "+msgType+", Message: "+message+", Exchange of Origin: "+origExchange+"\n");
        }
        //! [updatenewsbulletin]

        //! [position]
        public virtual void position(string account, Contract contract, double pos, double avgCost)
        {
            Console.WriteLine("Position. "+account+" - Symbol: "+contract.Symbol+", SecType: "+contract.SecType+", Currency: "+contract.Currency+", Position: "+pos+", Avg cost: "+avgCost);
        }
        //! [position]

        //! [positionend]
        public virtual void positionEnd()
        {
            Console.WriteLine("PositionEnd \n");
        }
        //! [positionend]

        //! [realtimebar]
        public virtual void realtimeBar(int reqId, long time, double open, double high, double low, double close, long volume, double WAP, int count)
        {
          //  Console.WriteLine("RealTimeBars. " + reqId + " - Time: " + time + ", Open: " + open + ", High: " + high + ", Low: " + low + ", Close: " + close + ", Volume: " + volume + ", Count: " + count + ", WAP: " + WAP);
            RT_price = close;

        }
        //! [realtimebar]

        //! [scannerparameters]
        public virtual void scannerParameters(string xml)
        {
            Console.WriteLine("ScannerParameters. "+xml+"\n");
        }
        //! [scannerparameters]

        //! [scannerdata]
        public virtual void scannerData(int reqId, int rank, ContractDetails contractDetails, string distance, string benchmark, string projection, string legsStr)
        {
            Console.WriteLine("ScannerData. "+reqId+" - Rank: "+rank+", Symbol: "+contractDetails.Contract.Symbol+", SecType: "+contractDetails.Contract.SecType+", Currency: "+contractDetails.Contract.Currency
                +", Distance: "+distance+", Benchmark: "+benchmark+", Projection: "+projection+", Legs String: "+legsStr);
        }
        //! [scannerdata]

        //! [scannerdataend]
        public virtual void scannerDataEnd(int reqId)
        {
            Console.WriteLine("ScannerDataEnd. "+reqId);
        }
        //! [scannerdataend]

        //! [receivefa]
        public virtual void receiveFA(int faDataType, string faXmlData)
        {
            Console.WriteLine("Receing FA: "+faDataType+" - "+faXmlData);
        }
        //! [receivefa]

        public virtual void bondContractDetails(int requestId, ContractDetails contractDetails)
        {
            Console.WriteLine("BondContractDetails begin. ReqId: " + requestId);
            printBondContractDetailsMsg(contractDetails);
            Console.WriteLine("BondContractDetails end. ReqId: " + requestId);
        }

        //! [historicaldataend]
        public virtual void historicalDataEnd(int reqId, string startDate, string endDate)
        {
            Console.WriteLine("HistoricalDataEnd - "+reqId+" from "+startDate+" to "+endDate);
            checkDataEnd = true;
        }
        //! [historicaldataend]

        public virtual void verifyMessageAPI(string apiData)
        {
            Console.WriteLine("verifyMessageAPI: " + apiData);
        }
        public virtual void verifyCompleted(bool isSuccessful, string errorText)
        {
            Console.WriteLine("verifyCompleted. IsSuccessfule: " + isSuccessful + " - Error: " + errorText);
        }
        public virtual void verifyAndAuthMessageAPI(string apiData, string xyzChallenge)
        {
            Console.WriteLine("verifyAndAuthMessageAPI: " + apiData + " " + xyzChallenge);
        }
        public virtual void verifyAndAuthCompleted(bool isSuccessful, string errorText)
        {
            Console.WriteLine("verifyAndAuthCompleted. IsSuccessful: " + isSuccessful + " - Error: " + errorText);
        }
        //! [displaygrouplist]
        public virtual void displayGroupList(int reqId, string groups)
        {
            Console.WriteLine("DisplayGroupList. Request: " + reqId + ", Groups" + groups);
        }
        //! [displaygrouplist]

        //! [displaygroupupdated]
        public virtual void displayGroupUpdated(int reqId, string contractInfo)
        {
            Console.WriteLine("displayGroupUpdated. Request: " + reqId + ", ContractInfo: " + contractInfo);
        }
        //! [displaygroupupdated]

        //! [positionmulti]
        public virtual void positionMulti(int reqId, string account, string modelCode, Contract contract, double pos, double avgCost)
        {
            Console.WriteLine("Position Multi. Request: " + reqId + ", Account: " + account + ", ModelCode: " + modelCode + ", Symbol: " + contract.Symbol + ", SecType: " + contract.SecType + ", Currency: " + contract.Currency + ", Position: " + pos + ", Avg cost: " + avgCost + "\n");
        }
        //! [positionmulti]

        //! [positionmultiend]
        public virtual void positionMultiEnd(int reqId)
        {
            Console.WriteLine("Position Multi End. Request: " + reqId + "\n");
        }
        //! [positionmultiend]

        //! [accountupdatemulti]
        public virtual void accountUpdateMulti(int reqId, string account, string modelCode, string key, string value, string currency)
        {
            Console.WriteLine("Account Update Multi. Request: " + reqId + ", Account: " + account + ", ModelCode: " + modelCode + ", Key: " + key + ", Value: " + value + ", Currency: " + currency + "\n");
        }
        //! [accountupdatemulti]

        //! [accountupdatemultiend]
        public virtual void accountUpdateMultiEnd(int reqId)
        {
            Console.WriteLine("Account Update Multi End. Request: " + reqId + "\n");
        }
        //! [accountupdatemultiend]

        //! [securityDefinitionOptionParameter]
        public void securityDefinitionOptionParameter(int reqId, string exchange, int underlyingConId, string tradingClass, string multiplier, HashSet<string> expirations, HashSet<double> strikes)
        {
            Console.WriteLine("Security Definition Option Parameter. Reqest: {0}, Exchange: {1}, Undrelying contract id: {2}, Trading class: {3}, Multiplier: {4}, Expirations: {5}, Strikes: {6}",
                              reqId, exchange, underlyingConId, tradingClass, multiplier, string.Join(", ", expirations), string.Join(", ", strikes));
        }
        //! [securityDefinitionOptionParameter]

        //! [securityDefinitionOptionParameterEnd]
        public void securityDefinitionOptionParameterEnd(int reqId)
        {
            Console.WriteLine("Security Definition Option Parameter End. Request: " + reqId + "\n");
        }
        //! [securityDefinitionOptionParameterEnd]

        //! [connectack]
        public void connectAck()
        {
            if (ClientSocket.AsyncEConnect)
                ClientSocket.startApi();
        }
        //! [connectack]

        //! [softDollarTiers]
        public void softDollarTiers(int reqId, SoftDollarTier[] tiers)
        {
            Console.WriteLine("Soft Dollar Tiers:");

            foreach (var tier in tiers)
            {
                Console.WriteLine(tier.DisplayName);
            }
        }
        //! [softDollarTiers]

        //! [familyCodes]
        public void familyCodes(FamilyCode[] familyCodes)
        {
            Console.WriteLine("Family Codes:");

            foreach (var familyCode in familyCodes)
            {
                Console.WriteLine("Account ID: {0}, Family Code Str: {1}", familyCode.AccountID, familyCode.FamilyCodeStr);
            }
        }
        //! [familyCodes]

        //! [symbolSamples]
        public void symbolSamples(int reqId, ContractDescription[] contractDescriptions) 
        {
            string derivSecTypes;
            Console.WriteLine("Symbol Samples. Request Id: {0}", reqId);

            foreach (var contractDescription in contractDescriptions)
            {
                derivSecTypes = "";
                foreach (var derivSecType in contractDescription.DerivativeSecTypes)
                {
                    derivSecTypes += derivSecType;
                    derivSecTypes += " ";
                }
                Console.WriteLine("Contract: conId - {0}, symbol - {1}, secType - {2}, primExchange - {3}, currency - {4}, derivativeSecTypes - {5}", 
                    contractDescription.Contract.ConId, contractDescription.Contract.Symbol, contractDescription.Contract.SecType, 
                    contractDescription.Contract.PrimaryExch, contractDescription.Contract.Currency, derivSecTypes);
            }
        }
        //! [symbolSamples]

        //! [mktDepthExchanges]
        public void mktDepthExchanges(DepthMktDataDescription[] depthMktDataDescriptions)
        {
            Console.WriteLine("Market Depth Exchanges:");

            foreach (var depthMktDataDescription in depthMktDataDescriptions)
            {
                Console.WriteLine("Depth Market Data Description: Exchange: {0}, Security Type: {1}, Listing Exch: {2}, Service Data Type: {3}, Agg Group: {4}", 
                    depthMktDataDescription.Exchange, depthMktDataDescription.SecType, 
                    depthMktDataDescription.ListingExch, depthMktDataDescription.ServiceDataType,
                    depthMktDataDescription.AggGroup != Int32.MaxValue ? depthMktDataDescription.AggGroup.ToString() : ""
                    );
            }
        }
        //! [mktDepthExchanges]

        //! [tickNews]
        public void tickNews(int tickerId, long timeStamp, string providerCode, string articleId, string headline, string extraData)
        {
            Console.WriteLine("Tick News. Ticker Id: {0}, Time Stamp: {1}, Provider Code: {2}, Article Id: {3}, headline: {4}, extraData: {5}", tickerId, timeStamp, providerCode, articleId, headline, extraData);
        }
        //! [tickNews]

        //! [smartcomponents]
        public void smartComponents(int reqId, Dictionary<int, KeyValuePair<string, char>> theMap)
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendFormat("==== Smart Components Begin (total={0}) reqId = {1} ====\n", theMap.Count, reqId);

            foreach (var item in theMap)
            {
                sb.AppendFormat("bit number: {0}, exchange: {1}, exchange letter: {2}\n", item.Key, item.Value.Key, item.Value.Value);
            }

            sb.AppendFormat("==== Smart Components Begin (total={0}) reqId = {1} ====\n", theMap.Count, reqId);

            Console.WriteLine(sb);
        }
        //! [smartcomponents]

        //! [tickReqParams]
        public void tickReqParams(int tickerId, double minTick, string bboExchange, int snapshotPermissions)
        {
        //    Console.WriteLine("id={0} minTick = {1} bboExchange = {2} snapshotPermissions = {3}", tickerId, minTick, bboExchange, snapshotPermissions);

            BboExchange = bboExchange;
        }
        //! [tickReqParams]

        //! [newsProviders]
        public void newsProviders(NewsProvider[] newsProviders)
        {
            Console.WriteLine("News Providers:");

            foreach (var newsProvider in newsProviders)
            {
                Console.WriteLine("News provider: providerCode - {0}, providerName - {1}",
                    newsProvider.ProviderCode, newsProvider.ProviderName);
            }
        }
        //! [newsProviders]

        //! [newsArticle]
        public void newsArticle(int requestId, int articleType, string articleText)
        {
            Console.WriteLine("News Article. Request Id: {0}, ArticleType: {1}", requestId, articleType);
            if (articleType == 0) {
                Console.WriteLine("News Article Text: {0}", articleText);
            }
            else if (articleType == 1) {
                Console.WriteLine("News Article Text: article text is binary/pdf and cannot be displayed");
            }
        }
        //! [newsArticle]
        
        //! [historicalNews]
        public void historicalNews(int requestId, string time, string providerCode, string articleId, string headline)
        {
            Console.WriteLine("Historical News. Request Id: {0}, Time: {1}, Provider Code: {2}, Article Id: {3}, headline: {4}", requestId, time, providerCode, articleId, headline);
        }
        //! [historicalNews]

        //! [historicalNewsEnd]
        public void historicalNewsEnd(int requestId, bool hasMore)
        {
            Console.WriteLine("Historical News End. Request Id: {0}, Has More: {1}", requestId, hasMore);
        }
        //! [historicalNewsEnd]

        //! [headTimestamp]
        public void headTimestamp(int reqId, string headTimestamp)
        {
            Console.WriteLine("Head time stamp. Request Id: {0}, Head time stamp: {1}", reqId, headTimestamp);
        }
        //! [headTimestamp]

        //! [histogramData]
        public void histogramData(int reqId, HistogramEntry[] data)
        {
            Console.WriteLine("Histogram data. Request Id: {0}, data size: {1}", reqId, data.Length);
            data.ToList().ForEach(i => Console.WriteLine("\tPrice: {0}, Size: {1}", i.Price, i.Size));
        }
        //! [histogramData]

        //! [historicalDataUpdate]
        public void historicalDataUpdate(int reqId, Bar bar)
        {
            Console.WriteLine("HistoricalDataUpdate. " + reqId + " - Time: " + bar.Time + ", Open: " + bar.Open + ", High: " + bar.High + ", Low: " + bar.Low + ", Close: " + bar.Close + ", Volume: " + bar.Volume + ", Count: " + bar.Count + ", WAP: " + bar.WAP);
        }
        //! [historicalDataUpdate]

        //! [rerouteMktDataReq]
        public void rerouteMktDataReq(int reqId, int conId, string exchange)
        {
            Console.WriteLine("Re-route market data request. Req Id: {0}, ConId: {1}, Exchange: {2}", reqId, conId, exchange);
        }
        //! [rerouteMktDataReq]

        //! [rerouteMktDepthReq]
        public void rerouteMktDepthReq(int reqId, int conId, string exchange)
        {
            Console.WriteLine("Re-route market depth request. Req Id: {0}, ConId: {1}, Exchange: {2}", reqId, conId, exchange);
        }
        //! [rerouteMktDepthReq]

        //! [marketRule]
        public void marketRule(int marketRuleId, PriceIncrement[] priceIncrements) 
        {
            Console.WriteLine("Market Rule Id: " + marketRuleId);
            foreach (var priceIncrement in priceIncrements) 
            {
                Console.WriteLine("Low Edge: {0}, Increment: {1}", ((decimal)priceIncrement.LowEdge).ToString(), ((decimal)priceIncrement.Increment).ToString());
            }
        }
        //! [marketRule]

		//! [pnl]
        public void pnl(int reqId, double dailyPnL, double unrealizedPnL, double realizedPnL)
        {
            Console.WriteLine("PnL. Request Id: {0}, Daily PnL: {1}, Unrealized PnL: {2}, Realized PnL: {3}", reqId, dailyPnL, unrealizedPnL, realizedPnL);
        }
		//! [pnl]

		//! [pnlsingle]
        public void pnlSingle(int reqId, int pos, double dailyPnL, double unrealizedPnL, double realizedPnL, double value)
        {
            Console.WriteLine("PnL Single. Request Id: {0}, Pos {1}, Daily PnL {2}, Unrealized PnL {3}, Realized PnL: {4}, Value: {5}", reqId, pos, dailyPnL, unrealizedPnL, realizedPnL, value);
        }
		//! [pnlsingle]

		//! [historicalticks]
        public void historicalTicks(int reqId, HistoricalTick[] ticks, bool done)
        {
            foreach (var tick in ticks)
            {
                Console.WriteLine("Historical Tick. Request Id: {0}, Time: {1}, Price: {2}, Size: {3}", reqId, Util.UnixSecondsToString(tick.Time, "yyyyMMdd-HH:mm:ss zzz"), tick.Price, tick.Size);
            }
        }
		//! [historicalticks]

		//! [historicalticksbidask]
        public void historicalTicksBidAsk(int reqId, HistoricalTickBidAsk[] ticks, bool done)
        {
            foreach (var tick in ticks)
            {
                Console.WriteLine("Historical Tick Bid/Ask. Request Id: {0}, Time: {1}, Mask: {2} Price Bid: {3}, Price Ask {4}, Size Bid: {5}, Size Ask {6}",
                    reqId, Util.UnixSecondsToString(tick.Time, "yyyyMMdd-HH:mm:ss zzz"), tick.Mask, tick.PriceBid, tick.PriceAsk, tick.SizeBid, tick.SizeAsk);
            }
        }
		//! [historicalticksbidask]

		//! [historicaltickslast]
        public void historicalTicksLast(int reqId, HistoricalTickLast[] ticks, bool done)
        {
            foreach (var tick in ticks)
            {
                Console.WriteLine("Historical Tick Last. Request Id: {0}, Time: {1}, Mask: {2}, Price: {3}, Size: {4}, Exchange: {5}, Special Conditions: {6}",
                    reqId, Util.UnixSecondsToString(tick.Time, "yyyyMMdd-HH:mm:ss zzz"), tick.Mask, tick.Price, tick.Size, tick.Exchange, tick.SpecialConditions);
            }
        }
		//! [historicaltickslast]

        //! [tickbytickalllast]
        public void tickByTickAllLast(int reqId, int tickType, long time, double price, int size, TickAttrib attribs, string exchange, string specialConditions)
        {
            Console.WriteLine("Tick-By-Tick. Request Id: {0}, TickType: {1}, Time: {2}, Price: {3}, Size: {4}, Exchange: {5}, Special Conditions: {6}, PastLimit: {7}, Unreported: {8}",
                reqId, tickType == 1 ? "Last" : "AllLast", Util.UnixSecondsToString(time, "yyyyMMdd-HH:mm:ss zzz"), price, size, exchange, specialConditions, attribs.PastLimit, attribs.Unreported);
        }
        //! [tickbytickalllast]

        //! [tickbytickbidask]
        public void tickByTickBidAsk(int reqId, long time, double bidPrice, double askPrice, int bidSize, int askSize, TickAttrib attribs)
        {
            Console.WriteLine("Tick-By-Tick. Request Id: {0}, TickType: BidAsk, Time: {1}, BidPrice: {2}, AskPrice: {3}, BidSize: {4}, AskSize: {5}, BidPastLow: {6}, AskPastHigh: {7}",
                reqId, Util.UnixSecondsToString(time, "yyyyMMdd-HH:mm:ss zzz"), bidPrice, askPrice, bidSize, askSize, attribs.BidPastLow, attribs.AskPastHigh);
        }
        //! [tickbytickbidask]

        //! [tickbytickmidpoint]
        public void tickByTickMidPoint(int reqId, long time, double midPoint)
        {
            Console.WriteLine("Tick-By-Tick. Request Id: {0}, TickType: MidPoint, Time: {1}, MidPoint: {2}",
                reqId, Util.UnixSecondsToString(time, "yyyyMMdd-HH:mm:ss zzz"), midPoint);
        }
        //! [tickbytickmidpoint] 

        /* account information that I need */
        public string account_ID { get; private set; }
        public string account_value { get; private set; }
        public string account_BuyingPower { get; private set; }
        public string account_InitMarginReq { get; private set; }
        public string account_MaintMarginReq { get; private set; }
        public string account_ExcessLiquidity { get; private set; }
        public string account_AvailableFunds { get; private set; }


        /* market data varables */
        public double mkt_price { get; private set; }
        public double ask_price { get; private set; }
        public double bid_price { get; private set; }

        public double RT_price { get; private set; }

        /**** contract details ******/
        public int contract_ID { get; private set; }

        /**** order info **************/
        public int currOrderId { get; set; }
        public int remainingOrderSize { get; set; }
        public bool checkOrderEnd { get; set; }
        public int ErrorCode { get; set; }
        public bool checkContractEnd { get; set; }
        public bool checkDataEnd { get; set; }
        public double InitMarginChange { get; set; }
        public double MaintMarginChange { get; set; }


    }
}
