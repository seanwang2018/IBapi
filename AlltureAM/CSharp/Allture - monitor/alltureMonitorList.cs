/* Copyright (C) 2018 Interactive Brokers LLC. All rights reserved. This code is subject to the terms
 * and conditions of the IB API Non-Commercial License or the IB API Commercial License, as applicable. */
using System;
using System.Net.Mail;
using IBApi;
using System.Threading;
using Samples;
using System.Collections.Generic;
using _Excel = Microsoft.Office.Interop.Excel; 


namespace Allture
{
    public class Allture_Monitor
    {
        /* IMPORTANT: always use your paper trading account. The code below will submit orders as part of the demonstration. */
        /* IB will not be responsible for accidental executions on your live account. */
        /* Any stock or option symbols displayed are for illustrative purposes only and are not intended to portray a recommendation. */
        /* Before contacting our API support team please refer to the available documentation. */

        static EWrapperImpl Allture = new EWrapperImpl();
        static TradeBooks openBook = new TradeBooks();

        public static int sleep0 = 200;
        public static int sleep1 = 1000;
        public static int sleep2 = 2000;

        public static int Main(string[] args)
        {
            // specify account and strategy           
            string IBaccount = TradeBooks.accountInit();
            string tradingModel = "monitor";

            double indexPrice = 0, lastPrice = 0,  vixPrice = 0, vix1Price = 0, vix2Price = 0;
            int legB_ID, legS_ID;
            double spread_premium;

            //*** set the monitoring parameters  **********
            bool autoExit = false;

            double SPXexitPoint = 30;
            double ESexitPoint = 30;
            double CLexitPoint = 3;
            double GCexitPoint = 20;

            double exitPoint = 100;

            int Alert1 = 0;
            int Alert2 = 0;

            bool[] alert1 = new bool[Constants.max_position];
            bool[] alert2 = new bool[Constants.max_position];
            bool[] alert3 = new bool[Constants.max_position];
            bool alert_vix = false; // vix/vix1 backwardation
            bool alert_vix1 = false; // vix1/vix2 backwardation

            List<double> putSPXlegS = new List<double>();
            List<double> callSPXlegS = new List<double>();

            List<double> putESlegS = new List<double>();
            List<double> callESlegS = new List<double>();

            List<double> putCLlegS = new List<double>();
            List<double> callCLlegS = new List<double>();

            List<double> putGClegS = new List<double>();
            List<double> callGClegS = new List<double>();

            //****  current date and time *****
            string date_time, today;
            date_time = DateTime.Now.ToString("h:mm:ss tt");
            today = DateTime.Now.ToString("MMdd");
          //   today = "0928";  // if today is not the trading close day

            //***** set up excel **************

            _Excel.Application excel = new _Excel.Application();
            excel.DisplayAlerts = false;
             _Excel.Workbook monitor_wb = excel.Workbooks.Open("C:\\Users\\Jack\\Documents\\Companies\\allture\\Trading\\option_spread\\Accounts\\All_Account_Oct_2018.xlsm");
             _Excel.Worksheet csheet = monitor_wb.Worksheets[today];
           // string path = TradeBooks.xlsPathSetup(IBaccount);
           // _Excel.Workbook monitor_wb = excel.Workbooks.Open(path + IBaccount + "\\account_" + IBaccount + ".xlsm");
           // _Excel.Worksheet csheet = monitor_wb.Worksheets["Positions"];


            Console.WriteLine("\n start calculating, time: " + DateTime.Now + " \n");

            //******* read position data to openBook
            position_xls_readin(csheet, openBook);

            monitor_wb.Close(0);
            excel.Quit();

            for (int n = 0; n < openBook.positionCNT; n++)
            {
                if (openBook.symbol[n] == "SPX")
                    if (openBook.CallPut[n] == "Put")
                        putSPXlegS.Add(openBook.legS[n]);
                    else callSPXlegS.Add(openBook.legS[n]);
                if (openBook.symbol[n] == "ES")
                    if (openBook.CallPut[n] == "Put")
                        putESlegS.Add(openBook.legS[n]);
                    else callESlegS.Add(openBook.legS[n]);
                if (openBook.symbol[n] == "CL")
                    if (openBook.CallPut[n] == "Put")
                        putCLlegS.Add(openBook.legS[n]);
                    else callCLlegS.Add(openBook.legS[n]);
                if (openBook.symbol[n] == "GC")
                    if (openBook.CallPut[n] == "Put")
                        putGClegS.Add(openBook.legS[n]);
                    else callGClegS.Add(openBook.legS[n]);
            }

            callSPXlegS.Sort((s1, s2) => s1.CompareTo(s2));
            callESlegS.Sort((s1, s2) => s1.CompareTo(s2));
            callCLlegS.Sort((s1, s2) => s1.CompareTo(s2));
            callGClegS.Sort((s1, s2) => s1.CompareTo(s2));

            putSPXlegS.Sort((s1, s2) => s2.CompareTo(s1));
            putESlegS.Sort((s1, s2) => s2.CompareTo(s1));
            putCLlegS.Sort((s1, s2) => s2.CompareTo(s1));
            putGClegS.Sort((s1, s2) => s2.CompareTo(s1));

            //!  start the connection ********************************//
            int[] channel = new int[2];
            channel = TradeBooks.channelSetup(IBaccount, tradingModel);

            EClientSocket clientSocket = Allture.ClientSocket;
            EReaderSignal readerSignal = Allture.Signal;
            //! [connect]
            clientSocket.eConnect("127.0.0.1", channel[0], channel[1]);

            //! [connect]
            //! [ereader]
            //Create a reader to consume messages from the TWS. The EReader will consume the incoming messages and put them in a queue
            var reader = new EReader(clientSocket, readerSignal);
            reader.Start();
            //Once the messages are in the queue, an additional thread can be created to fetch them
            new Thread(() => { while (clientSocket.IsConnected()) { readerSignal.waitForSignal(); reader.processMsgs(); } }) { IsBackground = true }.Start();
            //! [ereader]

            //**************************
            //*** Account Info       ***
            //**************************

            clientSocket.reqAccountSummary(9001, "All", AccountSummaryTags.GetAllTags());
            Thread.Sleep(1000);

            Console.WriteLine("account ID: " + Allture.account_ID);
            Console.WriteLine("account Value: " + Allture.account_value);
            Console.WriteLine("account_BuyingPower: " + Allture.account_BuyingPower);
            Console.WriteLine("account_ExcessLiquidity: " + Allture.account_ExcessLiquidity);
            Console.WriteLine("account_AvailableFunds: " + Allture.account_AvailableFunds);
            Console.WriteLine("\n");


            //******* calculate the adj std, return with spx, vix, vxx at order execuation in xls
            // runSPXstrategyCurr(trading_new, csheet, clientSocket);

            for (int n = 0; n < openBook.positionCNT; n++)
            {
                alert1[n] = false;
                alert2[n] = false;
                alert3[n] = false;
            }

            clientSocket.reqMarketDataType(1);
            Thread.Sleep(sleep1);

            while ((DateTime.Now > Convert.ToDateTime("09:30:00 AM")) && (DateTime.Now < Convert.ToDateTime("16:01:00 PM")))
            {
                do { indexPrice = getMktData(clientSocket, Allture, "SPX", "IND", "", "CBOE"); }
                while (indexPrice == lastPrice);
                lastPrice = indexPrice;  // to ensure vix, vix1, vix2 not consecutive,  since their prices aren't chaning very much each time, can't be consecutive for lastPrice. 

                do { vixPrice = getMktData(clientSocket, Allture, "VIX", "IND", "", "CBOE"); }
                while (vixPrice == lastPrice);
                lastPrice = vixPrice;

                do { indexPrice = getMktData(clientSocket, Allture, "SPX", "IND", "", "CBOE"); }
                while (indexPrice == lastPrice);
                lastPrice = indexPrice; // to ensure vix, vix1, vix2 not consecutive,  since their prices aren't chaning very much each time, can't be consecutive for lastPrice.

                do { vix1Price = getMktData(clientSocket, Allture, "VIX", "FUT", openBook.vx1Exp, "CFE"); }
                while (vix1Price == lastPrice);
                lastPrice = vix1Price; 

                do { indexPrice = getMktData(clientSocket, Allture, "SPX", "IND", "", "CBOE"); }
                while (indexPrice == lastPrice);
                lastPrice = indexPrice; // to ensure vix, vix1, vix2 not consecutive,  since their prices aren't chaning very much each time, can't be consecutive for lastPrice.

                do { vix2Price = getMktData(clientSocket, Allture, "VIX", "FUT", openBook.vx2Exp, "CFE"); }
                while (vix2Price == lastPrice);
                lastPrice = vix2Price;

                Console.WriteLine("\n UX2 price: " + vix2Price + "\n");

                if (!alert_vix && (vix1Price < vixPrice))
                {
                    SendMail("jzhang@alltuream.com", "Alert vix backwardation ! ", "VIX:" + vixPrice + " UX1:" + vix1Price);
                  //  SendMail("jzhang@leaderfunding.com", "Alert vix backwardation ! ", "VIX:" + vixPrice + " UX1:" + vix1Price);
                    alert_vix = true;
                }

                if (!alert_vix1 && (vix2Price < vix1Price))
                {
                    SendMail("jzhang@alltuream.com", "Alert UX1 backwardation ! ", "UX1:" + vix1Price + " UX2:" + vix2Price);
                  //  SendMail("jzhang@leaderfunding.com", "Alert UX1 backwardation ! ", "UX1:" + vix1Price + " UX2:" + vix2Price);
                    alert_vix1 = true;
                }

                for (int n = 0; n < openBook.positionCNT; n++)
                {
                    if (openBook.size[n] == 0) continue;

                    //**********************************************************
                    //*** Real time market price - Real Time Index bid/ask/mid**
                    //**********************************************************
                    if (openBook.symbol[n] == "SPX")
                    {
                        if (openBook.CallPut[n] == "Put")
                            if (openBook.legS[n] < putSPXlegS[0]) 
                                continue;

                        do { indexPrice = getMktData(clientSocket, Allture, "SPX", "IND", "", "CBOE"); }
                        while (indexPrice == lastPrice);
                        lastPrice = indexPrice;

                        Console.WriteLine("\n spx price: " + indexPrice + "\n");

                        exitPoint = SPXexitPoint;
                    }
                    else if (openBook.symbol[n] == "ES")
                    {
                        if (openBook.CallPut[n] == "Put")
                            if (openBook.legS[n] < putESlegS[0])
                                continue;

                        do { indexPrice = getMktData(clientSocket, Allture, "ES", "FUT", openBook.underlyingExp[n], "GLOBEX"); }
                        while (indexPrice == lastPrice);
                        lastPrice = indexPrice;

                        Console.WriteLine("\n es price: " + indexPrice + "\n");

                        exitPoint = ESexitPoint;
                    }
                    else if (openBook.symbol[n] == "CL")
                    {
                        if (openBook.CallPut[n] == "Put")
                            if (openBook.legS[n] < putCLlegS[0])
                                continue;

                        do { indexPrice = getMktData(clientSocket, Allture, "CL", "FUT", openBook.underlyingExp[n], "NYMEX"); }
                        while (indexPrice == lastPrice);
                        lastPrice = indexPrice;

                        Console.WriteLine("\n cl price: " + indexPrice + "\n");

                        exitPoint = CLexitPoint;
                    }
                    else if (openBook.symbol[n] == "GC")
                    {
                        if (openBook.CallPut[n] == "Put")
                            if (openBook.legS[n] < putGClegS[0])
                                continue;

                        do { indexPrice = getMktData(clientSocket, Allture, "GC", "FUT", openBook.underlyingExp[n], "NYMEX"); }
                        while (indexPrice == lastPrice);
                        lastPrice = indexPrice;

                        Console.WriteLine("\n GC price: " + indexPrice + "\n");

                        exitPoint = GCexitPoint;
                    }
                    else continue;

                    if (openBook.CallPut[n] == "Put")
                    {
                        if (!alert1[n] && (indexPrice <= (openBook.legS[n] + exitPoint + Alert1)))
                        {
                            // if alert1, sending email;
                            SendMail("jzhang@alltuream.com", "Alert1 !! " + openBook.model[n], openBook.symbol[n] + " " + indexPrice + ", " + openBook.legS[n] + ", " + openBook.expiration[n].Substring(4) + ", " + openBook.size[n] + ", " + openBook.accountID[n]);
                        //    SendMail("jzhang@leaderfunding.com", "Alert1 ! " + openBook.model[n], openBook.symbol[n] + " " + indexPrice + ", " + openBook.legS[n] + ", " + openBook.expiration[n].Substring(4) + ", " + openBook.size[n] + ", " + openBook.accountID[n]);
                            alert1[n] = true;
                        }

                     /*   if (!alert2[n] && (indexPrice <= (openBook.legS[n] + exitPoint + Alert2)))
                      //  {
                            // if alert2, sending email;
                      //      SendMail("jzhang@alltuream.com", "Alert2 !! " + openBook.model[n], openBook.symbol[n] + " " + indexPrice + ", " + openBook.legS[n] + ", " + openBook.expiration[n].Substring(4) + ", " + openBook.legS[n] + ", " + openBook.size[n] + ", " + openBook.accountID[n]);
                         //   SendMail("jzhang@leaderfunding.com", "Alert2 ! " + openBook.model[n], openBook.symbol[n] + " " + indexPrice + ", " + openBook.legS[n] + ", " + openBook.expiration[n].Substring(4) + ", " + openBook.size[n] + ", " + openBook.accountID[n]);
                      //      alert2[n] = true;
                     //   }
                        if (!alert3[n] && (indexPrice <= (openBook.legS[n] + exitPoint)))
                        {
                            // if alert3, sending email;
                            SendMail("jzhang@alltuream.com", "Alert3 !! " + openBook.model[n], openBook.symbol[n] + " " + indexPrice + ", " + openBook.legS[n] + ", " + openBook.expiration[n].Substring(4) + ", " + openBook.size[n] + ", " + openBook.accountID[n]);
                         //   SendMail("jzhang@leaderfunding.com", "Alert3 ! " + openBook.model[n], openBook.symbol[n] + " " + indexPrice + ", " + ", " + openBook.legS[n] + openBook.expiration[n].Substring(4) + ", " + openBook.size[n] + ", " + openBook.accountID[n]);
                            alert3[n] = true;
                        }
                        */
                        //  very complicated for auto exit and high risk, need more time to think through !!!
                        // seems if we don't directly get position info from IB, then should not do any auto exits
                        if (autoExit)
                            if (alert3[n])
                            {
                                //***  estimate the current spread's price and contract ID for combo orders                                    
                                if (openBook.symbol[n] == "SPX")
                                {
                                    spread_premium = 0;
                                    //  spread_premium = 0.05 + getSpreadPremium(clientSocket, Allture, openBook, n, openBook.legS[n], openBook.legB[n], "OPT", "SMART");
                                    legS_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legS[n], "OPT", "SMART");
                                    legB_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legB[n], "OPT", "SMART");

                                    //********* Place order ***************************
                                    clientSocket.reqIds(-1);
                                    Thread.Sleep(sleep1);
                                    clientSocket.placeOrder(Allture.NextOrderId, ContractSamples.Contract_Combo(openBook.symbol[n], "BAG", openBook.expiration[n], openBook.CallPut[n], legB_ID, legS_ID, "100", "SMART"), OrderSamples.ComboLimitOrder("SELL", openBook.size[n], -spread_premium, false));
                                    Thread.Sleep(sleep2);
                                    openBook.size[n] = 0;

                                    if (openBook.legS[n] == putSPXlegS[0]) putSPXlegS.RemoveAt(0);
                                 //   else if (openBook.legS[n] == putSPXlegS[1]) putSPXlegS.RemoveAt(1);
                                 //   else if (openBook.legS[n] == putSPXlegS[2]) putSPXlegS.RemoveAt(2);
                                }
                                else if (openBook.symbol[n] == "ES")
                                {
                                    spread_premium = 0;
                                    //  spread_premium = 0.05 + getSpreadPremium(clientSocket, Allture, openBook, n, openBook.legS[n], openBook.legB[n], "FOP", "GLOBEX");
                                    legS_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legS[n], "FOP", "GLOBEX");
                                    legB_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legB[n], "FOP", "GLOBEX");

                                    //********* Place order ************************
                                    clientSocket.reqIds(-1);
                                    Thread.Sleep(sleep1);
                                    clientSocket.placeOrder(Allture.NextOrderId, ContractSamples.Contract_Combo(openBook.symbol[n], "BAG", openBook.expiration[n], openBook.CallPut[n], legB_ID, legS_ID, "50", "GLOBEX"), OrderSamples.ComboLimitOrder("SELL", openBook.size[n], -spread_premium, false));
                                    Thread.Sleep(sleep2);
                                    openBook.size[n] = 0;

                                    if (openBook.legS[n] == putESlegS[0]) putESlegS.RemoveAt(0);
                                 //   else if (openBook.legS[n] == putESlegS[1]) putESlegS.RemoveAt(1);
                                  //  else if (openBook.legS[n] == putESlegS[2]) putESlegS.RemoveAt(2);
                                }
                                else if (openBook.symbol[n] == "CL")
                                {
                                    spread_premium = 0;
                                    //  spread_premium = getSpreadPremium(clientSocket, Allture, openBook, n, openBook.legS[n], openBook.legB[n], "FOP", "NYMEX");
                                    legS_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legS[n], "FOP", "NYMEX");
                                    legB_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legB[n], "FOP", "NYMEX");

                                    //********* Place order ************************
                                    clientSocket.reqIds(-1);
                                    Thread.Sleep(sleep1);
                                    clientSocket.placeOrder(Allture.NextOrderId, ContractSamples.Contract_Combo(openBook.symbol[n], "BAG", openBook.expiration[n], openBook.CallPut[n], legB_ID, legS_ID, "1000", "NYMEX"), OrderSamples.ComboLimitOrder("SELL", openBook.size[n], -spread_premium, false));
                                    Thread.Sleep(sleep2);
                                    openBook.size[n] = 0;

                                    if (openBook.legS[n] == putESlegS[0]) putCLlegS.RemoveAt(0);
                                    //   else if (openBook.legS[n] == putESlegS[1]) putESlegS.RemoveAt(1);
                                    //  else if (openBook.legS[n] == putESlegS[2]) putESlegS.RemoveAt(2);
                                }
                                else if (openBook.symbol[n] == "GC")
                                {
                                    spread_premium = 0;
                                    //  spread_premium = getSpreadPremium(clientSocket, Allture, openBook, n, openBook.legS[n], openBook.legB[n], "FOP", "NYMEX");
                                    legS_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legS[n], "FOP", "NYMEX");
                                    legB_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legB[n], "FOP", "NYMEX");

                                    //********* Place order ************************
                                    clientSocket.reqIds(-1);
                                    Thread.Sleep(sleep1);
                                    clientSocket.placeOrder(Allture.NextOrderId, ContractSamples.Contract_Combo(openBook.symbol[n], "BAG", openBook.expiration[n], openBook.CallPut[n], legB_ID, legS_ID, "100", "NYMEX"), OrderSamples.ComboLimitOrder("SELL", openBook.size[n], -spread_premium, false));
                                    Thread.Sleep(sleep2);
                                    openBook.size[n] = 0;

                                    if (openBook.legS[n] == putESlegS[0]) putGClegS.RemoveAt(0);
                                    //   else if (openBook.legS[n] == putESlegS[1]) putESlegS.RemoveAt(1);
                                    //  else if (openBook.legS[n] == putESlegS[2]) putESlegS.RemoveAt(2);
                                }
                                else continue;

                                //  very complicated, need more time to think through !!!  one thought - just simply place order, don't cancel it, until it filled
                                // seems if we don't directly get position info from IB, then should not do any auto exits                               
                            }
                    }

                    else if (openBook.CallPut[n] == "Call")
                    {
                        if (!alert3[n] && (indexPrice >= (openBook.legS[n] - exitPoint)))
                        {
                            // if alert3, sending email;
                            SendMail("jzhang@alltuream.com", "call Alert3 !! " + openBook.model[n], openBook.symbol[n] + " " + indexPrice + ", " + openBook.legS[n] + ", " + openBook.expiration[n].Substring(4) + ", " + openBook.size[n] + ", " + openBook.accountID[n]);
                            SendMail("jzhang@leaderfunding.com", "call Alert3 ! " + openBook.model[n], openBook.symbol[n] + " " + indexPrice + ", " + ", " + openBook.legS[n] + openBook.expiration[n].Substring(4) + ", " + openBook.size[n] + ", " + openBook.accountID[n]);
                            alert3[n] = true;
                        }

                        //  very complicated for auto exit and high risk, need more time to think through !!!
                        // seems if we don't directly get position info from IB, then should not do any auto exits
                        //  must have the auto exit.  otherwise, you can't handle that many position mannually.  and sychological/emotion is another issue.  
                        if (autoExit)
                            if (alert3[n])
                            {
                                //***  estimate the current spread's price and contract ID for combo orders                                    
                                if (openBook.symbol[n] == "SPX")
                                {
                                    // spread_premium = 0;
                                    spread_premium = 0.05 + getSpreadPremium(clientSocket, Allture, openBook, n, openBook.legS[n], openBook.legB[n], "OPT", "SMART");
                                    legS_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legS[n], "OPT", "SMART");
                                    legB_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legB[n], "OPT", "SMART");

                                    //********* Place order ***************************
                                    clientSocket.reqIds(-1);
                                    Thread.Sleep(sleep1);
                                    clientSocket.placeOrder(Allture.NextOrderId, ContractSamples.Contract_Combo(openBook.symbol[n], "BAG", openBook.expiration[n], openBook.CallPut[n], legB_ID, legS_ID, "100", "SMART"), OrderSamples.ComboLimitOrder("SELL", openBook.size[n], -spread_premium, false));
                                    Thread.Sleep(sleep2);
                                    openBook.size[n] = 0;

                                    if (openBook.legS[n] == callSPXlegS[0]) callSPXlegS.RemoveAt(0);
                                    //   else if (openBook.legS[n] == callSPXlegS[1]) callSPXlegS.RemoveAt(1);
                                    //   else if (openBook.legS[n] == callSPXlegS[2]) callSPXlegS.RemoveAt(2);
                                }
                                else if (openBook.symbol[n] == "ES")
                                {
                                    // spread_premium = 0;
                                    spread_premium = 0.05 + getSpreadPremium(clientSocket, Allture, openBook, n, openBook.legS[n], openBook.legB[n], "FOP", "GLOBEX");
                                    legS_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legS[n], "FOP", "GLOBEX");
                                    legB_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legB[n], "FOP", "GLOBEX");

                                    //********* Place order ************************
                                    clientSocket.reqIds(-1);
                                    Thread.Sleep(sleep1);
                                    clientSocket.placeOrder(Allture.NextOrderId, ContractSamples.Contract_Combo(openBook.symbol[n], "BAG", openBook.expiration[n], openBook.CallPut[n], legB_ID, legS_ID, "50", "GLOBEX"), OrderSamples.ComboLimitOrder("SELL", openBook.size[n], -spread_premium, false));
                                    Thread.Sleep(sleep2);
                                    openBook.size[n] = 0;

                                    if (openBook.legS[n] == callESlegS[0]) callESlegS.RemoveAt(0);
                                    //   else if (openBook.legS[n] == callESlegS[1]) callESlegS.RemoveAt(1);
                                    //  else if (openBook.legS[n] == callESlegS[2]) callESlegS.RemoveAt(2);
                                    else if (openBook.symbol[n] == "CL")
                                    {
                                        spread_premium = 0;
                                        //  spread_premium = getSpreadPremium(clientSocket, Allture, openBook, n, openBook.legS[n], openBook.legB[n], "FOP", "NYMEX");
                                        legS_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legS[n], "FOP", "NYMEX");
                                        legB_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legB[n], "FOP", "NYMEX");

                                        //********* Place order ************************
                                        clientSocket.reqIds(-1);
                                        Thread.Sleep(sleep1);
                                        clientSocket.placeOrder(Allture.NextOrderId, ContractSamples.Contract_Combo(openBook.symbol[n], "BAG", openBook.expiration[n], openBook.CallPut[n], legB_ID, legS_ID, "1000", "NYMEX"), OrderSamples.ComboLimitOrder("SELL", openBook.size[n], -spread_premium, false));
                                        Thread.Sleep(sleep2);
                                        openBook.size[n] = 0;

                                        if (openBook.legS[n] == callCLlegS[0]) callCLlegS.RemoveAt(0);
                                        //   else if (openBook.legS[n] == callCLlegS[1]) callCLlegS.RemoveAt(1);
                                        //  else if (openBook.legS[n] == callCLlegS[2]) callCLlegS.RemoveAt(2);
                                    }
                                    else if (openBook.symbol[n] == "GC")
                                    {
                                        spread_premium = 0;
                                        //  spread_premium = getSpreadPremium(clientSocket, Allture, openBook, n, openBook.legS[n], openBook.legB[n], "FOP", "NYMEX");
                                        legS_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legS[n], "FOP", "NYMEX");
                                        legB_ID = getContractID(clientSocket, Allture, openBook, n, openBook.legB[n], "FOP", "NYMEX");

                                        //********* Place order ************************
                                        clientSocket.reqIds(-1);
                                        Thread.Sleep(sleep1);
                                        clientSocket.placeOrder(Allture.NextOrderId, ContractSamples.Contract_Combo(openBook.symbol[n], "BAG", openBook.expiration[n], openBook.CallPut[n], legB_ID, legS_ID, "100", "NYMEX"), OrderSamples.ComboLimitOrder("SELL", openBook.size[n], -spread_premium, false));
                                        Thread.Sleep(sleep2);
                                        openBook.size[n] = 0;

                                        if (openBook.legS[n] == callCLlegS[0]) callGClegS.RemoveAt(0);
                                        //   else if (openBook.legS[n] == callCLlegS[1]) callCLlegS.RemoveAt(1);
                                        //  else if (openBook.legS[n] == callCLlegS[2]) callCLlegS.RemoveAt(2);
                                    }
                                    //  very complicated, need more time to think through !!!  one thought - just simply place order, don't cancel it, until it filled
                                    // seems if we don't directly get position info from IB, then should not do any auto exits                               
                                }
                            }

                    }


                }

            }

            Console.WriteLine("\n start finish, time: " + DateTime.Now + " \n");
            
            return 0;
            }




        private static void SendMail(string to, string body, string subject)
        {  
            MailMessage message = new MailMessage("myself@gmail.com", to, subject, body);
            SmtpClient mailClient = new SmtpClient("smtp.gmail.com", 587);
           
            mailClient.EnableSsl = true;
            mailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            mailClient.UseDefaultCredentials = false;
            mailClient.Credentials = new System.Net.NetworkCredential("jzhang18", "Frankzj!8");
            mailClient.Send(message);
            message.Dispose();
        }

        private static void position_xls_readin(_Excel.Worksheet wsheet, TradeBooks book)
        {
            int i=0;
            _Excel.Range myRange = wsheet.UsedRange;

            //These two lines clear up the blank cells.
            wsheet.Columns.ClearFormats();
            wsheet.Rows.ClearFormats();

            book.positionCNT = myRange.Rows.Count-2;
            
            for (i = 1; i < myRange.Rows.Count - 1; i++)
             {
                book.symbol[i-1] = wsheet.Cells[i+2, 6].value2;
                book.expiration[i-1] = wsheet.Cells[i + 2, 4].value2;
                book.CallPut[i-1] = wsheet.Cells[i + 2, 7].value2;
                book.legS[i-1] = Convert.ToDouble(wsheet.Cells[2 + i, 8].value2); 
                book.legB[i-1] = Convert.ToDouble(wsheet.Cells[2 + i, 9].value2);
                book.size[i-1] = (int)wsheet.Cells[i + 2, 10].value2;
                book.premium[i-1] = wsheet.Cells[i + 2, 12].value2;
                book.day_cnt[i-1] = (int)wsheet.Cells[i + 2, 15].value2;
                book.accountID[i-1] = wsheet.Cells[i + 2, 14].value2;
                book.model[i-1] = wsheet.Cells[i + 2, 24].value2;
                book.multipler[i-1] = (int)wsheet.Cells[i + 2, 30].value2;
                book.underlyingExp[i - 1] = wsheet.Cells[i + 2, 31].Text;
            }
        }

        private static double getMktData(EClientSocket client, EWrapperImpl Allture, string symbol, string sectype, string expiration, string exchange)
        {

            if (string.IsNullOrEmpty(expiration))
                client.reqMktData(1005, ContractSamples.Contract_RT(symbol, sectype, exchange), string.Empty, false, false, null);
            else
                client.reqMktData(1005, ContractSamples.Contract_FUT(symbol, sectype, expiration, exchange), string.Empty, false, false, null);

            Thread.Sleep(sleep2);
            client.cancelMktData(1005);
           // Thread.Sleep(sleep0);
            Console.WriteLine("\n price: " + Allture.mkt_price + "\n");
            return (Allture.mkt_price);
        }

        private static int getContractID(EClientSocket client, EWrapperImpl Allture, TradeBooks Book, int m, double leg, string sectype, string exchange)
        {
            Allture.checkContractEnd = false;
            client.reqContractDetails(20050, ContractSamples.Contract_Options(Book.symbol[m], sectype, Book.expiration[m], Book.CallPut[m], leg, Book.multipler[m].ToString(), exchange));
            //Thread.Sleep(sleep1);
            while (!Allture.checkContractEnd)
                Thread.Sleep(200);
            Console.WriteLine("contract_ID: " + Allture.contract_ID + "\n");
            return (Allture.contract_ID);
        }

        private static double getSpreadPremium(EClientSocket client, EWrapperImpl Allture, TradeBooks Book, int m, double LegS, double LegB, string sectype, string exchange)
        {
            double LegB_price, LegS_price;
            client.reqMktData(3011, ContractSamples.Contract_Options(Book.symbol[m], sectype, Book.expiration[m], Book.CallPut[m], LegB, Book.multipler[m].ToString(), exchange), string.Empty, false, false, null);
            Thread.Sleep(1000);
            client.cancelMktData(3011);
            LegB_price = Allture.ask_price;
            Console.WriteLine("legB ask: " + LegB_price + "\n");

            client.reqMktData(3012, ContractSamples.Contract_Options(Book.symbol[m], sectype, Book.expiration[m], Book.CallPut[m], LegS, Book.multipler[m].ToString(), exchange), string.Empty, false, false, null);
            Thread.Sleep(1000);
            client.cancelMktData(3012);
            LegS_price = Allture.ask_price;
            Console.WriteLine("legS ask: " + LegS_price + "\n");

            return (LegS_price - LegB_price);
        }
        

//! ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        private static void historicalTicks(EClientSocket client)
        {
            //! [reqhistoricalticks]
            client.reqHistoricalTicks(18001, ContractSamples.USStockAtSmart(), "20170712 21:39:33", null, 10, "TRADES", 1, true, null);
            client.reqHistoricalTicks(18002, ContractSamples.USStockAtSmart(), "20170712 21:39:33", null, 10, "BID_ASK", 1, true, null);
            client.reqHistoricalTicks(18003, ContractSamples.USStockAtSmart(), "20170712 21:39:33", null, 10, "MIDPOINT", 1, true, null);
            //! [reqhistoricalticks]
        }

        private static void pnl(EClientSocket client)
        {
            //! [reqpnl]
            client.reqPnL(17001, "DUD00029", "");
            //! [reqpnl]
            Thread.Sleep(1000);
            //! [cancelpnl]
            client.cancelPnL(17001);
            //! [cancelpnl]
        }

        private static void pnlSingle(EClientSocket client)
        {
            //! [reqpnlsingle]
            client.reqPnLSingle(17001, "DUD00029", "", 268084);
            //! [reqpnlsingle]
            Thread.Sleep(1000);
            //! [cancelpnlsingle]
            client.cancelPnLSingle(17001);
            //! [cancelpnlsingle]

        }

        private static void histogramData(EClientSocket client)
        {
            //! [reqHistogramData]
            client.reqHistogramData(15001, ContractSamples.USStockWithPrimaryExch(), false, "1 week");
            //! [reqHistogramData]
            Thread.Sleep(2000);
            //! [cancelHistogramData]
            client.cancelHistogramData(15001);
            //! [cancelHistogramData]
        }

        private static void headTimestamp(EClientSocket client)
        {
            //! [reqHeadTimeStamp]
            client.reqHeadTimestamp(14001, ContractSamples.USStock(), "TRADES", 1, 1);
            //! [reqHeadTimeStamp]
            Thread.Sleep(1000);
            //! [cancelHeadTimestamp]
            client.cancelHeadTimestamp(14001);
            //! [cancelHeadTimestamp]

        }

        private static void realTimeBars(EClientSocket client)
        {

            client.reqRealTimeBars(3004, ContractSamples.FuturesOnOptionsES(), 5, "TRADES", true, null);
            //! [reqrealtimebars]
            Thread.Sleep(1000);
            /*** Canceling real time bars ***/
            //! [cancelrealtimebars]
            client.cancelRealTimeBars(3004);
            //! [cancelrealtimebars]
        }

        private static void marketDataType(EClientSocket client)
        {
            //! [reqmarketdatatype]
            /*** Switch to live (1) frozen (2) delayed (3) or delayed frozen (4)***/
            client.reqMarketDataType(1);
            //! [reqmarketdatatype]
        }

        private static void historicalDataRequests(EClientSocket client)
        {
            /*** Requesting historical data ***/
            //! [reqhistoricaldata]
            String queryTime = DateTime.Now.AddMonths(-6).ToString("yyyyMMdd HH:mm:ss");
            client.reqHistoricalData(4001, ContractSamples.EurGbpFx(), queryTime, "1 M", "1 day", "MIDPOINT", 1, 1, false, null);
            client.reqHistoricalData(4002, ContractSamples.EuropeanStock(), queryTime, "10 D", "1 min", "TRADES", 1, 1, false, null);
            //! [reqhistoricaldata]
            Thread.Sleep(2000);
            /*** Canceling historical data requests ***/
            client.cancelHistoricalData(4001);
            client.cancelHistoricalData(4002);
        }

        private static void marketScanners(EClientSocket client)
        {
            /*** Requesting all available parameters which can be used to build a scanner request ***/
            //! [reqscannerparameters]
            client.reqScannerParameters();
            //! [reqscannerparameters]
            Thread.Sleep(2000);

            /*** Triggering a scanner subscription ***/
            //! [reqscannersubscription]
            client.reqScannerSubscription(7001, ScannerSubscriptionSamples.HighOptVolumePCRatioUSIndexes(), null);
            //! [reqscannersubscription]

            Thread.Sleep(2000);
            /*** Canceling the scanner subscription ***/
            //! [cancelscannersubscription]
            client.cancelScannerSubscription(7001);
            //! [cancelscannersubscription]
        }


        private static void ConditionSamples(EClientSocket client, int nextOrderId)
        {
            //! [order_conditioning_activate]
            Order mkt = OrderSamples.MarketOrder("BUY", 100);
            //Order will become active if conditioning criteria is met
            mkt.ConditionsCancelOrder = true;
            mkt.Conditions.Add(OrderSamples.PriceCondition(208813720, "SMART", 600, false, false));
            mkt.Conditions.Add(OrderSamples.ExecutionCondition("EUR.USD", "CASH", "IDEALPRO", true));
            mkt.Conditions.Add(OrderSamples.MarginCondition(30, true, false));
            mkt.Conditions.Add(OrderSamples.PercentageChangeCondition(15.0, 208813720, "SMART", true, true));
            mkt.Conditions.Add(OrderSamples.TimeCondition("20160118 23:59:59", true, false));
            mkt.Conditions.Add(OrderSamples.VolumeCondition(208813720, "SMART", false, 100, true));
            client.placeOrder(nextOrderId++, ContractSamples.EuropeanStock(), mkt);
            //! [order_conditioning_activate]

            //Conditions can make the order active or cancel it. Only LMT orders can be conditionally canceled.
            //! [order_conditioning_cancel]
            Order lmt = OrderSamples.LimitOrder("BUY", 100, 20);
            //The active order will be cancelled if conditioning criteria is met
            lmt.ConditionsCancelOrder = true;
            lmt.Conditions.Add(OrderSamples.PriceCondition(208813720, "SMART", 600, false, false));
            client.placeOrder(nextOrderId++, ContractSamples.EuropeanStock(), lmt);
            //! [order_conditioning_cancel]
        }
    }

}
