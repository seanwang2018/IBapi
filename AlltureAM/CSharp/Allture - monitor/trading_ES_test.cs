/* Copyright (C) 2018 Interactive Brokers LLC. All rights reserved. This code is subject to the terms
 * and conditions of the IB API Non-Commercial License or the IB API Commercial License, as applicable. */
using System;
using IBApi;
using System.Threading;
using Samples;
using _Excel = Microsoft.Office.Interop.Excel; 


namespace Allture
{
    public class Trading_FUT
    {
        /* IMPORTANT: always use your paper trading account. The code below will submit orders as part of the demonstration. */
        /* IB will not be responsible for accidental executions on your live account. */
        /* Any stock or option symbols displayed are for illustrative purposes only and are not intended to portray a recommendation. */
        /* Before contacting our API support team please refer to the available documentation. */

        static EWrapperImpl Allture = new EWrapperImpl();
        static TradeBooks openBook = new TradeBooks();

        static ExecutionFilter executionFilter = new ExecutionFilter();

        public static int sleep0 = 200;
        public static int sleep1 = 1000;
        public static int sleep2 = 2000;

        public static int Main(string[] args)
        {
            // specify account and strategy           
            string IBaccount = TradeBooks.accountInit();


            string tradingModel = "es_growth";
            string tradingInstrument = "ES";
            int multipler = 50;


            int tryCNT = 1;
            int tryGap = (int)openBook.gap * 20;
            int bestCNT = 0;

            double[] Returns = new double[2];
            double[] Premium = new double[tryCNT];
            int[] LegB_ID = new int[tryCNT];
            int[] LegB = new int[tryCNT];

            int legS;
            int legS_ID;

            openBook.positionCNT = 1;

            for (int n = 0; n < openBook.positionCNT; n++)
            {
                openBook.symbol[n] = tradingInstrument;
                openBook.multipler[n] = multipler;
            }


            //!  start the connection
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


            //****  current time *****
            string date_time, last_time;
            date_time = DateTime.Now.ToString("hh:mm:ss tt");
            //**************************
            //*** Account Info       ***
            //**************************

            clientSocket.reqAccountSummary(9001, "All", AccountSummaryTags.GetAllTags());
            Thread.Sleep(1000);

            Console.WriteLine("account ID: " + Allture.account_ID);
            Console.WriteLine("account Value: " + Allture.account_value);
            Console.WriteLine("account_BuyingPower: " + Allture.account_BuyingPower);
            Console.WriteLine("account_InitMarginReq: " + Allture.account_InitMarginReq);
            Console.WriteLine("account_MaintMarginReq: " + Allture.account_MaintMarginReq);
            Console.WriteLine("account_ExcessLiquidity: " + Allture.account_ExcessLiquidity);
            Console.WriteLine("account_AvailableFunds: " + Allture.account_AvailableFunds);
            Console.WriteLine("\n");



            Allture.remainingOrderSize = 100000; // set to maxium order size

            clientSocket.reqMarketDataType(1);
            // clientSocket.reqGlobalCancel();
            Thread.Sleep(sleep1);

            while ((DateTime.Now > Convert.ToDateTime("09:30:00 AM")) && (DateTime.Now < Convert.ToDateTime("16:00:00 PM")))
            {
                last_time = date_time;

                for (int n = 0; n < openBook.positionCNT; n++)
                {
                    if (openBook.status[n] == "complete") continue;

                    if (openBook.status[n] == "submit")
                    {
                        Allture.currOrderId = openBook.currOrderID[n];
                        Allture.remainingOrderSize = 0;
                        Allture.checkOrderEnd = false;
                        clientSocket.reqOpenOrders();
                        while (!Allture.checkOrderEnd)
                            Thread.Sleep(200);

                        if (Allture.remainingOrderSize == 0)
                        {
                            openBook.status[n] = "complete";
                            openBook.capital[n] = 0;
                            openBook.size[n] = 0;
                            continue;  // if remaining size = 0, means this order has completely filled. then go to the next position
                        }

                    }


                    legS = 2700;

                    legS_ID = getContractID(clientSocket, Allture, openBook, n, legS, "FOP", "GLOBEX");

                    LegB[0] = legS - 100;

                    LegB_ID[0] = getContractID(clientSocket, Allture, openBook, n, LegB[0], "FOP", "GLOBEX");


                    //******* check the order status again.  if still exiting, cancel the order order, update the order size and capital
                    if (openBook.status[n] == "submit")
                    {
                        Allture.currOrderId = openBook.currOrderID[n];
                        Allture.remainingOrderSize = 0;
                        Allture.checkOrderEnd = false;
                        clientSocket.reqOpenOrders();
                        while (!Allture.checkOrderEnd)
                            Thread.Sleep(200);

                        if (Allture.remainingOrderSize == 0)
                        {
                            openBook.status[n] = "complete";
                            openBook.capital[n] = 0;
                            openBook.size[n] = 0;
                            continue;
                        }
                        else
                        {
                            clientSocket.cancelOrder(openBook.currOrderID[n]); //cancel existing order
                            Thread.Sleep(sleep1);
                            openBook.status[n] = "cancel";

                            if (Allture.remainingOrderSize > openBook.size[n]) Allture.remainingOrderSize = openBook.size[n];

                            openBook.capital[n] = openBook.capital[n] - (openBook.size[n] - Allture.remainingOrderSize) * openBook.margin[n];
                            openBook.size[n] = (int)(openBook.usingCapital * openBook.capital[n] / 10000);
                        }
                    }
                    else openBook.size[n] = (int)(100);



                    //********* Place order ************************
                    clientSocket.reqIds(-1);
                    Thread.Sleep(sleep1);
                    clientSocket.placeOrder(Allture.NextOrderId, ContractSamples.Contract_Combo(openBook.symbol[n], "BAG", openBook.expiration[n], openBook.CallPut[n], LegB_ID[0], legS_ID, openBook.multipler[n].ToString(), "GLOBEX"), OrderSamples.ComboLimitOrder("BUY", openBook.size[n], -1.0, false));
                    Thread.Sleep(sleep1);

                    //********* update remaining of the trading book **************
                    openBook.status[n] = "submit";
                    openBook.currOrderID[n] = Allture.NextOrderId;
                    openBook.legS[n] = legS;
                    openBook.legB[n] = LegB[0];
                    openBook.premium[n] = 1.0;

                }
            }

            clientSocket.reqGlobalCancel();
            Thread.Sleep(sleep1);

            Console.WriteLine("today's trading is done...time: " + DateTime.Now);
            clientSocket.eDisconnect();

            return 0;
        }


        private static int getContractID(EClientSocket client, EWrapperImpl Allture, TradeBooks Book, int m, int leg, string sectype, string exchange)
        {
            Allture.checkContractEnd = false;
            client.reqContractDetails(20010, ContractSamples.Contract_Options(Book.symbol[m], sectype, Book.expiration[m], Book.CallPut[m], leg, Book.multipler[m].ToString(), exchange));
            //Thread.Sleep(sleep1);
            while (!Allture.checkContractEnd)
                Thread.Sleep(200);
            Console.WriteLine("contract_ID: " + Allture.contract_ID + "\n");
            return (Allture.contract_ID);
        }
    }
}
