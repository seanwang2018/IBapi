using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IBApi
{ 
    public class TradeBooks
    {
        /* order info */
        public string[] symbol;
        public string[] CallPut;
        public string[] expiration;
        public string[] underlyingExp;
        public double[] legB;
        public double[] legS;
        public double[] margin;
        public double[] premium;
        public double[] currStd;
        public int[] currOrderID;
        public int[] size;
        public int[] day_cnt;
        public string[] accountID;
        public int positionCNT { get; set; }


        public string[] status;

        /* capital info */
        public double[] capital;

        /**** contract info ******/
        public int[] multipler;
        public double gap;
      
        //! risk control parameters
        public double exitBuffer { get; set; }
        public double probProfitTarget { get; set; }
        public double hisProfitTarget { get; set; }
        public double lossAdjustor { get; set; }
        public double lossDayAdjustor { get; set; }
        public double marginAdjustor { get; set; }
        public double adjStd { get; set; }
        public double usingCapital { get; set; }
        public int usingDays { get; set; }
        public string mktTrend { get; set; }

        //! expiration date
        public string vx1Exp { get; set; }
        public string vx2Exp { get; set; }
        public string vx3Exp { get; set; }

        /***** strategy info ******/
        public string[] model;
        public TradeBooks()
        {
            symbol = new string[Constants.max_position];
            expiration = new string[Constants.max_position];
            underlyingExp = new string[Constants.max_position];
            CallPut = new string[Constants.max_position];
            legB = new double[Constants.max_position];
            legS = new double[Constants.max_position];
            margin = new double[Constants.max_position];
            premium = new double[Constants.max_position];
            currStd = new double[Constants.max_position];
            currOrderID = new int[Constants.max_position];
            size = new int[Constants.max_position];
            day_cnt = new int[Constants.max_position];
            multipler = new int[Constants.max_position];
            capital = new double[Constants.max_position];
            status = new string[Constants.max_position];
            accountID = new string[Constants.max_position];
            positionCNT = 0;         
            model = new string[Constants.max_position];

            exitBuffer = 50;
            probProfitTarget = 0.5;
            hisProfitTarget = 0.3;
            lossAdjustor = 1.0;
            lossDayAdjustor = 0.6;
            marginAdjustor = 1.1;
            adjStd = 2.0;
            usingCapital = 1.0;
            usingDays = 10;
            gap = 5;

            mktTrend = "neutral";

            vx1Exp = "20190821";
            vx2Exp = "20190918";
            vx3Exp = "20191016";
            //8/21  9/18 10/16 11/20, 12/18
    }
 
        //! set up IB API channel 
        public static int[] channelSetup(string account, string model)
        {
            int[] channel = new int[2];

            if (account == "zhang882") channel[0] = 7496;

            if (model == "spx_safe") channel[1] = 1;
            if (model == "spx_growth") channel[1] = 2;
            if (model == "es_call") channel[1] = 3;
            if (model == "es_growth") channel[1] = 4;
            if (model == "spx_HR") channel[1] = 5;
            if (model == "spx_call") channel[1] = 6;
            if (model == "CL_put") channel[1] = 11;
            if (model == "CL_call") channel[1] = 12;
            if (model == "GC_put") channel[1] = 15;
            if (model == "GC_call") channel[1] = 16;

            if (model == "monitor") channel[1] = 21;
            if (model == "report") channel[1] = 99;

            return (channel);
        }


        public static List<string> account = new List<string> { "zhang882", "marong882", "zhangi882" };

        public static string xlsPathSetup(string IBaccount)
        {
            string path;
            int simulateflag = account.IndexOf(IBaccount);
            if (simulateflag == -1) path = "C:\\Users\\Jack\\Documents\\Companies\\allture\\Trading\\option_spread\\Accounts\\_simulate\\";
            else path = "C:\\Users\\Jack\\Documents\\Companies\\allture\\Trading\\option_spread\\Accounts\\";

            return (path);
        }

        public static string accountInit()
        {
            string IBaccount;
            string userInput;
            int selectID = 1;
            Console.WriteLine("Select an account: \n 1. zhang882 (default);  \n");
            userInput = Console.ReadLine();
            if (userInput != "")
                selectID = Convert.ToInt32(userInput);
            else
                selectID = 1;

            switch (selectID)
            {
                case 1:
                    IBaccount = "zhang882";
                    break;
                default:
                    IBaccount = "zhang882";
                    break;
            }

            return (IBaccount);
        }

        public static string trendInit()
        {
            string trend;
            string userInput;
            int selectID = 1;
            Console.WriteLine("Select mkt trend: \n 1. neutral (default); 2. bull; 3. bear; \n ");
            userInput = Console.ReadLine();
            if(userInput != "")
                selectID = Convert.ToInt32(userInput);
            else
                selectID = 1;

            switch (selectID)
            {
                case 1:
                    trend = "neutral";
                    break;
                case 2:
                    trend = "bull";
                    break;
                case 3:
                    trend = "bear";
                    break;
                default:
                    trend = "neutral";
                    break;
            }

            return (trend);
        }
    }
}