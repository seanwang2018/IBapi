using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IBApi
{
    public class TradeStrategy
    {
        public static void setParameters(TradeBooks openBook, double Vix, double Ux1, double Ux2, double extraStd)
        {
            double vixUx1 = Ux1 / Vix;
            double Ux1Ux2 = Ux2 / Ux1;
            int usingDay = 10;
            double usingCapital = 1.0;

            if ((vixUx1 > 1.07) && (Ux1Ux2 > 1.07))
            {
                openBook.exitBuffer = 50;
                openBook.probProfitTarget = 0.15;
                openBook.hisProfitTarget = 0.1;
                openBook.lossAdjustor = 1.6;
                openBook.adjStd = 2.06 + extraStd;  // 98% probability
                openBook.usingCapital = usingCapital;
                openBook.usingDays = usingDay;
            }
            else if ((vixUx1 >= 1.02) && (Ux1Ux2 >= 1.02))
            {
                openBook.exitBuffer = 50;
                openBook.probProfitTarget = 0.20;
                openBook.hisProfitTarget = 0.15;
                openBook.lossAdjustor = 1.4;
                openBook.adjStd = 1.96 + extraStd;  //97.5% probability
                openBook.usingCapital = usingCapital;
                openBook.usingDays = usingDay;
            }
            else if ((vixUx1 >= 1.02) || (Ux1Ux2 >= 1.02))
            {
                openBook.exitBuffer = 50;
                openBook.probProfitTarget = 0.30;
                openBook.hisProfitTarget = 0.25;
                openBook.lossAdjustor = 1.3;
                openBook.adjStd = 1.76 + extraStd;  //96% prob
                openBook.usingCapital = usingCapital;
                openBook.usingDays = usingDay;
            }
            else
            {
                openBook.exitBuffer = 50;
                openBook.probProfitTarget = 0.4;
                openBook.hisProfitTarget = 0.3;
                openBook.lossAdjustor = 1.3;
                openBook.adjStd = 1.66 + extraStd; // 95% prob
                openBook.usingCapital = usingCapital;
                openBook.usingDays = usingDay;
            }
        }
    }
}