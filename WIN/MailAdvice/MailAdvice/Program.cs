using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailAdvice
{
    class Program
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            try
            {
                string transferType = "";
                string company = "";
                string secretKey = "";

                if (args.Length > 0)
                {
                    transferType = args[0].ToString();
                    company = args[1].ToString();
                    secretKey = args[2].ToString();
                }

                switch (transferType)
                {
                    case "DailyDeliveryMailAdvice":
                        Console.WriteLine("每日出貨明細...");
                        DailyDeliveryMailAdvice dailyDeliveryMailAdvice = new DailyDeliveryMailAdvice(company, secretKey);
                        dailyDeliveryMailAdvice.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "DailyDeliveryWipUnLinkDetailMailAdvice":
                        Console.WriteLine("每日出貨未綁定製令明細...");
                        DailyDeliveryWipUnLinkDetailMailAdvice dailyDeliveryWipUnLinkDetailMailAdvice = new DailyDeliveryWipUnLinkDetailMailAdvice(company, secretKey);
                        dailyDeliveryWipUnLinkDetailMailAdvice.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "SuggestionsForPurchaseRecord":
                        Console.WriteLine("粗胚建議請購清單(無條件)...");
                        SuggestionsForPurchaseRecord suggestionsForPurchaseRecord = new SuggestionsForPurchaseRecord(company, secretKey);
                        suggestionsForPurchaseRecord.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "SuggestionsForPurchaseConditionRecord":
                        Console.WriteLine("粗胚建議請購清單(有條件)...");
                        SuggestionsForPurchaseConditionRecord suggestionsForPurchaseConditionRecord = new SuggestionsForPurchaseConditionRecord(company, secretKey);
                        suggestionsForPurchaseConditionRecord.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "PrConfirmedNotProcuredRecord":
                        Console.WriteLine("請購單已確認但未採購的清單...");
                        PrConfirmedNotProcuredRecord prConfirmedNotProcuredRecord = new PrConfirmedNotProcuredRecord(company, secretKey);
                        prConfirmedNotProcuredRecord.Init();
                        Console.WriteLine("執行完成。");
                        break;
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                Console.WriteLine(e.Message);
            }
            finally
            {
                ReadLine(1000);
            }
        }

        static void ReadLine(int millisecond)
        {
            ReadLineDelegate readLineDelegate = Console.ReadLine;
            IAsyncResult asyncResult = readLineDelegate.BeginInvoke(null, null);
            asyncResult.AsyncWaitHandle.WaitOne(millisecond);
            if (asyncResult.IsCompleted)
            {
                string result = readLineDelegate.EndInvoke(asyncResult);
                Console.WriteLine("Read: " + result);
            }
            else
            {
                Console.WriteLine("程式自動關閉");
            }
        }

        delegate string ReadLineDelegate();
    }
}
