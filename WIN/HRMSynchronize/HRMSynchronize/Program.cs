using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HRMSynchronize
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
                    case "DepartmentSynchronize":
                        Console.WriteLine("HRM部門資料同步中...");
                        DepartmentSynchronize departmentSynchronize = new DepartmentSynchronize(company, secretKey);
                        departmentSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    case "UserSynchronize":
                        Console.WriteLine("HRM使用者資料同步中...");
                        UserSynchronize userSynchronize = new UserSynchronize(company, secretKey);
                        userSynchronize.Init();
                        Console.WriteLine("執行完成。");
                        break;
                    default:
                        Console.WriteLine("未執行任何程式。");
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
