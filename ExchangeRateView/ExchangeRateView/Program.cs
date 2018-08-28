using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExchangeRateView
{
    public class Program
    {
        public static void Main(string[] args)
        {
            String res = "";
            var ie = new SHDocVw.InternetExplorer();
            //ネット接続
            try
            {
                ie.Navigate("https://moneykit.net/visitor/sb_rate/");
                Wait(ie);
                var doc = ie.Document as mshtml.IHTMLDocument3;
                var tmpCol = doc.getElementsByTagName("tbody") as mshtml.IHTMLElementCollection;
                tmpCol = tmpCol.item(0).getElementsByTagName("td") as mshtml.IHTMLElementCollection;
                res = tmpCol.item(2).innerText;
            }
            catch (Exception)
            {
                res = "接続エラー";
            }
            finally
            {
                ie.Quit();
                FileOutput(res);
            }
        }
        public static void Wait(SHDocVw.InternetExplorer ie)
        {
            while (ie.Busy || ie.ReadyState != SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE)
            {
                System.Threading.Thread.Sleep(100);
            }
        }
        public static void FileOutput(String str)
        {
            string fileName = Directory.GetCurrentDirectory() + @"\kida\result.txt";
            if (!Directory.Exists(Directory.GetCurrentDirectory() + @"\kida"))
            {
                Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\kida");
            }
            File.WriteAllText(fileName, str);
        }
    }
}
