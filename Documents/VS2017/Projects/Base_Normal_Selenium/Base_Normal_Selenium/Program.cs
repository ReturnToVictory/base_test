
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Support.UI;
using System.Text.RegularExpressions;

namespace Base_Normal_Selenium
{
    class Program
    {
        public static string reg_exp(string buf, string ful_adress)
        {
            string pattern = buf;
            string answer = "";
            Regex regex1 = new Regex(pattern, RegexOptions.IgnoreCase);
            Match match1 = regex1.Match(ful_adress);
            answer = match1.Groups[1].Value;
            return answer;
        }

        static string find(string name, List<string> list_test)
        {
            int index = list_test.LastIndexOf(name);
            if (index == -1 || index == 0)
                return "";

            name = list_test[index + 1];
            return name;
        }
        public static void read(string text, string cell, int index, Excel.Workbook excelappworkbook)
        {
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            Excel.Range excelcells;

            cell = cell + index;
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelcells = excelworksheet.get_Range(cell, cell);
            excelcells.Value2 = text;
        }
        static void Main(string[] args)
        {
            Excel.Application excelapp;
            Excel.Range excelcells;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            Excel.Workbook excelappworkbook;

            Regex VAT = new Regex(@"[A-Z]{2}([0-9]{2-12}|\\w{9}|[U]\\w{8}|\\w[0-9]{7}\\w|\\w[0-9]{8}|GD[0-9]{3}|HA[0-9]{3}|[0-9]{7}\\wW|[0-9]\\w[0-9]{5}\\w|[0-9]{9}B[0-9]{2})");

            string file_name = "D:\\Users\\ysibirkin\\export.xlsx";

            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelappworkbook = excelapp.Workbooks.Open(@file_name);
            excelsheets = excelappworkbook.Worksheets;


            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://ec.europa.eu/taxation_customs/vies/vatResponse.html");
            driver.Manage();

            string cell = "P";
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            long count = 0;
            for (long i = 2; i < 139200; i++)
            {
                excelcells = excelworksheet.get_Range((cell + i), (cell + i));
                string Nalog_number = excelcells.Value2;
                if (excelcells.Value2 == null)
                {
                    Console.WriteLine("Empty");//delete 
                    continue;
                }
                else
                {

                   // Match match1 = VAT.Match(Nalog_number);
                   // Console.WriteLine("REG_EXP===" + match1.Groups[1].Value);
                   // Console.WriteLine("Nalog_Number" + Nalog_number);
                    //Console.ReadKey();
                    
                    if ( Nalog_number[1]> 'A' && Nalog_number[1] < 'Z' )
                    {
                        Console.WriteLine("BEZ STRANI" + Nalog_number);
                        count++;
                    }

                    /*
                    string country = "EE";

                    IWebElement searchInput;
                    searchInput = driver.FindElement(By.Id("number"));

                    searchInput.SendKeys("100366327");

                    //ввод страны , выбор из выпадающего окна
                    // select the drop down list
                    var memberStateCode = driver.FindElement(By.Name("memberStateCode"));
                    //create select element object 
                    var selectElement = new SelectElement(memberStateCode);
                    //select by value
                    selectElement.SelectByValue(country);

                    // select by text
                    // selectElement.SelectByText("HighSchool");

                    searchInput = driver.FindElement(By.Id("submit"));
                    //searchInput.
                    searchInput.SendKeys(Keys.Enter);

                    IList<IWebElement> all = driver.FindElements(By.TagName("td"));
                    List<string> list_test = new List<string>();

                    foreach (IWebElement element in all)
                    {
                        list_test.Add(element.Text);
                        //  Console.WriteLine("test==" + element.Text);
                    }
                    //Member State
                    string state = "Member State";
                    state = find(state, list_test);
                    Console.WriteLine("Member State===" + state);

                    // Регистрационный номер "S"
                    string registre = "VAT Number";
                    registre = find(registre, list_test);
                    Console.WriteLine("INN===" + registre);

                    //Date when request received
                    string date = "Date when request received";
                    date = find(date, list_test);
                    Console.WriteLine("Date===" + date);

                    //Date when request received
                    string name = "Name";
                    name = find(name, list_test);
                    Console.WriteLine("Name===" + name);

                    //Address
                    string adress = "Address";
                    adress = find(adress, list_test);
                    Console.WriteLine("Address===" + adress);

                    //Consultation Number
                    string number = "Consultation Number";
                    number = find(number, list_test);
                    Console.WriteLine("Consultation Number===" + number);
                    */
                }
            }
            Console.WriteLine("COUNT====" + count);
        }
    }
}
