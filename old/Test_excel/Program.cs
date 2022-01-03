using EasyInteropExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_excel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Person> people = new List<Person>();

            for (int i = 10; i < 100; i++)
            {
                people.Add(new Person() { Idade = i, Nome = "teste" + i });
            }


            //OExcel.ToExcel(people, Environment.CurrentDirectory, "teste.xlsx", OExcel.XlFileFormat.xlWorkbookDefault);
            OExcel.ExcelToWriteTxt(
                $"{Environment.CurrentDirectory}\\VIVO_MEC_OBM_03092021_2110.xlsx",
                Environment.CurrentDirectory, 
                OExcel.TextFormat.txt,
                "BillingOffer",
                new string[] { "B", "C", "D", "E", "H", "BH" },
                5,
                "B",
                ";"
                );
        }

        public class Person
        {
            public int Idade { get; set; }
            public string Nome { get; set; }
        }
    }
}
