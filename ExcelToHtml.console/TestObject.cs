using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtml.console
{
    public class Person
    {
        public string Name { get; set; }
        public string Surname { get; set; }

    }

    public class Company
    {
        public string CompanyName;
        public string CompanyCode;
        public  List<Person> People = new List<Person>();

        public Company() {
            CompanyName = "Acme";
            CompanyCode = "CODE13";

            Person x = new Person();
            x.Name = "John";
            x.Surname = "Wick";

            People.Add(x);
            People.Add(x);
            People.Add(x);

        }


    }
}
