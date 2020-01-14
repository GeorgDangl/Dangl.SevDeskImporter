using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dangl.SevDeskImporter
{
    class Program
    {
        private static async Task Main(string[] args)
        {
            Console.WriteLine("Please pass two arguments:");
            Console.WriteLine("1. The path to the input Excel file");
            Console.WriteLine("2. The path to the output SevDesk CSV file");

            if (args.Length != 2)
            {
                Console.WriteLine("Please provide exactly two arguments.");
                return;
            }

            try
            {
                var excelImportPath = args[0];
                var sevdeskCsvOutputPath = args[1];

                var customerData = new CustomerDataParser().ParseCustomerDataFromExcel(excelImportPath);

                foreach (var customer in customerData)
                {
                    var serialized = JsonConvert.SerializeObject(customer, Formatting.Indented);
                    Console.WriteLine(serialized);
                }

                using (var fs = File.Create(sevdeskCsvOutputPath))
                {
                    using (var sw = new StreamWriter(fs, Encoding.UTF8))
                    {
                        var customerCsv = GetSevdeskCustomerImportCsv(customerData);
                        await sw.WriteAsync(customerCsv);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private static string GetSevdeskCustomerImportCsv(List<CustomerData> customers)
        {
            // This is the CSV header from the sevDesk contact import template
            var headings = "Kunden-Nr.;Anrede;Titel;Nachname;Vorname;Organisation;Namenszusatz;Position;Kategorie;IBAN;BIC;Umsatzsteuer-ID;Strasse;PLZ;Ort;Land;Adress-Kategorie;Telefon;Telefon-Kategorie;Mobil;Fax;E-Mail;E-Mail-Kategorie;Webseite;Webseiten-Kategorie;Beschreibung;Geburtstag;Tags;Debitoren/Kreditoren-Nr."
                .Split(';')
                .ToList();

            var sevDeskCsv = headings.Aggregate((c, n) => c + ";" + n) + Environment.NewLine;

            foreach (var customer in customers)
            {
                var customerCsvData = new List<string>();
                foreach (var _ in headings)
                {
                    customerCsvData.Add(string.Empty);
                }

                SetDataForCustomerList(headings, customerCsvData, "Kunden-Nr.", customer.CustomerNumber);
                SetDataForCustomerList(headings, customerCsvData, "Organisation", customer.Name);
                SetDataForCustomerList(headings, customerCsvData, "Kategorie", customer.IsActiveCustomer ? "Kunde" : "Interessent");
                SetDataForCustomerList(headings, customerCsvData, "Strasse", customer.Address?.StreetWithNumber);
                SetDataForCustomerList(headings, customerCsvData, "PLZ", customer.Address?.ZipCode);
                SetDataForCustomerList(headings, customerCsvData, "Ort", customer.Address?.City);
                SetDataForCustomerList(headings, customerCsvData, "Land", customer.Address?.Country);
                SetDataForCustomerList(headings, customerCsvData, "Webseite", customer.Website);
                SetDataForCustomerList(headings, customerCsvData, "Beschreibung", customer.Notes);

                sevDeskCsv += customerCsvData.Aggregate((c, n) => c + ";" + n) + Environment.NewLine;

                if (customer.Contacts?.Any() ?? false)
                {
                    foreach (var contact in customer.Contacts)
                    {
                        var contactCsvData = new List<string>();
                        foreach (var _ in headings)
                        {
                            contactCsvData.Add(string.Empty);
                        }

                        SetDataForCustomerList(headings, contactCsvData, "Kunden-Nr.", customer.CustomerNumber);
                        if (!string.IsNullOrWhiteSpace(contact.Email))
                        {
                            SetDataForCustomerList(headings, contactCsvData, "E-Mail", contact.Email);
                        }
                        if (!string.IsNullOrWhiteSpace(contact.Phone))
                        {
                            SetDataForCustomerList(headings, contactCsvData, "Telefon", contact.Phone);
                        }
                        SetDataForCustomerList(headings, contactCsvData, "Vorname", contact.FirstName);
                        SetDataForCustomerList(headings, contactCsvData, "Nachname", contact.LastName);
                        SetDataForCustomerList(headings, contactCsvData, "Anrede", contact.Salutation);

                        sevDeskCsv += contactCsvData.Aggregate((c, n) => c + ";" + n) + Environment.NewLine;
                    }
                }
            }

            return sevDeskCsv;
        }

        private static void SetDataForCustomerList(List<string> headings, List<string> csvLine, string heading, string propertyValue)
        {
            var index = headings.IndexOf(heading);
            if (index < 0)
            {
                throw new NotImplementedException();
            }

            if (propertyValue == null)
            {
                return;
            }

            propertyValue = propertyValue.Replace("\"", string.Empty);

            // CSV supports line breaks when the value is wrapped in quotes
            propertyValue = (propertyValue.Contains("\r") || propertyValue.Contains("\n"))
                ? $"\"{propertyValue}\""
                : propertyValue;

            csvLine[index] = propertyValue.Replace(";", string.Empty);
        }
    }
}
