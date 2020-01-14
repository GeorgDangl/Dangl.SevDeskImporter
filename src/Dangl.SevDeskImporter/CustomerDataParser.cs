using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Dangl.SevDeskImporter
{
    public class CustomerDataParser
    {
        public List<CustomerData> ParseCustomerDataFromExcel(string excelFilePath)
        {
            using (var fs = File.OpenRead(excelFilePath))
            {
                using (var excelDocument = new XLWorkbook(fs))
                {
                    var worksheet = excelDocument.Worksheet("Kunden");
                    var customerData = ReadCustomerData(worksheet);
                    return customerData;
                }
            }
        }

        private List<CustomerData> ReadCustomerData(IXLWorksheet worksheet)
        {
            var customerData = new List<CustomerData>();
            var currentLine = 1; // Skipping headings
            var emptyLines = 0;
            while (emptyLines < 10)
            {
                currentLine++;

                var parsedNumber = worksheet.Cell(currentLine, 1).GetValue<string>();
                if (string.IsNullOrWhiteSpace(parsedNumber))
                {
                    emptyLines++;
                    continue;
                }

                emptyLines = 0;

                var customer = new CustomerData
                {
                    CustomerNumber = parsedNumber
                };
                customerData.Add(customer);

                var parsedName = worksheet.Cell(currentLine, 2).GetValue<string>();
                var addressRaw = worksheet.Cell(currentLine, 4).GetValue<string>();
                var address = ParseAddress(addressRaw);

                customer.Name = address?.Name ?? parsedName;
                customer.Address = address;

                customer.Website = worksheet.Cell(currentLine, 3).GetValue<string>();
                customer.Notes = worksheet.Cell(currentLine, 7).GetValue<string>();

                customer.IsActiveCustomer = worksheet.Cell(currentLine, 8).GetValue<string>().Equals("Ja", System.StringComparison.InvariantCultureIgnoreCase);

                var contact01Raw = worksheet.Cell(currentLine, 5).GetValue<string>();
                if (!string.IsNullOrWhiteSpace(contact01Raw))
                {
                    var parsedContact = ParseContact(contact01Raw);
                    if (parsedContact != null)
                    {
                        if (customer.Contacts == null)
                        {
                            customer.Contacts = new List<CustomerContact>();
                        }

                        customer.Contacts.Add(parsedContact);
                    }
                }

                var contact02Raw = worksheet.Cell(currentLine, 6).GetValue<string>();
                if (!string.IsNullOrWhiteSpace(contact02Raw))
                {
                    var parsedContact = ParseContact(contact02Raw);
                    if (parsedContact != null)
                    {
                        if (customer.Contacts == null)
                        {
                            customer.Contacts = new List<CustomerContact>();
                        }

                        customer.Contacts.Add(parsedContact);
                    }
                }
            }

            return customerData;
        }

        private Address ParseAddress(string addressRaw)
        {
            if (string.IsNullOrWhiteSpace(addressRaw))
            {
                return null;
            }

            var splitAddress = Regex.Split(addressRaw, "\r\n?|\n")
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();

            if (!splitAddress.Any())
            {
                return null;
            }

            var address = new Address();

            address.Name = splitAddress[0];
            address.StreetWithNumber = splitAddress[1];
            address.ZipCode = splitAddress[2].TakeWhile(c => c != ' ').Select(c => c.ToString()).Aggregate((c, n) => c + n).Trim();
            address.City = splitAddress[2].SkipWhile(c => c != ' ').Select(c => c.ToString()).Aggregate((c, n) => c + n).Trim();
            address.Country = splitAddress.Count == 4 ? splitAddress[3] : "Deutschland";

            return address;
        }

        private CustomerContact ParseContact(string contactRaw)
        {
            if (string.IsNullOrWhiteSpace(contactRaw))
            {
                return null;
            }

            var splitContact = Regex.Split(contactRaw, "\r\n?|\n")
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Skip(1)
                .ToList();

            if (!contactRaw.Any())
            {
                return null;
            }

            var name = splitContact[0].Split(" ");
            if (name.Length != 3 && name.Length != 2 && name.Length != 4)
            {
                throw new System.NotImplementedException("Can not deserialize name: " + splitContact[0]);
            }

            var contact = new CustomerContact();

            if(name.Length == 2)
            {
                contact.FirstName = name[0];
                contact.LastName = name[1];
            }
            else if (name.Length == 3)
            {
                contact.Salutation = name[0];
                contact.FirstName = name[1];
                contact.LastName = name[2];
            }
            else if (name.Length == 4)
            {
                contact.Salutation = name[0];
                contact.FirstName = name[1];
                contact.LastName = name[2] + " " + name[3];
            }

            if (splitContact.Count > 1)
            {
                if (splitContact[1].StartsWith("Tel"))
                {
                    contact.Phone = splitContact[1].Substring(3).Trim('.').Trim(':').Trim();
                }
                else if (splitContact[1].StartsWith("Email"))
                {
                    contact.Email = splitContact[1].Substring(6).Trim('.').Trim(':');
                }
                else
                {
                    throw new System.NotImplementedException("Can not deserialize contact detail: " + splitContact[1]);
                }
            }

            if (splitContact.Count > 2)
            {
                if (splitContact[2].StartsWith("Tel"))
                {
                    contact.Phone = splitContact[2].Substring(3).Trim('.').Trim(':').Trim();
                }
                else if (splitContact[2].StartsWith("Email:"))
                {
                    contact.Email = splitContact[2].Substring(6).Trim('.').Trim(':');
                }
                else
                {
                    throw new System.NotImplementedException("Can not deserialize contact detail: " + splitContact[2]);
                }
            }

            return contact;
        }
    }
}
