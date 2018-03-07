using Netplate;

using System;
using System.Collections.Generic;

namespace NetplateTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string template = @"Dear [Title][~ ][Forename][~ ][Surname],";
            Console.WriteLine(template);
            Console.WriteLine("\n----------\n");

            var data = new Dictionary<string, object>()
            {
                {"Title", "Mr"},
                {"Forename", "John"},
                {"Surname", "Doe"}
            };
            string result = MailMerge.Merge(template, data);
            Console.WriteLine(result);

            Console.ReadLine();
        }
    }
}
