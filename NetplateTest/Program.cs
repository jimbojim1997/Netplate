using Netplate;

using System;
using System.Collections.Generic;

namespace NetplateTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string template = @"[if [Min] == [Max]][~true]";
            Console.WriteLine(template);
            Console.WriteLine("\n------------------\n");

            var data = new Dictionary<string, object>()
            {
                {"Min", 100 },
                {"Max", 100 }
            };

            string result = MailMerge.Merge(template, data);
            Console.WriteLine(result);

            Console.ReadLine();
        }
    }
}
