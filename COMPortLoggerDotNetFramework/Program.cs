using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;

namespace COMPortLoggerDotNetFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("~ COM Port logger 5000 ~");

            // Get and print all port names
            var portInfo = GetAllComPortInfo();
            Console.WriteLine("Type the number of the port that you want to use:");
            var id = 1;
            foreach (var onePort in portInfo)
            {
                Console.WriteLine($"{id} - {onePort.Name} - {onePort.Caption}");
                id++;
            }

            // Retrieve the port ID from the user
            var portIdInput = Console.ReadLine();
            if (!int.TryParse(portIdInput, out var portId) || portId > portInfo.Count || portId <= 0)
            {
                Console.WriteLine("Invalid port ID! Relaunch this app... It's your own fault... You stupid >:(");
                return;
            }

            // Init the port
            var baudRate = 9600;
            var selectedPort = portInfo[portId - 1];
            var fileName = $@"{Environment.CurrentDirectory}\{DateTime.Now:yyyy_MM_dd__hh_mm_ss}_log.txt";
            Console.WriteLine($"Using port '{selectedPort.Name}' ({selectedPort.Caption}) with baud rate {baudRate}. Writing all data to file '{fileName}'");
            var port = new SerialPort(selectedPort.Name, baudRate);
            port.DataReceived += (sender, eventArgs) =>
            {
                var newData = port.ReadExisting();
                var sanitizedData = newData.Replace("\r", "").Replace("\n", "");
                Console.WriteLine($"Received data: '{sanitizedData}'");
                File.AppendAllText(fileName, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss},{newData}");
            };

            port.Open();

            Console.WriteLine("Type 'q' to quit...");
            while (Console.ReadLine() != "q") {}

            Console.WriteLine("Closing the port... Goodbye :)");
            port.Close();
        }

        public static List<PortInfo> GetAllComPortInfo()
        {
            using (var searcher = new ManagementObjectSearcher
                ("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portInfo = searcher.Get().Cast<ManagementBaseObject>();
                return portInfo.Select(one => new PortInfo(one["Caption"].ToString())).ToList();
            }
        }

        public class PortInfo
        {
            public string Name { get; set; }
            public string Caption { get; set; }

            public PortInfo(string caption)
            {
                Name = ExtractName(caption);
                Caption = caption;
            }

            private string ExtractName(string caption)
            {
                var matchInfo = Regex.Match(caption, @"(?<=\().+?(?=\))");
                if (!matchInfo.Success)
                {
                    return "No name found! You cannot use this port :(";
                }

                return matchInfo.Groups[0].Value;
            }
        }
    }
}
