using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using CsvHelper;
using Fclp;
using NLog;
using NLog.Config;
using NLog.Targets;
using ServiceStack;

namespace KenConsole
{
    internal class Program
    {
        private static Logger _logger;

        private static FluentCommandLineParser<Args> _fluentCommandLineParser;

        private static readonly string BaseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        private static void SetupNLog()
        {
            if (File.Exists(Path.Combine(BaseDirectory, "Nlog.config")))
            {
                return;
            }

            var config = new LoggingConfiguration();
            var loglevel = LogLevel.Info;

            var layout = @"${message}";

            var consoleTarget = new ColoredConsoleTarget();

            config.AddTarget("console", consoleTarget);

            consoleTarget.Layout = layout;

            var rule1 = new LoggingRule("*", loglevel, consoleTarget);
            config.LoggingRules.Add(rule1);

            LogManager.Configuration = config;
        }

        private static void Main(string[] args)
        {
            SetupNLog();

            _logger = LogManager.GetLogger("MFTECmd");

            _fluentCommandLineParser = new FluentCommandLineParser<Args>
            {
                IsCaseSensitive = false
            };

            _fluentCommandLineParser.Setup(arg => arg.FileName)
                .As('f')
                .WithDescription("File to process. This or -d required\r\n");

            _fluentCommandLineParser.Setup(arg => arg.Directory)
                .As('d')
                .WithDescription("Directory to process. This or -f required");

            _fluentCommandLineParser.Setup(arg => arg.CsvOut)
                .As("csv")
                .WithDescription(
                    "Directory to save CSV formatted results to. Required\r\n");


            _fluentCommandLineParser.Setup(arg => arg.Debug)
                .As("debug")
                .WithDescription("Show debug information during processing").SetDefault(false);

            _fluentCommandLineParser.Setup(arg => arg.Trace)
                .As("trace")
                .WithDescription("Show trace information during processing\r\n").SetDefault(false);

            var header =
                $"EmailCounter version {Assembly.GetExecutingAssembly().GetName().Version}" +
                "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)";

            var footer = @"Examples: EmailCounter.exe -f ""C:\Temp\someFile.txt"" --csv ""c:\temp\out"" " +
                         "\r\n\t " +
                         "\r\n\t" +
                         "  Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes\r\n";

            _fluentCommandLineParser.SetupHelp("?", "help")
                .WithHeader(header)
                .Callback(text => _logger.Info(text + "\r\n" + footer));

            var result = _fluentCommandLineParser.Parse(args);

            if (result.HelpCalled)
            {
                return;
            }

            if (result.HasErrors)
            {
                _logger.Error("");
                _logger.Error(result.ErrorText);

                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);

                return;
            }

            if (_fluentCommandLineParser.Object.FileName.IsNullOrEmpty() &&
                _fluentCommandLineParser.Object.Directory.IsNullOrEmpty())
            {
                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);

                _logger.Warn("-f or -d is required. Exiting");
                return;
            }

            var files = new List<string>();

            if (_fluentCommandLineParser.Object.FileName.IsNullOrEmpty() == false)
            {
                if (File.Exists(_fluentCommandLineParser.Object.FileName))
                {
                    files.Add(_fluentCommandLineParser.Object.FileName);
                }
                else
                {
                    _logger.Warn($"File '{_fluentCommandLineParser.Object.FileName}' does not exist. Exiting");
                    Environment.Exit(0);
                }
            }
            else
            {
                if (Directory.Exists(_fluentCommandLineParser.Object.Directory))
                {
                    files = Directory.GetFiles(_fluentCommandLineParser.Object.Directory, "*",
                        SearchOption.AllDirectories).ToList();
                }
                else
                {
                    _logger.Warn($"Directory '{_fluentCommandLineParser.Object.Directory}' does not exist. Exiting");
                    Environment.Exit(0);
                }
            }

            _logger.Info(header);
            _logger.Info("");
            _logger.Info($"Command line: {string.Join(" ", Environment.GetCommandLineArgs().Skip(1))}\r\n");

            if (_fluentCommandLineParser.Object.Debug)
            {
                LogManager.Configuration.LoggingRules.First().EnableLoggingForLevel(LogLevel.Debug);
            }

            if (_fluentCommandLineParser.Object.Trace)
            {
                LogManager.Configuration.LoggingRules.First().EnableLoggingForLevel(LogLevel.Trace);
            }

            LogManager.ReconfigExistingLoggers();

            var regexObj = new Regex(@"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", RegexOptions.IgnoreCase);

            var outName = $"{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_EmailCounter_Output.csv";

            var outFile = Path.Combine(_fluentCommandLineParser.Object.CsvOut, outName);

            if (Directory.Exists(_fluentCommandLineParser.Object.CsvOut) == false)
            {
                Directory.CreateDirectory(_fluentCommandLineParser.Object.CsvOut);
            }

            _logger.Warn($"CSV output will be saved to '{outFile}'\r\n");

            var swCsv = new StreamWriter(outFile, false, Encoding.UTF8);

            var _csvWriter = new CsvWriter(swCsv);

            var foo = _csvWriter.Configuration.AutoMap<CsvOut>();

            _csvWriter.Configuration.RegisterClassMap(foo);
            _csvWriter.WriteHeader<CsvOut>();
            _csvWriter.NextRecord();

            var emails = new Dictionary<string, int>();

            _logger.Fatal($"Files found: {files.Count:N0}");
            Console.WriteLine();

            foreach (var file in files)
            {
                try
                {
                    _logger.Info($"Processing '{file}'");

                    foreach (var readLine in File.ReadLines(file))
                    {
                        var matchResult = regexObj.Match(readLine);
                        while (matchResult.Success)
                        {
                            var key = matchResult.Value.ToLowerInvariant();
                            if (emails.ContainsKey(key) == false)
                            {
                                _logger.Debug($"Found new email address '{key}'!");
                                emails.Add(key, 0);
                            }

                            emails[key] += 1;

                            matchResult = matchResult.NextMatch();
                        }
                    }
                }
                catch (Exception e)
                {
                    _logger.Error($"Unable to process file '{file}'. Error: {e.Message}");
                }
            }

            _logger.Debug("Writing results to CSV...");
            foreach (var email in emails.OrderBy(t => t.Value))
            {
                var csvo = new CsvOut();
                csvo.EmailAddress = email.Key;
                csvo.Count = email.Value;

                _csvWriter.WriteRecord(csvo);
                _csvWriter.NextRecord();
            }

            _csvWriter.Flush();
            swCsv.Flush();
            swCsv.Close();

            var suffix = "s";
            if (files.Count == 1)
            {
                suffix = string.Empty;
            }

            Console.WriteLine();
            _logger.Fatal($"Finished. Found {emails.Count:N0} unique emails in {files.Count:N0} file{suffix}");
            Console.WriteLine();
        }

        public class Args
        {
            public string FileName { get; set; }
            public string Directory { get; set; }
            public string CsvOut { get; set; }

            public bool Debug { get; set; }
            public bool Trace { get; set; }
        }
    }
}


public class CsvOut
{
    public string EmailAddress { get; set; }
    public int Count { get; set; }
}