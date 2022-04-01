using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Diagnostics;

namespace SpoLvtTest
{
    internal class Program
    {
        //These values are set in the App.config file...
        private static readonly string SpoUrl = ConfigurationManager.AppSettings.Get("SpoUrl");
        private static readonly string SpoUsername = ConfigurationManager.AppSettings.Get("SpoUsername");
        private static readonly string SpoPassword = ConfigurationManager.AppSettings.Get("SpoPassword");
        private static readonly string SpoDocLibTitle = ConfigurationManager.AppSettings.Get("SpoDocLibTitle");
        private static readonly int SpoListItemID = int.Parse(ConfigurationManager.AppSettings.Get("SpoListItemId"));


        static void Main()
        {
            var logFactory = LoggerFactory.Create(builder => builder.AddSimpleConsole(opt =>
            {
                opt.IncludeScopes = true;
                opt.TimestampFormat = "[hh:mm:ss.fff] ";
                opt.SingleLine = true;
            }));

            ExecuteTests(logFactory.CreateLogger("SPO Tester"));

            logFactory.Dispose();

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey(true);
        }


        private static void ExecuteTests(ILogger logger)
        {
            var sw = Stopwatch.StartNew();

            using (var triggerLvtScope = logger.BeginScope("Testing simple LVT..."))
            {
                ExecuteSimpleExample(logger);
            }

            using (var docLibScope = logger.BeginScope("Testing document libary..."))
            {
                ExecuteDocLibTest(logger);
            }

            using (var listItemScope = logger.BeginScope("Testing list items..."))
            {
                ExecuteListItemTest(logger);
            }

            sw.Stop();

            logger.LogInformation($"Finished! Elasped time: {sw.Elapsed}");
        }

        private static void ExecuteDocLibTest(ILogger logger)
        {
            ClientContext clientContext = null;

            try
            {
                clientContext = ClientContextUtility.GetClientContext(SpoUrl, SpoUsername, SpoPassword);

                ListViewThresholdExamples.TestLvt_EnumerateListItems(logger, clientContext, SpoDocLibTitle);

                logger.LogInformation("Success!");
            }
            catch (Exception ex)
            {
                logger.LogError(1, $"Failed! {ex.GetType()}: {ex.Message}");
            }
            finally
            {
                clientContext?.Dispose();
                clientContext = null;
            }
        }

        private static void ExecuteListItemTest(ILogger logger)
        {
            ClientContext clientContext = null;

            try
            {
                clientContext = ClientContextUtility.GetClientContext(SpoUrl, SpoUsername, SpoPassword);

                ListViewThresholdExamples.TestLvt_GetSingleListItem(logger, clientContext, SpoDocLibTitle, SpoListItemID);

                logger.LogInformation("Success!");
            }
            catch (Exception ex)
            {
                logger.LogError(1, $"Failed! {ex.GetType()}: {ex.Message}");
            }
            finally
            {
                clientContext?.Dispose();
                clientContext = null;
            }
        }

        private static void ExecuteSimpleExample(ILogger logger)
        {
            ClientContext clientContext = null;

            try
            {
                clientContext = ClientContextUtility.GetClientContext(SpoUrl, SpoUsername, SpoPassword);

                ListViewThresholdExamples.TestLvt_SimpleExample(logger, clientContext, SpoDocLibTitle);

                logger.LogInformation("Success!");
            }
            catch (Exception ex)
            {
                logger.LogError(1, $"Failed! {ex.GetType()}: {ex.Message}");
            }
            finally
            {
                clientContext?.Dispose();
                clientContext = null;
            }
        }
    }
}
