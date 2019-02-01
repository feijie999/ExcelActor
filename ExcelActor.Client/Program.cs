using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using ExcelActor.Interface;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Orleans;
using Orleans.Configuration;
using Orleans.Hosting;
using Orleans.Runtime;

namespace ExcelActor.Client
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();


            var client = await ClientFactory.Build(() =>
            {
                var builder = new ClientBuilder()
                    .UseLocalhostClustering()
                    .Configure<ClusterOptions>(options =>
                    {
                        options.ClusterId = "dev";
                        options.ServiceId = "excelapp";
                    })
                    .UseAdoNetClustering(option =>
                    {
                        option.ConnectionString = config["ConnectionStrings:OrleansCluster"];
                        option.Invariant = "MySql.Data.MySqlClient";
                    })
                    .ConfigureApplicationParts(parts =>
                        parts.AddApplicationPart(typeof(IExcelGrain).Assembly).WithReferences())
                    .ConfigureLogging(logging => logging.AddConsole());
                return builder;
            });
            //await TestExcel(client);
            await TestAdd(client);
            Console.ReadKey();
        }

        private static async Task TestAdd(IClusterClient client)
        {
            while (true)
            {
                await Task.Delay(500);
                var testGrain = client.GetGrain<ITestGrain>(0);
                var result = await testGrain.Add(1);
                Console.WriteLine(result);
            }
        }

        static async Task TestExcel(IClusterClient client)
        {
            while (true)
            {
                //var input = Console.ReadLine();
                //if (input == "exit")
                //{
                //    break;
                //}
                var excelGrain = client.GetGrain<IExcelGrain>(2);
                //var result = await excelGrain.Test(input);
                //Console.WriteLine(result);
                using (var fs = new FileStream("1.xlsx", FileMode.Open, FileAccess.Read))
                {
                    var bytes = new byte[fs.Length];

                    await fs.ReadAsync(bytes);
                    await excelGrain.Load(bytes);
                }

                var stopwatch = Stopwatch.StartNew();
                var json = await excelGrain.ExportAllToText();
                Console.WriteLine("耗时:" + stopwatch.ElapsedMilliseconds + "毫秒");
                Console.WriteLine(json);

            }
        }
    }
}
