using System;
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
                    .ConfigureApplicationParts(parts => parts.AddApplicationPart(typeof(IExcelGrain).Assembly).WithReferences())
                    .ConfigureLogging(logging => logging.AddConsole());
                return builder;
            });
            while (true)
            {
                var input = Console.ReadLine();
                if (input == "exit")
                {
                    break;
                }
                var excelGrain = client.GetGrain<IExcelGrain>(1);
                //var result = await excelGrain.Test(input);
                //Console.WriteLine(result);
                using (var fs = new FileStream("template.xlsx", FileMode.Open, FileAccess.Read))
                {
                    var bytes = new byte[fs.Length];

                    await fs.ReadAsync(bytes);
                    await excelGrain.Load(bytes);
                }

               var json = await excelGrain.ExportAllToText();
                Console.WriteLine(json);

            }
            Console.ReadKey();
        }
    }
}
