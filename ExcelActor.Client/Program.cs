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
                var excelGrain = client.GetGrain<IExcelGrain>(0);
                var result = await excelGrain.Test(input);
                Console.WriteLine(result);
                if (Console.ReadLine() == "exit")
                {
                    break;
                }
            }
            Console.ReadKey();
        }
    }
}
