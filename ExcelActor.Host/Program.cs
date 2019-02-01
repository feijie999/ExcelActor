using System;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Orleans;
using ExcelActor.Grain;
using Microsoft.Extensions.Configuration;
using Orleans.Configuration;
using Orleans.Hosting;
using Microsoft.Extensions.Logging;

namespace ExcelActor.Host
{
   public class Program
    {
       public static async Task Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json");

            var config = builder.Build();
            var silo = new SiloHostBuilder()
                .Configure<EndpointOptions>(options =>
                {
                    options.AdvertisedIPAddress = IPAddress.Loopback;
                })
                //.ConfigureEndpoints(siloPort: siloPort, gatewayPort: gatewayPort, listenOnAnyHostAddress: true,advertisedIP: IPAddress.Loopback)
                .Configure<ClusterOptions>(options =>
                {
                    options.ClusterId = "dev";
                    options.ServiceId = "excelapp";
                })
                .ConfigureApplicationParts(parts => parts.AddApplicationPart(typeof(ExcelGrain).Assembly).WithReferences())
                .UseAdoNetClustering(option =>
                {
                    option.ConnectionString = config["ConnectionStrings:OrleansCluster"];
                    option.Invariant = "MySql.Data.MySqlClient";
                })
                .AddAdoNetGrainStorageAsDefault(option =>
                {
                    option.Invariant = "MySql.Data.MySqlClient";
                    option.ConnectionString = config["ConnectionStrings:OrleansGrain"];

                })
                .ConfigureLogging(logging => logging.AddConsole())
                .Build();

            var stopEvent = new ManualResetEvent(false);
            Console.CancelKeyPress += (sender, eventArgs) =>
            {
                eventArgs.Cancel = false;
                stopEvent.Set();
            };

            Console.WriteLine("Starting");
            await silo.StartAsync();
            Console.WriteLine("Started");

            stopEvent.WaitOne();
            Console.WriteLine("Shutting down");
            await silo.StopAsync();
        }
    }
}
