using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelActor.Interface;

namespace ExcelActor.Grain
{
    public class ExcelGrain : Orleans.Grain<object>,IExcelGrain
    {
        public Task Load(byte[] excelBytes)
        {
            State = excelBytes;
            return Task.CompletedTask;
        }

        public async Task<string> Test(string name)
        {
            await Task.Delay(3000);
            Console.WriteLine(name);
            return "Hi " + name;
        }
    }
}
