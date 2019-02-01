using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using ExcelActor.Interface;
using Orleans;

namespace ExcelActor.Grain
{
    public class TestGrain :Grain<int> ,ITestGrain
    {
        public override async Task OnActivateAsync()
        {
            //await ReadStateAsync();
            await base.OnActivateAsync();
        }

        public async Task<int> Add(int i)
        {
            State += i;
            await Task.Delay(1000);
            Console.WriteLine(State);
            return State;
        }

        public override async Task OnDeactivateAsync()
        {
            await WriteStateAsync();
            await base.OnDeactivateAsync();
        }
    }
}
