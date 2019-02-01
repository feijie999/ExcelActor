using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Orleans;

namespace ExcelActor.Interface
{
    public interface ITestGrain : IGrainWithIntegerKey
    {
        Task<int> Add(int i);
    }
}
