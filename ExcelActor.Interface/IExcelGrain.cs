using System;
using System.Threading.Tasks;
using Orleans;

namespace ExcelActor.Interface
{
    public interface IExcelGrain : IGrainWithIntegerKey
    {
        Task Load(byte[] excelBytes);

        Task<string> ExportAllToText();

        Task<string> Test(string name);
    }
}
