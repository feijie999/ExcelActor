using Orleans;

namespace ExcelActor.Client
{
    public interface IClientFactory
    {
        IClusterClient Create();
    }
}