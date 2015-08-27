namespace SQLServerForExcel_Addin
{
    public interface IDataConnectionConfiguration
    {
        string GetSelectedSource();
        void SaveSelectedSource(string provider);

        string GetSelectedProvider();
        void SaveSelectedProvider(string provider);
    }
}