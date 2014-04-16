using System.Data;
using System.Security.Cryptography.X509Certificates;

namespace PMSImporter
{
    public interface IDataSource
    {
        DataSet ReadData(string fileName);
    }
}
