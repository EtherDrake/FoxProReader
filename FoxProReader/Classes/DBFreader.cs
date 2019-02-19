using System.Data;
using System.Data.OleDb;

namespace FoxProReader.Classes
{
    public class DBFreader
    {
        public static void Get(string SQL, ref DataSet set, string FilePath, string TableName)
        {
            string connetionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FilePath + "; Extended Properties=dBASE IV;";
            OleDbConnection Conn = new OleDbConnection(connetionString);
            Conn.Open();

            OleDbDataAdapter t1 = new OleDbDataAdapter(SQL, Conn);
            t1.Fill(set, TableName);

            Conn.Close();
        }
    }
}
