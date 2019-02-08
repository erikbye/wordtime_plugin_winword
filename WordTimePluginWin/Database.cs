using System.Data.SQLite;

namespace WordTimePluginWin
{
    class Database
    {
        private static Database instance = null;
        private string _dbfile;

        SQLiteConnection dbConnection;
  
        private Database()
        {                 
            var homepath = System.Environment.GetEnvironmentVariable("homepath");
            var homedrive = System.Environment.GetEnvironmentVariable("homedrive");

            _dbfile = homedrive + homepath + "/wordtime.db";

            SQLiteConnection.CreateFile(_dbfile);            
        }

        public static Database Instance {
            get 
            {
                if (instance == null) instance = new Database();
                return instance;
            }
        }

        public void Connect()
        {
            using (dbConnection = new SQLiteConnection("Data Source=" + _dbfile + ";Version=3;"))
            {
                dbConnection.Open();

                var command = new SQLiteCommand("CREATE TABLE documents (id integer primary key, document TEXT, project TEXT)", dbConnection);
                command.ExecuteNonQuery();

                command = new SQLiteCommand("CREATE TABLE projects (id integer primary key, project TEXT)", dbConnection);
                command.ExecuteNonQuery();

                command = new SQLiteCommand("CREATE TABLE time (id integer primary key, heartbeat TEXT)", dbConnection);
                command.ExecuteNonQuery();
            }


            
            
            

            // string sql = "CREATE TABLE projects (document TEXT, project TEXT)";
        }
    }
}
