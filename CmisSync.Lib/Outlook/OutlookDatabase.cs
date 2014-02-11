using log4net;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data.SQLite;
using System.IO;

namespace CmisSync.Lib.Outlook
{

    /// <summary>
    /// Database to cache remote information from Oris4.
    /// Implemented with SQLite.
    /// </summary>
    public class OutlookDatabase : IDisposable
    {
        /// <summary>
        /// Log.
        /// </summary>
        private static readonly ILog Logger = LogManager.GetLogger(typeof(OutlookDatabase));


        /// <summary>
        /// Name of the SQLite database file.
        /// </summary>
        private string databaseFileName;


        /// <summary>
        ///  SQLite connection to the underlying database.
        /// </summary>
        private SQLiteConnection sqliteConnection;


        /// <summary>
        /// Track whether <c>Dispose</c> has been called.
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Constructor.
        /// </summary>
        public OutlookDatabase(string dataPath)
        {
            this.databaseFileName = dataPath;
        }


        /// <summary>
        /// Destructor.
        /// </summary>
        ~OutlookDatabase()
        {
            Dispose(false);
        }


        /// <summary>
        /// Implement IDisposable interface. 
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        /// <summary>
        /// Dispose pattern implementation.
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (this.sqliteConnection != null)
                    {
                        this.sqliteConnection.Dispose();
                    }
                }
                this.disposed = true;
            }
        }


        /// <summary>
        ///  Connection to the database.
        /// The sqliteConnection must not be used directly, used this method instead.
        /// </summary>
        public SQLiteConnection GetSQLiteConnection()
        {
            if (sqliteConnection == null || sqliteConnection.State == System.Data.ConnectionState.Broken)
            {
                try
                {
                    Logger.Info(String.Format("Checking whether database {0} exists", databaseFileName));
                    bool createDatabase = !File.Exists(databaseFileName);

                    sqliteConnection = new SQLiteConnection("Data Source=" + databaseFileName + ";PRAGMA journal_mode=WAL;");
                    sqliteConnection.Open();

                    if (createDatabase)
                    {
                        string command =
                            @"CREATE TABLE emails (
                            dataHash TEXT PRIMARY KEY,
                            folderPath TEXT,
                            uploaded DATE);
                        CREATE TABLE attachments (
                            emailDataHash TEXT NOT NULL,
                            dataHash TEXT NOT NULL,
                            fileName TEXT NOT NULL,
                            folderPath TEXT,
                            uploaded DATE,
                            PRIMARY KEY (emailDataHash, dataHash, fileName));
                        CREATE TABLE general (
                            key TEXT PRIMARY KEY,
                            value TEXT);";    /* Generic values */
                        ExecuteSQLAction(command, null);
                        Logger.Info("Database created");
                    }
                }
                catch (Exception e)
                {
                    Logger.Error("Error creating database: " + e.Message, e);
                    throw;
                }
            }
            return sqliteConnection;
        }

        //
        // Database operations.
        // 

        /// <summary>
        /// Add a file to the database.
        /// If checksum is not null, it will be used for the database entry
        /// </summary>
        public void AddEmail(string dataHash, string folderPath, DateTime uploaded)
        {
            Logger.DebugFormat("Starting database email addition: {0}\\{1}", folderPath, dataHash);
            // Make sure that the uploaded date is always UTC, because sqlite has no concept of Time-Zones
            // See http://www.sqlite.org/datatype3.html
            if (null != uploaded)
            {
                uploaded = ((DateTime)uploaded).ToUniversalTime();
            }

            if (String.IsNullOrEmpty(dataHash))
            {
                Logger.WarnFormat("Bad dataHash for {0}\\{1}", folderPath, dataHash);
                return;
            }

            // Insert into database.
            string command =
                @"INSERT OR REPLACE INTO emails (dataHash, folderPath, uploaded)
                    VALUES (@dataHash, @folderPath, @uploaded)";
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("dataHash", dataHash);
            parameters.Add("folderPath", folderPath);
            parameters.Add("uploaded", uploaded);
            ExecuteSQLAction(command, parameters);
            Logger.DebugFormat("Completed database email addition: {0}\\{1}", folderPath, dataHash);
        }

        /// <summary>
        /// Remove a file from the database.
        /// </summary>
        public void RemoveEmail(string dataHash)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("dataHash", dataHash);
            ExecuteSQLAction("DELETE FROM emails WHERE dataHash=@dataHash", parameters);
            ExecuteSQLAction("DELETE FROM attachments WHERE emailDataHash=@dataHash", parameters);
        }

        /// <summary>
        /// Get the time at which the file was uploaded.
        /// </summary>
        public DateTime? GetEmailUploadedDate(string dataHash)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("dataHash", dataHash);
            object obj = ExecuteScalarSQLFunction("SELECT uploaded FROM emails WHERE dataHash=@dataHash", parameters);
            if (null != obj)
            {
                obj = ((DateTime)obj).ToUniversalTime();
            }
            return (DateTime?)obj;
        }

        /// <summary>
        /// Checks whether the database contains a given email.
        /// </summary>
        public bool ContainsEmail(string dataHash)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("dataHash", dataHash);
            return null != ExecuteScalarSQLFunction("SELECT dataHash FROM emails WHERE dataHash=@dataHash", parameters);
        }

        /// <summary>
        /// List all email data hashes.
        /// </summary>
        public HashSet<string> ListEmailDataHashes()
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            List<NameValueCollection> rowList = ExecuteListSQLFunction("SELECT dataHash FROM emails", parameters);
            HashSet<string> dataHashList = new HashSet<string>();
            if (rowList != null)
            {
                foreach (NameValueCollection nameValueCollection in rowList)
                {
                    string dataHash = nameValueCollection.Get("dataHash");
                    if (!string.IsNullOrWhiteSpace(dataHash))
                    {
                        dataHashList.Add(dataHash);
                    }
                }
            }
            return dataHashList;
        }

        /// <summary>
        /// Add a file to the database.
        /// If checksum is not null, it will be used for the database entry
        /// </summary>
        public void AddAttachment(string emailDataHash, string dataHash, string fileName, string folderPath, DateTime uploaded)
        {
            Logger.DebugFormat("Starting database attachment addition: {0}\\{1}", folderPath, fileName);
            // Make sure that the uploaded date is always UTC, because sqlite has no concept of Time-Zones
            // See http://www.sqlite.org/datatype3.html
            if (null != uploaded)
            {
                uploaded = ((DateTime)uploaded).ToUniversalTime();
            }

            if (String.IsNullOrEmpty(emailDataHash))
            {
                Logger.WarnFormat("Bad emailDataHash for {0}\\{1}", folderPath, fileName);
                return;
            }
            
            if (String.IsNullOrEmpty(dataHash))
            {
                Logger.WarnFormat("Bad dataHash for {0}\\{1}", folderPath, fileName);
                return;
            }

            // Insert into database.
            string command =
                @"INSERT OR REPLACE INTO attachments (emailDataHash, dataHash, fileName, folderPath, uploaded)
                    VALUES (@emailDataHash, @dataHash, @fileName, @folderPath, @uploaded)";
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("emailDataHash", emailDataHash);
            parameters.Add("dataHash", dataHash);
            parameters.Add("fileName", fileName);
            parameters.Add("folderPath", folderPath);
            parameters.Add("uploaded", uploaded);
            ExecuteSQLAction(command, parameters);
            Logger.DebugFormat("Completed database attachment addition: {0}\\{1}", folderPath, fileName);
        }

        /// <summary>
        /// Get the time at which the file was uploaded.
        /// </summary>
        public DateTime? GetAttachmentUploadedDate(string emailDataHash, string dataHash, string fileName)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("emailDataHash", emailDataHash);
            parameters.Add("dataHash", dataHash);
            parameters.Add("fileName", fileName);
            object obj = ExecuteScalarSQLFunction(
                @"SELECT uploaded FROM attachments WHERE emailDataHash=@emailDataHash AND
                    dataHash=@dataHash AND fileName=@fileName", parameters);
            if (null != obj)
            {
                obj = ((DateTime)obj).ToUniversalTime();
            }
            return (DateTime?)obj;
        }

        /// <summary>
        /// Checks whether the database contains a given email.
        /// </summary>
        public bool ContainsAttachment(string emailDataHash, string dataHash, string fileName)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("emailDataHash", emailDataHash);
            parameters.Add("dataHash", dataHash);
            parameters.Add("fileName", fileName);
            return null != ExecuteScalarSQLFunction(
                @"SELECT dataHash FROM attachments WHERE emailDataHash=@emailDataHash AND
                    dataHash=@dataHash AND fileName=@fileName", parameters);
        }

        
        /// <summary>
        /// Helper method to execute an SQL command that does not return anything.
        /// </summary>
        /// <param name="text">SQL query, optionnally with @something parameters.</param>
        /// <param name="parameters">Parameters to replace in the SQL query.</param>
        private void ExecuteSQLAction(string text, Dictionary<string, object> parameters)
        {
            using (var command = new SQLiteCommand(GetSQLiteConnection()))
            {
                try
                {
                    ComposeSQLCommand(command, text, parameters);
                    command.ExecuteNonQuery();
                }
                catch (SQLiteException e)
                {
                    Logger.Error(String.Format("Could not execute SQL: {0}; {1}", text, string.Join(";", parameters)), e);
                    throw;
                }
            }
        }


        /// <summary>
        /// Helper method to execute an SQL command that returns something.
        /// </summary>
        /// <param name="text">SQL query, optionnally with @something parameters.</param>
        /// <param name="parameters">Parameters to replace in the SQL query.</param>
        private object ExecuteScalarSQLFunction(string text, Dictionary<string, object> parameters)
        {
            using (var command = new SQLiteCommand(GetSQLiteConnection()))
            {
                try
                {
                    ComposeSQLCommand(command, text, parameters);
                    return command.ExecuteScalar();
                }
                catch (SQLiteException e)
                {
                    Logger.Error(String.Format("Could not execute SQL: {0}; {1}", text, string.Join(";", parameters)), e);
                    throw;
                }
            }
        }

        /// <summary>
        /// Helper method to execute an SQL command that returns something.
        /// </summary>
        /// <param name="text">SQL query, optionnally with @something parameters.</param>
        /// <param name="parameters">Parameters to replace in the SQL query.</param>
        private List<NameValueCollection> ExecuteListSQLFunction(string text, Dictionary<string, object> parameters)
        {
            using (var command = new SQLiteCommand(GetSQLiteConnection()))
            {
                try
                {
                    ComposeSQLCommand(command, text, parameters);
                    SQLiteDataReader reader = command.ExecuteReader();
                    List<NameValueCollection> rowList = new List<NameValueCollection>();
                    while (reader.Read())
                    {
                        rowList.Add(reader.GetValues());
                    }
                    return rowList;
                }
                catch (SQLiteException e)
                {
                    Logger.Error(String.Format("Could not execute SQL: {0}; {1}", text, string.Join(";", parameters)), e);
                    throw;
                }
            }
        }

        /// <summary>
        /// Helper method to fill the parameters inside an SQL command.
        /// </summary>
        /// <param name="command">The SQL command object to fill. This method modifies it.</param>
        /// <param name="text">SQL query, optionnally with @something parameters.</param>
        /// <param name="parameters">Parameters to replace in the SQL query.</param>
        private void ComposeSQLCommand(SQLiteCommand command, string text, Dictionary<string, object> parameters)
        {
            command.CommandText = text;
            if (null != parameters)
            {
                foreach (KeyValuePair<string, object> pair in parameters)
                {
                    command.Parameters.AddWithValue(pair.Key, pair.Value);
                }
            }
        }
    }
}
