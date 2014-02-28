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
                            folderPath TEXT NOT NULL,
                            entryId TEXT NOT NULL,
                            dataHash TEXT NOT NULL,
                            key INTEGER NOT NULL,
                            uploaded DATE NOT NULL,
                            PRIMARY KEY (folderPath, entryId));
                        CREATE TABLE attachments (
                            folderPath TEXT NOT NULL,
                            entryId TEXT NOT NULL,
                            fileName TEXT NOT NULL,
                            fileSize INTEGER NOT NULL,
                            emailDataHash TEXT NOT NULL,
                            dataHash TEXT NOT NULL,
                            uploaded DATE NOT NULL,
                            PRIMARY KEY (folderPath, entryId, fileName, fileSize));
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
        public void AddEmail(string folderPath, string entryId, string dataHash, long key, DateTime uploaded)
        {
            Logger.DebugFormat("Starting database email addition: {0}\\{1}", folderPath, entryId);
            // Make sure that the uploaded date is always UTC, because sqlite has no concept of Time-Zones
            // See http://www.sqlite.org/datatype3.html
            if (null != uploaded)
            {
                uploaded = ((DateTime)uploaded).ToUniversalTime();
            }

            if (String.IsNullOrEmpty(dataHash))
            {
                Logger.WarnFormat("Bad dataHash for {0}\\{1}", folderPath, entryId);
                return;
            }

            // Insert into database.
            string command =
                @"INSERT OR REPLACE INTO emails (folderPath, entryId, dataHash, key, uploaded)
                    VALUES (@folderPath, @entryId, @dataHash, @key, @uploaded)";
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            parameters.Add("entryId", entryId);
            parameters.Add("dataHash", dataHash);
            parameters.Add("key", key);
            parameters.Add("uploaded", uploaded);
            ExecuteSQLAction(command, parameters);
            Logger.DebugFormat("Completed database email addition: {0}\\{1}", folderPath, entryId);
        }

        /// <summary>
        /// Remove a file from the database.
        /// </summary>
        public void RemoveEmail(string folderPath, string entryId)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            parameters.Add("entryId", entryId);
            ExecuteSQLAction("DELETE FROM emails WHERE folderPath=@folderPath AND entryId=@entryId", parameters);
            ExecuteSQLAction("DELETE FROM attachments WHERE folderPath=@folderPath AND entryId=@entryId", parameters);
        }

        /// <summary>
        /// Remove all emails (and attachments) from the database.
        /// </summary>
        public void RemoveAllEmails()
        {
            ExecuteSQLAction("DELETE FROM emails", null);
            ExecuteSQLAction("DELETE FROM attachments", null);
        }

        /// <summary>
        /// Get the time at which the file was uploaded.
        /// </summary>
        public DateTime? GetEmailUploadedDate(string folderPath, string entryId)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            parameters.Add("entryId", entryId);
            object obj = ExecuteScalarSQLFunction("SELECT uploaded FROM emails WHERE folderPath=@folderPath AND entryId=@entryId", parameters);
            if (null != obj)
            {
                return ((DateTime)obj).ToUniversalTime();
            }
            return null;
        }

        /// <summary>
        /// Get data hash for an email in the database.
        /// </summary>
        public string GetEmailDataHash(string folderPath, string entryId)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            parameters.Add("entryId", entryId);
            return (string)ExecuteScalarSQLFunction("SELECT dataHash FROM emails WHERE folderPath=@folderPath AND entryId=@entryId", parameters);
        }

        /// <summary>
        /// <summary>
        /// Checks whether the database contains a given email.
        /// </summary>
        public bool ContainsEmail(string folderPath, string entryId)
        {
            return null != GetEmailDataHash(folderPath, entryId);
        }

        /// <summary>
        /// List all email entry ids for folder.
        /// </summary>
        public HashSet<string> ListEntryIds(string folderPath)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            List<NameValueCollection> rowList = ExecuteListSQLFunction("SELECT entryId FROM emails WHERE folderPath=@folderPath", parameters);
            HashSet<string> entryIdSet = new HashSet<string>();
            if (rowList != null)
            {
                foreach (NameValueCollection nameValueCollection in rowList)
                {
                    string entryId = nameValueCollection.Get("entryId");
                    if (!string.IsNullOrWhiteSpace(entryId))
                    {
                        entryIdSet.Add(entryId);
                    }
                }
            }
            return entryIdSet;
        }
        
        /// <summary>
        /// List all email data hashes for folder.
        /// </summary>
        public HashSet<string> ListDistinctFolders()
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            List<NameValueCollection> rowList = ExecuteListSQLFunction("SELECT DISTINCT folderPath FROM emails", parameters);
            HashSet<string> folderPathList = new HashSet<string>();
            if (rowList != null)
            {
                foreach (NameValueCollection nameValueCollection in rowList)
                {
                    string folderPath = nameValueCollection.Get("folderPath");
                    if (!string.IsNullOrWhiteSpace(folderPath))
                    {
                        folderPathList.Add(folderPath);
                    }
                }
            }
            return folderPathList;
        }

        /// <summary>
        /// Get the time at which the file was uploaded.
        /// </summary>
        public DateTime? GetLastUploadedDate(string folderPath)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            object obj = ExecuteScalarSQLFunction("SELECT uploaded FROM emails WHERE folderPath=@folderPath ORDER BY uploaded DESC LIMIT 1", parameters);
            if (null != obj && obj is DateTime)
            {
                return ((DateTime)obj).ToUniversalTime();
            }
            return null;
        }
        
        /// <summary>
        /// Add a file to the database.
        /// If checksum is not null, it will be used for the database entry
        /// </summary>
        public void AddAttachment(string folderPath, string entryId, string fileName, long fileSize, string emailDataHash, string dataHash, DateTime uploaded)
        {
            Logger.DebugFormat("Starting database attachment addition: {0}\\{1}\\{2}", folderPath, entryId, fileName);
            // Make sure that the uploaded date is always UTC, because sqlite has no concept of Time-Zones
            // See http://www.sqlite.org/datatype3.html
            if (null != uploaded)
            {
                uploaded = ((DateTime)uploaded).ToUniversalTime();
            }

            if (String.IsNullOrEmpty(emailDataHash))
            {
                Logger.WarnFormat("Bad emailDataHash for {0}\\{1}\\{2}", folderPath, entryId, fileName);
                return;
            }
            
            if (String.IsNullOrEmpty(dataHash))
            {
                Logger.WarnFormat("Bad dataHash for {0}\\{1}\\{2}", folderPath, entryId, fileName);
                return;
            }

            // Insert into database.
            string command =
                @"INSERT OR REPLACE INTO attachments (folderPath, entryId, fileName, fileSize, emailDataHash, dataHash, uploaded)
                    VALUES (@folderPath, @entryId, @fileName, @fileSize, @emailDataHash, @dataHash, @uploaded)";
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            parameters.Add("entryId", entryId);
            parameters.Add("fileName", fileName);
            parameters.Add("fileSize", fileSize);
            parameters.Add("emailDataHash", emailDataHash);
            parameters.Add("dataHash", dataHash);
            parameters.Add("uploaded", uploaded);
            ExecuteSQLAction(command, parameters);
            Logger.DebugFormat("Completed database attachment addition: {0}\\{1}\\{2}", folderPath, entryId, fileName);
        }

        /// <summary>
        /// Get the time at which the file was uploaded.
        /// </summary>
        public DateTime? GetAttachmentUploadedDate(string folderPath, string entryId, string fileName, long fileSize)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            parameters.Add("entryId", entryId);
            parameters.Add("fileName", fileName);
            parameters.Add("fileSize", fileSize);
            object obj = ExecuteScalarSQLFunction(
                @"SELECT uploaded FROM attachments WHERE folderPath=@folderPath AND entryId=@entryId AND
                    fileName=@fileName AND fileSize=@fileSize", parameters);
            if (null != obj)
            {
                return ((DateTime)obj).ToUniversalTime();
            }
            return null;
        }

        /// <summary>
        /// Checks whether the database contains a given email.
        /// </summary>
        public bool ContainsAttachment(string folderPath, string entryId, string fileName, long fileSize)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            parameters.Add("entryId", entryId);
            parameters.Add("fileName", fileName);
            parameters.Add("fileSize", fileSize);
            return null != ExecuteScalarSQLFunction(
                @"SELECT folderPath FROM attachments WHERE folderPath=@folderPath AND entryId=@entryId AND
                    fileName=@fileName AND fileSize=@fileSize", parameters);
        }

        /// <summary>
        /// Count the number of email attachments.
        /// </summary>
        public int CountAttachments(string folderPath, string entryId)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("folderPath", folderPath);
            parameters.Add("entryId", entryId);
            object obj = ExecuteScalarSQLFunction(
                @"SELECT count(*) FROM attachments WHERE folderPath=@folderPath AND entryId=@entryId", parameters);
            return (int)(long)obj;
        }
        
        /// <summary>
        /// Get client ID.
        /// </summary>
        public string GetClientId()
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("key", "ClientId");
            return (string)ExecuteScalarSQLFunction("SELECT value FROM general WHERE key=@key", parameters);
        }

        /// <summary>
        /// Set the client ID (overwrites).
        /// </summary>
        public void SetClientId(string ClientId)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("key", "ClientId");
            parameters.Add("value", ClientId);
            ExecuteSQLAction("INSERT OR REPLACE INTO general (key, value) VALUES (@key, @value)", parameters);
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
