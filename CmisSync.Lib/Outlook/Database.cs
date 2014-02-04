using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SQLite;
using System.IO;
using System.Security.Cryptography;
using log4net;

namespace CmisSync.Lib.Outlook
{

    /// <summary>
    /// Database to cache remote information from Oris4.
    /// Implemented with SQLite.
    /// </summary>
    public class Database : IDisposable
    {
        /// <summary>
        /// Log.
        /// </summary>
        private static readonly ILog Logger = LogManager.GetLogger(typeof(Database));


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
        public Database(string dataPath)
        {
            this.databaseFileName = dataPath;
        }


        /// <summary>
        /// Destructor.
        /// </summary>
        ~Database()
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
                            entryId TEXT PRIMARY KEY,
                            folderPath TEXT,
                            uploaded DATE,
                            dataHash TEXT);
                        CREATE TABLE attachments (
                            entryId TEXT NOT NULL,
                            position INTEGER NOT NULL,
                            fileName TEXT,
                            uploaded DATE,
                            dataHash TEXT,
                            PRIMARY KEY (entryId, position));
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
        public void AddEntry(string entryId, string folderPath, DateTime uploaded, string dataHash)
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
                @"INSERT OR REPLACE INTO emails (entryId, folderPath, uploaded, dataHash)
                    VALUES (@entryId, @folderPath, @uploaded, @dataHash)";
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("entryId", entryId);
            parameters.Add("folderPath", folderPath);
            parameters.Add("uploaded", uploaded);
            parameters.Add("dataHash", dataHash);
            ExecuteSQLAction(command, parameters);
            Logger.DebugFormat("Completed database email addition: {0}\\{1}", folderPath, entryId);
        }

        /// <summary>
        /// Remove a file from the database.
        /// </summary>
        public void RemoveFile(string entryId)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("entryId", entryId);
            ExecuteSQLAction("DELETE FROM emails WHERE entryId=@entryId", parameters);
        }

        /// <summary>
        /// Get the time at which the file was uploaded.
        /// </summary>
        public DateTime? GetUploadedDate(string entryId)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("entryId", entryId);
            object obj = ExecuteSQLFunction("SELECT uploaded FROM emails WHERE entryId=@entryId", parameters);
            if (null != obj)
            {
                obj = ((DateTime)obj).ToUniversalTime();
            }
            return (DateTime?)obj;
        }

        /// <summary>
        /// Checks whether the database contains a given email.
        /// </summary>
        public bool ContainsEmail(string entryId)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("entryId", entryId);
            return null != ExecuteSQLFunction("SELECT entryId FROM emails WHERE entryId=@entryId", parameters);
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
        private object ExecuteSQLFunction(string text, Dictionary<string, object> parameters)
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
