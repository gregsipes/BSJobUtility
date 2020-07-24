using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace BSGlobals
{
    public static class DataIO
    {
        public static void LogException(Exception ex, string message, string jobName)
        {
            Mail.SendMail(message, ex.ToString(), false);
            WriteToJobLog(JobLogMessageType.ERROR, ex.ToString(), jobName);
        }

        public static void WriteToJobLog(JobLogMessageType type, string message, string jobName)
        {
            Console.WriteLine($"{DateTime.Now.ToString()} {type.ToString("g"),-7}  Message: {message}");


            using (SqlCommand command = new SqlCommand())
            {
                try
                {
                    command.Connection = new SqlConnection(Config.GetConnectionStringTo(DatabaseConnectionStringNames.EventLogs));
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "dbo.InsertJobLog";
                    command.Parameters.Add(new SqlParameter("@JobName", jobName));
                    command.Parameters.Add(new SqlParameter("@MessageType", type.ToString("d")));
                    command.Parameters.Add(new SqlParameter("@Message", message));

                    command.Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error inserting log record. {ex.Message}");
                }
                finally
                {
                    if (command != null && command.Connection != null)
                        command.Connection.Close();
                }
            }
        }

        public static void ExecuteNonQuery(DatabaseConnectionStringNames connectionStringName, string commandText, params SqlParameter[] parameters)
        {
            ExecuteSQLCommand(connectionStringName, CommandType.StoredProcedure, commandText, parameters);
        }

        public static void ExecuteNonQuery(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            ExecuteSQLCommand(connectionStringName, commandType, commandText, parameters);
        }

        private static void ExecuteSQLCommand(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                try
                {
                    command.Connection = new SqlConnection(Config.GetConnectionStringTo(connectionStringName));
                    command.CommandType = commandType;
                    command.CommandText = commandText;
                    command.CommandTimeout = 0;

                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            command.Parameters.Add(param);
                        }
                    }

                    command.Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error executing query. {ex.Message}");
                }
                finally
                {
                    if (command != null && command.Connection != null)
                        command.Connection.Close();
                }
            }
        }

        public static List<Dictionary<string, object>> ExecuteSQL(DatabaseConnectionStringNames connectionStringName, string commandText, params SqlParameter[] parameters)
        {
            return ExecuteSQLQuery(connectionStringName, CommandType.StoredProcedure, commandText, parameters);
        }

        public static List<Dictionary<string, object>> ExecuteSQL(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            return ExecuteSQLQuery(connectionStringName, commandType, commandText, parameters);
        }

        private static List<Dictionary<string, object>> ExecuteSQLQuery(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {

            List<Dictionary<string, object>> rowsToReturn = new List<Dictionary<string, object>>();

            using (SqlDataReader reader = ExecuteQuery(connectionStringName, commandType, commandText, parameters))
            {
                while (reader.Read())
                {
                    Dictionary<string, object> dictionary = new Dictionary<string, object>();

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        dictionary.Add(reader.GetName(i), reader.GetValue(i));
                    }

                    rowsToReturn.Add(dictionary);
                }
            }

            return rowsToReturn;
        }

        public static SqlDataReader ExecuteQuery(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                command.Connection = new SqlConnection(Config.GetConnectionStringTo(connectionStringName));
                command.CommandType = commandType;
                command.CommandText = commandText;

                if (parameters != null)
                {
                    foreach (var param in parameters)
                    {
                        command.Parameters.Add(param);  //new SqlParameter(param.Key, param.Value)
                    }
                }
                command.Connection.Open();

                //https://docs.microsoft.com/en-us/dotnet/api/system.data.sqlclient.sqlcommand?redirectedfrom=MSDN&view=netframework-4.6
                // When using CommandBehavior.CloseConnection, the connection will be closed when the 
                // IDataReader is closed.
                SqlDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

                return reader;

            }
        }
    }
}
