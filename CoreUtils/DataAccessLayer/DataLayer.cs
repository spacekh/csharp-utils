using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace CoreUtils.DataAccessLayer
{
    /// <summary>
    /// Summary description for DataLayer
    /// </summary>
    public class DataLayer
    {
        private SqlConnection dbConn;

        /// <summary>
        /// Creates a new DataLayer object with the given SQL Connection String
        /// </summary>
        /// <param name="dbConnectionString">The connection string to use (usually the name of the database)</param>
        public DataLayer(string dbConnectionString)
        {
            dbConn = new SqlConnection(dbConnectionString);
        }

        /// <summary>
        /// Creates a new SqlParameter with the given name, type, and value
        /// </summary>
        /// <param name="paramName">The name of the SQL parameter</param>
        /// <param name="sqlDbType">The SqlDbType of the parameter</param>
        /// <param name="paramval">The value that the parameter should take in a SQL query</param>
        /// <returns>A new SqlParameter with the given attributes</returns>
        public SqlParameter CreateParam(string paramName, SqlDbType sqlDbType, object paramval)
        {
            SqlParameter param = new SqlParameter(paramName, sqlDbType);
            if (paramval != null)
                param.Value = paramval;
            return param;
        }
        
        // Various datasets with different signatures
        #region Datasets

        /// <summary>
        /// Returns the DataSet returned by the given stored procedure
        /// </summary>
        /// <param name="spName">The stored procedure to run</param>
        /// <returns>The DataSet returned by the given stored procedure, or null if an error occurs</returns>
        public DataSet GetDataSet(string spName)
        {
            return GetDataSet(spName, (SqlParameter)null);
        }
        /// <summary>
        /// Returns the DataSet returned by the given stored procedure
        /// </summary>
        /// <param name="spName">The stored procedure to run</param>
        /// <param name="parameter">The parameter for this stored procedure</param>
        /// <returns>The DataSet returned by the given stored procedure, or null if an error occurs</returns>
        public DataSet GetDataSet(string spName, SqlParameter parameter)
        {
            List<SqlParameter> l = new List<SqlParameter>();
            if (parameter != null) l.Add(parameter);
            else l = null;
            return GetDataSet(spName, l);
        }
        /// <summary>
        /// Returns the DataSet returned by the given stored procedure
        /// </summary>
        /// <param name="spName">The stored procedure to execute</param>
        /// <param name="parameters">The parameters for the given stored procedure</param>
        /// <returns>The DataSet returned by the given stored procedure, or null if an error occurs</returns>
        public DataSet GetDataSet(string spName, SqlParameter[] parameters)
        {
            List<SqlParameter> l = new List<SqlParameter>();
            foreach (SqlParameter p in parameters)
            {
                l.Add(p);
            }
            return GetDataSet(spName, l);
        }
        /// <summary>
        /// Returns the DataSet returned by the given stored procedure
        /// </summary>
        /// <param name="spName">The stored procedure to run</param>
        /// <param name="parameters">The parameters for this stored procedure</param>
        /// <returns>The DataSet returned by the given stored procedure, or null if an error occurs</returns>
        public DataSet GetDataSet(string spName, List<SqlParameter> parameters)
        {
            try
            {
                SqlCommand cmd = new SqlCommand() { CommandType = CommandType.StoredProcedure, CommandText = spName };
                if (parameters != null)
                {
                    //for (int i = 0; i < parameters.Count; i++)
                    //    cmd.Parameters.Add(parameters[i]);
                    cmd.Parameters.AddRange(parameters.ToArray());
                }
                return GetDataSet(cmd);
            }
            catch (Exception)
            {
                // TODO: Log errors to file and email
                //sendMail.ErrorNotice("developers@pbjcal.org", "Error: GetDataset:List<SqlParameter> : DataLayer", ex.ToString());
                return null;
            }
        }
        /// <summary>
        /// Returns the DataSet returned by the given SqlCommand
        /// </summary>
        /// <param name="cmd">The SQL query to perform</param>
        /// <returns>The DataSet returned by the given SqlCommand.  If the DataSet includes a TABLE_NAMES table at the end, the tables will be named accordingly.</returns>
        public DataSet GetDataSet(SqlCommand cmd)
        {
            try
            {
                DataSet ds = new DataSet();
                cmd.Connection = dbConn;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                dbConn.Open();
                da.Fill(ds);
                //if the DataSet has a table at the end with table names, name the tables.
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable names = ds.Tables[ds.Tables.Count - 1];
                    try
                    {
                        if (names.Columns.Contains("TABLE_NAMES"))
                        {
                            for (int i = 0; i < ds.Tables.Count; i++)
                            {
                                ds.Tables[i].TableName = names.Rows[0][i].ToString();
                            }
                        }
                    }
                    catch (Exception) { /*if something goes wrong, just give up on naming tables*/ }
                }
                return ds;
            }
            catch (Exception)
            {
                // TODO: Log errors to file and email
                //sendMail.ErrorNotice("developers@pbjcal.org", "Error: GetDataSet SqlCommand", ex.ToString());
                return null;
            }
            finally
            {
                if (dbConn.State == ConnectionState.Open)
                    dbConn.Close();
            }

        }
        #endregion

        /// <summary>
        /// Executes the given stored procedure
        /// </summary>
        /// <param name="spName">the stored procedure to execute</param>
        /// <returns>The number of rows affected, or -1 if it fails</returns>
        public int ExecuteNonQuery(string spName)
        {
            return ExecuteNonQuery(spName, null);
        }

        /// <summary>
        /// Executes the given SqlCommand against the database
        /// </summary>
        /// <param name="command">the SqlCommand to execute</param>
        /// <returns>The number of rows affected, or -1 if it fails</returns>
        public int ExecuteNonQuery(SqlCommand command)
        {
            try
            {
                command.Connection = dbConn;
                dbConn.Open();
                return command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.WriteLine(command.CommandText);
                return -1;
            }
            finally
            {
                if (dbConn.State == ConnectionState.Open)
                    dbConn.Close();
            }
        }

        /// <summary>
        /// Executes a stored procedure that is not supposed to return any rows of data
        /// </summary>
        /// <param name="spName">the stored procedure to execute</param>
        /// <param name="parameters">the parameters to pass to the stored procedure</param>
        /// <returns>The number of rows affected; -1 if an error occurred</returns>
        public int ExecuteNonQuery(string spName, List<SqlParameter> parameters)
        {
            SqlCommand command = new SqlCommand(spName, dbConn)
            {
                CommandType = CommandType.StoredProcedure
            };
            command.Parameters.Clear();

            if (parameters != null)
            {
                command.Parameters.AddRange(parameters.ToArray());
            }
            return ExecuteNonQuery(command);
        }

        /// <summary>
        /// Executes a stored procedure that is only supposed to return a scalar value (not a table)
        /// </summary>
        /// <param name="spName">The stored procedure to execute</param>
        /// <param name="parameters">The list of parameters to pass to the SP</param>
        /// <returns>The value returned by the SP, or null if the SP execution fails.</returns>
        public object ExecuteScalar(string spName, List<SqlParameter> parameters)
        {
            SqlCommand command = new SqlCommand(spName, dbConn)
            {
                CommandType = CommandType.StoredProcedure
            };

            if (parameters != null) command.Parameters.AddRange(parameters.ToArray());

            return ExecuteScalar(command);
        }

        /// <summary>
        /// Executes a SqlCommand that is only supposed to return a scalar value (not a table)
        /// </summary>
        /// <param name="command">The command to execute</param>
        /// <returns>The value returned by the command, or null if the command execution fails.</returns>
        public object ExecuteScalar(SqlCommand command)
        {
            try
            {
                command.Connection = dbConn;
                dbConn.Open();
                return command.ExecuteScalar();
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {
                if (dbConn.State == ConnectionState.Open) dbConn.Close();
            }
        }
    }
}
