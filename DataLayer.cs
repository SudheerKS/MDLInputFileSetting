using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Oracle.DataAccess.Client;
using System.Collections;

namespace MDL
{
    public sealed class DataLayer
    {
        private string strConnection;
        private OracleConnection ocnConn;
        private OracleTransaction octTransaction;

        public DataLayer()
        { }

        public DataLayer(string strConnectionString)
        {
            strConnection = strConnectionString;
            ocnConn = new OracleConnection(strConnection);
            ocnConn.Open();
        }

        public DataLayer(string strConnectionString, bool bolTransaction)
        {
            strConnection = strConnectionString;
            ocnConn = new OracleConnection(strConnection);
            ocnConn.Open();

            if (bolTransaction)
            {
                if (octTransaction == null)
                {
                    octTransaction = ocnConn.BeginTransaction();
                }
            }
        }

        #region Create Oracle Connection
        /// <summary>
        /// Open Connections
        /// </summary>
        /// <returns>Return true if executed</returns>
        private bool CreateConnection()
        {
            try
            {
                ocnConn = new OracleConnection(strConnection);
                ocnConn.Open();
                return true;
            }
            catch (Exception expConnection)
            {
                throw expConnection;
            }
        }
        #endregion

        #region Check if conenction is open
        /// <summary>
        /// Check if conenction is open
        /// </summary>
        /// <returns></returns>
        public bool CheckConnectionOpen()
        {
            if (ocnConn != null && ocnConn.State == ConnectionState.Open)
                return true;
            return false;
        }
        #endregion

        #region Begin Transaction
        /// <summary>
        /// Check if conenction is open
        /// </summary>
        /// <returns></returns>
        public bool BeginTransaction()
        {
            if (ocnConn != null && ocnConn.State == ConnectionState.Open)
            {
                octTransaction = ocnConn.BeginTransaction();
                return true;
            }
            return false;
        }
        #endregion

        #region Close Oracle Connection
        /// <summary>
        /// Close Connections for connections that are maintained across queries
        /// </summary>
        /// <returns></returns>
        public bool CloseConnection(bool bolClose)
        {
            try
            {

                if (bolClose && ocnConn != null)
                {
                    if (ocnConn.State == ConnectionState.Open)
                    {
                        if (octTransaction != null)
                        {
                            RollBackMaintainConnection();
                            //CCH 070408
                            throw new Exception("A transaction connection cannot be closed using this function");
                            //CCH 070408 end
                        }
                        ocnConn.Close();
                        ocnConn.Dispose();
                        ocnConn = null;
                        return true;
                    }
                }
                return false;
            }
            catch (Exception expConnection)
            {
                throw expConnection;
            }
        }
        #endregion
        #region RollBack Database
        /// <summary>
        /// RollBack Database
        /// </summary>
        /// <returns></returns>
        private void RollBackMaintainConnection()
        {
            try
            {
                if (ocnConn != null)
                {
                    if (ocnConn.State == ConnectionState.Open)
                    {
                        octTransaction.Rollback();
                        octTransaction.Dispose();
                        octTransaction = null;
                    }
                }
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
        }
        #endregion

        #region Table Operation
        /// <summary>
        /// This function is to insert/alter/drop a table
        /// Please make sure connection is opened before call this fucntion
        /// </summary>
        /// <param name="ocOperation">Oracle command</param>
        /// <returns>Return true if executed</returns>
        public bool TableOperation(OracleCommand ocdOperation, bool bolCloseConnOnErr)
        {
            try
            {
                ocdOperation.Connection = ocnConn;
                ocdOperation.ExecuteNonQuery();

                ocdOperation.Dispose();

                if (octTransaction == null)
                    CloseConnection();
            }
            catch (Exception expQuery)
            {
                if (bolCloseConnOnErr)
                {
                    if (octTransaction != null)
                        RollBack();
                    else
                        CloseConnection();
                }

                throw expQuery;
            }
            return true;
        }
        #endregion
        #region Close Oracle Connection
        /// <summary>
        /// Close Connections
        /// </summary>
        /// <returns></returns>
        private void CloseConnection()
        {
            try
            {

                if (ocnConn != null)
                {
                    if (ocnConn.State == ConnectionState.Open)
                    {
                        if (octTransaction != null)
                        {
                            RollBack();
                            //CCH 070408
                            throw new Exception("A transaction connection cannot be closed using this function");
                            //CCH 070408 end
                        }
                        ocnConn.Close();
                        ocnConn.Dispose();
                    }
                }
            }
            catch (Exception expConnection)
            {
                throw expConnection;
            }
        }
        #endregion
        #region RollBack Database and close the connection
        /// <summary>
        /// RollBack Database and close the connection
        /// </summary>
        /// <returns></returns>
        private void RollBack()
        {
            try
            {
                if (ocnConn != null)
                {
                    if (ocnConn.State == ConnectionState.Open)
                    {
                        octTransaction.Rollback();
                        octTransaction.Dispose();
                        octTransaction = null;
                        ocnConn.Close();
                        ocnConn.Dispose();
                        ocnConn = null;
                    }
                }
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
        }
        #endregion

        #region Check if Transaction and Connection is active
        /// <summary>
        /// Checked if the Transaction is opened or closed with the connection
        /// Pass in parameter true to commit transaction or false to rollback
        /// </summary>
        /// <returns>Return true if the connection and transaction is closed 
        ///          else commit/rollback transaction and close connection
        /// </returns>
        public bool CheckTransaction(bool bolAction)
        {
            try
            {
                if ((octTransaction != null) && (octTransaction.Connection != null && ocnConn.State != ConnectionState.Closed))
                {
                    if (bolAction)
                        CloseConnectionWithTransaction();
                    else
                        RollBack();
                }
                return true;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
        }
        #endregion

        #region Close Oracle Connection With Transaction
        /// <summary>
        /// Close Connections With Transaction
        /// </summary>
        /// <returns></returns>
        private void CloseConnectionWithTransaction()
        {
            try
            {
                if (ocnConn != null)
                {
                    if (ocnConn.State == ConnectionState.Open)
                    {
                        octTransaction.Commit();
                        octTransaction.Dispose();
                        octTransaction = null;
                        ocnConn.Close();
                        ocnConn.Dispose();
                        ocnConn = null;
                    }
                }
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
        }
        #endregion

        #region Fetch Data without closing connection
        /// <summary>
        /// Get from datatable
        /// </summary>
        /// <param name="ocGet">Oracle Command</param>
        /// <returns>Returns Datatable of result</returns>
        public DataTable GetResultDTMaintainConnection(OracleCommand ocdGet)
        {
            try
            {
                DataTable dtbResult = new DataTable();

                OracleDataAdapter odaResult = new OracleDataAdapter(ocdGet);
                ocdGet.Connection = ocnConn;

                odaResult.Fill(dtbResult);

                ocdGet.Dispose();
                return dtbResult;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
        }

        /// <summary>
        /// Get data from Datatable with safe mapping to sync datatype differences between oracle and .NET
        /// </summary>
        /// <param name="ocdGet">orce comamnd object with the query</param>
        /// <returns> datatable</returns>
        public DataTable GetResultDTWithSafeMapping(OracleCommand ocdGet)
        {
            try
            {
                DataTable dtbResult = new DataTable();
                if (ocnConn == null)
                {
                    if (!CreateConnection())
                        return dtbResult;
                }

                OracleDataAdapter odaResult = new OracleDataAdapter(ocdGet);
                ocdGet.Connection = ocnConn;
                odaResult.SafeMapping.Add("*", typeof(string));

                odaResult.Fill(dtbResult);

                ocdGet.Dispose();
                return dtbResult;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
        }
        #endregion

        #region Get Single Query Return Datatable
        /// <summary>
        /// Get from datatable
        /// </summary>
        /// <param name="ocGet">Oracle Command</param>
        /// <returns>Returns Datatable of result</returns>
        public DataTable GetResultDT(OracleCommand ocdGet)
        {
            try
            {
                DataTable dtbResult = new DataTable();
                if (ocnConn == null)
                {
                    if (!CreateConnection())
                        return dtbResult;
                }

                OracleDataAdapter odaResult = new OracleDataAdapter(ocdGet);
                ocdGet.Connection = ocnConn;

                odaResult.Fill(dtbResult);

                ocdGet.Dispose();
                return dtbResult;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
            finally
            {
                CloseConnection();
            }
        }

        public DataTable GetResultDT2(OracleCommand ocdGet)
        {
            try
            {
                DataTable dtbResult = new DataTable();
                if (ocnConn == null)
                {
                    if (!CreateConnection())
                        return dtbResult;
                }
                ocdGet.Connection = ocnConn;
                DateTime dtestart = DateTime.Now;
                OracleDataReader odrResult = ocdGet.ExecuteReader();
                odrResult.FetchSize = 500000 * ocdGet.RowSize;
                //odrResult.Read();
                dtbResult.Load(odrResult);
                DateTime dteend = DateTime.Now;
                TimeSpan ts = dteend - dtestart;
                odrResult.Close();
                ocdGet.Dispose();

                return dtbResult;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
            finally
            {
                CloseConnection();
            }
        }

        public DataTable GetResultDT(OracleCommand ocdGet, bool bolCloseConn)
        {
            try
            {
                DataTable dtbResult = new DataTable();
                if (ocnConn == null)
                {
                    if (!CreateConnection())
                        return dtbResult;
                }
                OracleDataAdapter odaResult = new OracleDataAdapter(ocdGet);
                ocdGet.Connection = ocnConn;
                odaResult.Fill(dtbResult);

                //odaResult.Close();

                ocdGet.Dispose();
                return dtbResult;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
            finally
            {
                if (bolCloseConn)
                    CloseConnection();
            }
        }

        #endregion

        #region Single Operation
        /// <summary>
        /// This function is for single Insert / Delete / Update
        /// Please make sure connection is opened before call this fucntion
        /// </summary>
        /// <param name="ocOperation">Oracle command</param>
        /// <param name="bolCloseConn">True : Close connection after function executes, False : Keeps connection open after function executes</param>
        /// <returns>Return true if executed</returns>
        public bool SingleOperation(OracleCommand ocdOperation, bool bolCloseConn)
        {
            try
            {
                ocdOperation.Connection = ocnConn;
                long intRowOperation = ocdOperation.ExecuteNonQuery();

                if (intRowOperation < 0)
                    throw new Exception("Error occoured in SingleOperation!");

                ocdOperation.Dispose();

                if (octTransaction == null && bolCloseConn)
                    CloseConnection();

                return true;
            }
            catch (Exception expQuery)
            {
                if (octTransaction != null)
                    RollBack();
                else
                    CloseConnection();

                throw expQuery;
            }
        }

        /// <summary>
        /// This function is for single Insert / Delete / Update
        /// Please make sure connection is opened before call this fucntion
        /// </summary>
        /// <param name="ocOperation">Oracle command</param>
        /// <param name="bolCloseConn">True : Close connection after function executes, False : Keeps connection open after function executes</param>
        /// <param name="intResult">Returns the number of rows affected</param>
        /// <returns>Return true if executed</returns>
        public bool SingleOperation(OracleCommand ocdOperation, bool bolCloseConn, out int intResult)
        {
            try
            {
                ocdOperation.Connection = ocnConn;
                intResult = ocdOperation.ExecuteNonQuery();

                if (intResult < 0)
                    throw new Exception("Error occoured in SingleOperation!");

                ocdOperation.Dispose();

                if (octTransaction == null && bolCloseConn)
                    CloseConnection();

                return true;
            }
            catch (Exception expQuery)
            {
                if (octTransaction != null)
                    RollBack();
                else
                    CloseConnection();

                throw expQuery;
            }
        }
        #endregion

        //start BGRF-1795
        #region Get Single Column Query Return Object
        /// <summary>
        /// Get Single Column Query Result return single object
        /// </summary>
        /// <param name="ocGet">Oracle Command</param>
        /// <returns>Return single object</returns>
        public Object GetResultObj(OracleCommand ocdGet, bool bolCloseConn)
        {
            try
            {
                object objResult = new object();
                if (ocnConn == null)
                {
                    if (!CreateConnection())
                        return false;
                }
                ocdGet.Connection = ocnConn;
                objResult = ocdGet.ExecuteScalar();
                ocdGet.Dispose();

                return objResult;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
            finally
            {
                if (bolCloseConn)
                    CloseConnection();
            }
        }

        #region Get Single Column Query Return Object

        public Object GetResultObj_BySQL(string strSQL, bool bolCloseConn)
        {
            try
            {
                object objResult = new object();
                if (ocnConn == null)
                {
                    if (!CreateConnection())
                        return false;
                }

                OracleCommand ocdGet = new OracleCommand(strSQL);
                ocdGet.CommandType = System.Data.CommandType.Text;

                ocdGet.Connection = ocnConn;
                objResult = ocdGet.ExecuteScalar();
                ocdGet.Dispose();

                return objResult;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
            finally
            {
                if (bolCloseConn)
                    CloseConnection();
            }
        }
        #endregion
        #endregion

        #region Table Operation
        /// <summary>
        /// This function is to insert/alter/drop a table
        /// Please make sure connection is opened before call this fucntion
        /// </summary>
        /// <param name="ocOperation">Oracle command</param>
        /// <returns>Return true if executed</returns>

        public bool TableOperation_BySQL(string strSQL, bool bolCloseConn)
        {
            try
            {
                OracleCommand ocdOperation = new OracleCommand(strSQL);
                ocdOperation.CommandType = CommandType.Text;

                ocdOperation.Connection = ocnConn;
                ocdOperation.ExecuteNonQuery();

                ocdOperation.Dispose();

                if (octTransaction == null && bolCloseConn)
                    CloseConnection();
            }
            catch (Exception expQuery)
            {
                if (octTransaction != null)
                    RollBack();
                else
                    CloseConnection();

                throw expQuery;
            }
            return true;
        }

        public long TableOperation_BySQL_ReturnRow(string strSQL, bool bolCloseConn)
        {
            long lRet = 0;
            try
            {
                OracleCommand ocdOperation = new OracleCommand(strSQL);
                ocdOperation.CommandType = CommandType.Text;

                ocdOperation.Connection = ocnConn;
                lRet = ocdOperation.ExecuteNonQuery();

                ocdOperation.Dispose();

                if (octTransaction == null && bolCloseConn)
                    CloseConnection();
            }
            catch (Exception expQuery)
            {
                lRet = -1;

                if (octTransaction != null)
                    RollBack();
                else
                    CloseConnection();

                throw expQuery;
            }

            return lRet;
        }
        #endregion



        #region Get Single Query Return ArrayList
        /// <summary>
        /// Get from datatable
        /// </summary>
        /// <param name="ocGet">Oracle Command</param>
        /// <returns>Returns Arraylist of result</returns>
        public ArrayList GetResultArrayList(OracleCommand ocdGet)
        {
            try
            {
                ArrayList alsResult = new ArrayList();
                if (ocnConn == null)
                {
                    if (!CreateConnection())
                        return alsResult;
                }
                ocdGet.Connection = ocnConn;
                OracleDataReader odrResult = ocdGet.ExecuteReader();

                while (odrResult.Read())
                {
                    object objResult = odrResult.GetValue(0);
                    if (objResult != null)
                    {
                        alsResult.Add(objResult.ToString());
                    }
                }
                ocdGet.Dispose();
                return alsResult;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
            finally
            {
                CloseConnection();
            }
        }

        /// <summary>
        /// Get from datatable
        /// </summary>
        /// <param name="ocGet">Oracle Command</param>
        /// <returns>Returns Arraylist of result</returns>
        public ArrayList GetResultArrayList(OracleCommand ocdGet, bool bolCloseConn)
        {
            try
            {
                ArrayList alsResult = new ArrayList();
                if (ocnConn == null)
                {
                    if (!CreateConnection())
                        return alsResult;
                }
                ocdGet.Connection = ocnConn;
                OracleDataReader odrResult = ocdGet.ExecuteReader();

                while (odrResult.Read())
                {
                    object objResult = odrResult.GetValue(0);
                    if (objResult != null)
                    {
                        alsResult.Add(objResult.ToString());
                    }
                }
                ocdGet.Dispose();
                return alsResult;
            }
            catch (Exception expQuery)
            {
                throw expQuery;
            }
            finally
            {
                if (bolCloseConn)
                    CloseConnection();
            }
        }
        #endregion
        //end BGRF-1795
    }
}
