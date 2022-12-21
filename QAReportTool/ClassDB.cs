using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QAReportTool
{
    public class ClassDB
    {
        public DataTable SelectQueryNoLock(string query, string conn) //without transaction
        {
            DataTable dt_result = new DataTable();
            SqlConnection _conn = new SqlConnection(conn);
            try
            {
                _conn.Open();
                SqlCommand cmd = new SqlCommand(query, _conn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                da.Fill(dt_result);
            }
            catch (Exception ex)
            {

                //throw;
            }
            finally
            {
                _conn.Close();
            }


            return dt_result;

        }
        public DataSet SelectQueryNoLocks(string query, string conn) //without transaction
        {
            DataSet ds_result = new DataSet();
            SqlConnection _conn = new SqlConnection(conn);
            try
            {
                _conn.Open();
                SqlCommand cmd = new SqlCommand(query, _conn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                da.Fill(ds_result);
            }
            catch (Exception ex)
            {

                //throw;
            }
            finally
            {
                _conn.Close();
            }


            return ds_result;

        }

        public bool ExecQueryNoLock(string query, string conn) //without transaction
        {
            bool result = true;
            query = " BEGIN TRY BEGIN TRAN " + query + @"  COMMIT END TRY BEGIN CATCH  
            
                    DECLARE @ErrorMessage nvarchar(max), @ErrorSeverity int, @ErrorState int
                    SELECT @ErrorMessage = N'Error Number: ' + CONVERT(nvarchar(5), ERROR_NUMBER()) + N'. ' + ERROR_MESSAGE() + ' Line ' + CONVERT(nvarchar(5), ERROR_LINE()), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
                    ROLLBACK TRANSACTION
                    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
            
            END CATCH  ";
            SqlConnection _conn = new SqlConnection(conn);
            try
            {
                _conn.Open();
                SqlCommand cmd = new SqlCommand(query, _conn);
                int hh = cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                result = false;

                //throw;
            }
            finally
            {
                _conn.Close();
            }


            return result;

        }

        public DataTable ExecStoreProcNoLock(string query, string conn) //without transaction
        {
            DataTable dt_result = new DataTable();
            SqlConnection _conn = new SqlConnection(conn);
            try
            {
                _conn.Open();
                SqlCommand cmd = new SqlCommand(query, _conn);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt_result);
            }
            catch (Exception)
            {


                //throw;
            }
            finally
            {
                _conn.Close();
            }


            return dt_result;

        }



    }
}
