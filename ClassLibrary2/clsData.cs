using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft;
using System.Collections;
using System.Text.RegularExpressions;
using System.Reflection;

namespace prjData
{
    public class clsData
    {
        static clsDataBase _MRDB;
    
        public static clsDataBase MRData
        {
            get
            {
                if (_MRDB == null)
                {
                    _MRDB = new clsDataBase();
                    _MRDB.connString = "Data Source=SQL02;Initial Catalog=BI_Methodology;Persist Security Info=True;User ID=mrcharts;Password='pwdmrchae0_d';Connect Timeout=0";       }


                return _MRDB;
            }
            set
            {
                _MRDB = value;
            }
        }



    }


    public class clsDataBase
    {
        string x_connstring = "";

        public string connString
        {
            get
            {
                return x_connstring;
            }
            set
            {
                x_connstring = value;
            }
        }

        public  string HtmlToPlainText(string html)
        {
            const string tagWhiteSpace = @"(>|$)(\W|\n|\r)+<";
            const string stripFormatting = "<[^>]*(>|$)";
            const string lineBreak = @"<(br|BR)\s{0,1}\/{0,1}>";
            var lineBreakRegex = new Regex(lineBreak, RegexOptions.Multiline);
            var stripFormattingRegex = new Regex(stripFormatting, RegexOptions.Multiline);
            var tagWhiteSpaceRegex = new Regex(tagWhiteSpace, RegexOptions.Multiline);
            var text = html;
            // text = WebUtility.System.Net.utiHtmlDecode(text)
            text = tagWhiteSpaceRegex.Replace(text, "><");


            text = lineBreakRegex.Replace(text, Environment.NewLine);
            text = stripFormattingRegex.Replace(text, string.Empty);
            return text;
        }

        public bool justExecute(string mySql)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection(connString))
                {
                    SqlCommand myCommand = new SqlCommand(mySql, myConnection);

                    myCommand.CommandTimeout = 0;
                    myConnection.Open();

                    if (!String.IsNullOrEmpty(mySql))
                    {
                      int i=  myCommand.ExecuteNonQuery();
                    }

                    return true;
                }

            }
            catch (Exception e)
            {
                return false;
            }
        }

        public int  justExecute2(string mySql)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection(connString))
                {
                    SqlCommand myCommand = new SqlCommand(mySql, myConnection);

                    myCommand.CommandTimeout = 0;
                    myConnection.Open();

                    if (!String.IsNullOrEmpty(mySql))
                    {
                        return myCommand.ExecuteNonQuery();
                    }

                    else
                    {

                        return -1;
                    }

                     
                }

            }
            catch (Exception e)
            {
                return -1;
            }
        }

        public string getStrValue2(string mySql)
        {

            string toret = "";
            using (SqlConnection myConnection = new SqlConnection(connString))
            {
                SqlCommand myCommand = new SqlCommand(mySql, myConnection);
                myCommand.CommandTimeout = 0;
                myConnection.Open();

                using (SqlDataReader myReader = myCommand.ExecuteReader())
                {
                    if (myReader.HasRows)
                    {
                        if (myReader.Read())
                        {
                            if (!myReader.IsDBNull(0))
                            {
                                toret = myReader.GetString(0);
                            }
                        }
                    }
                }
            }

            return toret;


        }

        public async Task<string> getStrValue(string mySql)
        {

            string toret = "";
            using (SqlConnection myConnection = new SqlConnection(connString))
            {
                SqlCommand myCommand = new SqlCommand(mySql, myConnection);
                myCommand.CommandTimeout = 0;
                myConnection.Open();

                using (SqlDataReader myReader = myCommand.ExecuteReader())
                {
                    if (myReader.HasRows)
                    {
                        if (myReader.Read())
                        {
                            if (!myReader.IsDBNull(0))
                            {
                                toret = myReader.GetString(0);
                            }
                        }
                    }
                }
            }

            return toret;


        }

        public int getIntValue(string mySql)
        {
            int toret = 0;
            using (SqlConnection myConnection = new SqlConnection(connString))
            {
                SqlCommand myCommand = new SqlCommand(mySql, myConnection);
                myCommand.CommandTimeout = 0;
                myConnection.Open();

                using (SqlDataReader myReader = myCommand.ExecuteReader())
                {
                    if (myReader.HasRows)
                    {
                        if (myReader.Read())
                        {
                            if (!myReader.IsDBNull(0))
                            {
                                toret = myReader.GetInt32(0);
                            }
                        }
                    }
                }
            }

            return toret;

        }

        public double getDblValue(string mySql)
        {
            double toret = 0;
            using (SqlConnection myConnection = new SqlConnection(connString))
            {
                SqlCommand myCommand = new SqlCommand(mySql, myConnection);
                myCommand.CommandTimeout = 0;
                myConnection.Open();

                using (SqlDataReader myReader = myCommand.ExecuteReader())
                {
                    if (myReader.HasRows)
                    {
                        if (myReader.Read())
                        {
                            if (!myReader.IsDBNull(0))
                            {
                                toret = myReader.GetDouble(0);
                            }
                        }
                    }
                }
            }

            return toret;
        }

        public bool getBoolValue(string mySql)
        {
            bool toret = false;
            using (SqlConnection myConnection = new SqlConnection(connString))
            {
                SqlCommand myCommand = new SqlCommand(mySql, myConnection);
                myCommand.CommandTimeout = 0;
                myConnection.Open();

                using (SqlDataReader myReader = myCommand.ExecuteReader())
                {
                    if (myReader.HasRows)
                    {
                        if (myReader.Read())
                        {
                            if (!myReader.IsDBNull(0))
                            {
                                toret = myReader.GetBoolean(0);
                            }
                        }
                    }
                }
            }


            return toret;
        }

        public bool chkExisitingemail(string stremail, string strCompany, string strContactDB)
        {
            if (stremail != "")
            {
                string sqlstring = "[BIContacts].[dbo].[chkExisitingemail] '" + stremail + "';";

                // Dim sqlstring As String = "select count(*) from nb_contacts where email= " + " '" + stremail + "' " + " and companyid =(Select id from companies where company= " + "'" + Me.txtClient.Text + "'" + ")"

                int cid = prjData.clsData.MRData.getCount(sqlstring);


                if (cid > 0)
                    return true;
                else
                    return false;
            }
            else
                return false;
        }






        public string fnreplacesingleQuote(string strinput)
        {
            if(!string.IsNullOrEmpty(strinput))
            {
                strinput = strinput.Replace("'", "''");
            }
         
            return strinput;
        }

        public string fnRemoveNonAlphaNumerical(string strInput)
        {
            string str = strInput;

            str = Regex.Replace(str, "[^a-zA-Z0-9]", String.Empty);
          return str;
        }

        public List<string> getArrayList(string mySql)
        {

            List<string> toret = new List<string>();
            using (SqlConnection myConnection = new SqlConnection(connString))
            {
                SqlCommand myCommand = new SqlCommand(mySql, myConnection);
                myCommand.CommandTimeout = 0;
                myConnection.Open();

                using (SqlDataReader myReader = myCommand.ExecuteReader())
                {
                    if (myReader.HasRows)
                    {
                        while (myReader.Read())
                        {
                            toret.Add(myReader.GetString(0));
                        }
                    }
                }
            }

            return toret;
        }

        public SqlDataReader getReader(string mySql)
        {

            using (SqlConnection myConnection = new SqlConnection(connString))
            {
                SqlCommand myCommand = new SqlCommand(mySql, myConnection);
                myCommand.CommandTimeout = 0;
                myConnection.Open();

                using (SqlDataReader myReader = myCommand.ExecuteReader())
                {
                    if (myReader.HasRows)
                    {
                        return myReader;
                    }
                }
            }

            return null;
        }

        public DataTable getDataTable(string mySql)
        {
            DataTable toret = new DataTable { TableName = "Results" };


            using (SqlConnection myConnection = new SqlConnection(connString))
            {
                SqlCommand myCommand = new SqlCommand(mySql, myConnection);
                myCommand.CommandTimeout = 0;
                myConnection.Open();

                SqlDataAdapter myDataAdapter = new SqlDataAdapter(myCommand);
                myDataAdapter.Fill(toret);

            }

            return toret;
        }


        public string getJSON(string mySql)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection(connString))
                {
                    SqlCommand myCommand = new SqlCommand(mySql, myConnection);

                    myCommand.CommandTimeout = 0;
                    myConnection.Open();

                    SqlDataReader myReader = myCommand.ExecuteReader();
                    return WriteDataReader(myReader);
                }

            }
            catch (Exception e)
            {
                return "";
            }
        }


        public string getJSONWithGroupby(string mySql)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection(connString))
                {
                    SqlCommand myCommand = new SqlCommand(mySql, myConnection);

                    myCommand.CommandTimeout = 0;
                    myConnection.Open();

                    SqlDataReader myReader = myCommand.ExecuteReader();
                    return WriteDataReaderWithGroupby(myReader);
                }

            }
            catch (Exception e)
            {
                return "";
            }
        }




        private IEnumerable Groupby(DataTable dt)
        {
            var dttemp = dt.AsEnumerable().GroupBy(x => new { Company = x["Company"], Contact = x["Contact"], Title = x["Title"], Email = x["Email"], ContactDB = x["ContactDB"], NewContactID = x["NewContactID"], DateClicked = x["DateClicked"], EmailOnly=x["EmailOnly"],Color=x["Color"], EmailMessage=x["EmailMessage"] }).Select(y => new
            {
                y.Key.Title,
                y.Key.Company,
                y.Key.ContactDB,
                y.Key.Email,
                y.Key.Contact,
                y.Key.NewContactID,
                y.Key.DateClicked,
                y.Key.EmailOnly,
                y.Key.Color,
                y.Key.EmailMessage,
                ClickType = string.Join(", ", y.Select(z => z.Field<string>("ClickType")))
            });


            return dttemp;

        }


        public List<T> ConvertDataTableToListGeneric<T>(DataTable dataTable) where T : new()
        {
            var list = new List<T>();

            // Iterate through the rows in the DataTable
            foreach (DataRow row in dataTable.Rows)
            {
                T obj = new T();

                // Iterate through the columns in the DataTable
                foreach (DataColumn column in dataTable.Columns)
                {
                    // Get the property in T with the same name as the DataTable column
                    PropertyInfo property = typeof(T).GetProperty(column.ColumnName);

                    // If the property exists in T, and it's not null or empty, set the value
                    if (property != null && row[column] != DBNull.Value)
                    {
                        // Convert the value from the DataTable to the property type and set it
                        property.SetValue(obj, Convert.ChangeType(row[column], property.PropertyType));
                    }
                }

                // Add the object to the list
                list.Add(obj);
            }

            return list;
        }



        public string CompletePipelineTotals(string sql)
        {

            DataTable dt = getDataTable(sql);

            int intp1 = 0;
            int intp2 = 0;
            int intp3 = 0;
            int inttot = 0;


            for (var i = 0; i <= dt.Rows.Count - 1; i++)
            {
                intp1 = intp1 + Convert.ToInt16(dt.Rows[i]["Phase1"]);
                intp2 = intp2 + Convert.ToInt16(dt.Rows[i]["Phase2"]);
                intp3 = intp3 + Convert.ToInt16(dt.Rows[i]["Phase3"]);
                inttot = inttot + Convert.ToInt16(dt.Rows[i]["Total"]);
            }

            DataRow drow = dt.NewRow();

            drow["Company"] = "Total";
            drow["Phase1"] = intp1;
            drow["Phase2"] = intp2;
            drow["Phase3"] = intp3;
            drow["Total"] = inttot;

            dt.Rows.Add(drow);
            dt.AcceptChanges();

            return DataTableToJSON(dt);


        }

        public int getCount(string sql)
        {
            string mySelectQuery = sql;
            int toret=0;
            using (SqlConnection myConnection = new SqlConnection(connString))
            {
                SqlCommand myCommand = new SqlCommand(mySelectQuery, myConnection);
                SqlDataReader myReader;

                myConnection.Open();
                try
                {
                    myCommand.CommandTimeout = 180;
                    myReader = myCommand.ExecuteReader();
                    {
                        var withBlock = myReader;
                        if (withBlock.Read())
                            toret = myReader.GetInt32(0);
                        withBlock.Close();
                    }
                }
                catch (Exception e)
                {
                    //string n = sql + Constants.vbCrLf + tmpConn + Constants.vbCrLf + e.Message;
                    //DataLayer.sendError(n);
                }
                myConnection.Close();
            }
            return toret;
        }

        public string getJSONobj(string mySql)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection(connString))
                {
                    SqlCommand myCommand = new SqlCommand(mySql, myConnection);

                    myCommand.CommandTimeout = 0;
                    myConnection.Open();

                    SqlDataReader myReader = myCommand.ExecuteReader();
                    return WriteDataReaderObj(myReader);
                }

            }
            catch (Exception e)
            {
                return "";
            }
        }

        public string getJSONobjForEmailTemplate(string mySql)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection(connString))
                {
                    SqlCommand myCommand = new SqlCommand(mySql, myConnection);

                    myCommand.CommandTimeout = 0;
                    myConnection.Open();

                    SqlDataReader myReader = myCommand.ExecuteReader();
                    return WriteDataReaderObjForEmailTemplate2(myReader);
                }

            }
            catch (Exception e)
            {
                return "";
            }
        }


        public string DataTableToJSON(DataTable table)
        {
            string JSONString = string.Empty;
            JSONString = JsonConvert.SerializeObject(table);
            return JSONString;
        }


        private string WriteDataReaderObj(IDataReader reader)
        {
            StringBuilder sb = new StringBuilder();
            if (reader == null || reader.FieldCount == 0)
            {
                return "";
            }

            int rowCount = 0;

            while (reader.Read())
            {
                if (rowCount > 0) sb.Append(",");
                sb.Append("{");
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    if (i > 0) sb.Append(",");
                    // sb.Append("\"" + reader.GetName(i) + "\":");
                    sb.Append("" +"\""+ reader.GetName(i) + "\"" + ":");
                    sb.Append("" + JsonConvert.ToString(reader.GetValue(i).ToString()) + "");
                }
                sb.Append("}");
                rowCount++;
            }
           // sb = sb.Replace("\\","");

            string stemp = sb.ToString();
            // stemp = Regex.Replace(sb.ToString(), @"\r\n?|\n", "");

           // stemp = stemp.Trim().Replace("\r\n", string.Empty);
           // stemp = stemp.Trim().Replace("\n", string.Empty);
            stemp = stemp.Replace(Environment.NewLine, string.Empty);

            stemp = stemp.Replace(@"\r\n", "");
           // stemp = stemp.Replace("rn", "");
            sb.Clear();
            sb.Append(stemp);
            return sb.ToString();

        }

        private string WriteDataReaderObjForEmailTemplate(IDataReader reader)
        {
            StringBuilder sb = new StringBuilder();
            if (reader == null || reader.FieldCount == 0)
            {
                return "";
            }

            int rowCount = 0;

            while (reader.Read())
            {
                if (rowCount > 0) sb.Append(",");
                //sb.Append("[");
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    if (i > 0) sb.Append(",");
                    // sb.Append("\"" + reader.GetName(i) + "\":");
                    sb.Append("" + "\"" + reader.GetName(i) + "\"" + ":");
                    // sb.Append("" + JsonConvert.ToString(reader.GetValue(i)) + "");
                    
                    //this is works but cannot be parsed in Json
                    sb.Append("" + reader.GetValue(i).ToString() + "");


                }
              //  sb.Append("]");
                rowCount++;
            }
            sb = sb.Replace("\\", "");
            return sb.ToString();

        }

        private string WriteDataReaderObjForEmailTemplate2(IDataReader reader)
        {
            StringBuilder sb = new StringBuilder();
            if (reader == null || reader.FieldCount == 0)
            {
                return "";
            }

            int rowCount = 0;

            while (reader.Read())
            {
                if (rowCount > 0) sb.Append(",");
                //sb.Append("[");
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    if (i > 0) sb.Append(",");
                    // sb.Append("\"" + reader.GetName(i) + "\":");
                    sb.Append("" + "\"" + reader.GetName(i) + "\"" + ":");
                    // sb.Append("" + JsonConvert.ToString(reader.GetValue(i)) + "");

                    ////this is works but cannot be parsed in Json
                    //if(reader.GetName(i) == "EBody")
                    //{
                    //    sb.Append("" + reader.GetValue(i).ToString() + "");

                    //}

                    //else if(reader.GetName(i) == "ESubject")
                    //{
                    //    sb.Append("" + reader.GetValue(i).ToString() + "");
                    //}
                    //else
                    sb.Append("" + "\"" + reader.GetValue(i).ToString()+ "\"" + "");


                }
                //  sb.Append("]");
                rowCount++;
            }
            sb = sb.Replace("\\", "");
            return sb.ToString();

        }
        private string WriteDataReader(SqlDataReader myReader)
        {
            try
            {
                var dataTable = new DataTable();
                dataTable.Load(myReader);
                string JSONString = string.Empty;
                JSONString = JsonConvert.SerializeObject(dataTable);
                return JSONString;
            }
            catch (Exception e)
            {
                return "";
            }

        }


        private string WriteDataReaderWithGroupby(SqlDataReader myReader)
        {
            try
            {
                var dt = new DataTable();
                dt.Load(myReader);

                var dataTable = Groupby(dt);

                string JSONString = string.Empty;
                JSONString = JsonConvert.SerializeObject(dataTable);
                return JSONString;
            }
            catch (Exception e)
            {
                return "";
            }

        }
    }
}
