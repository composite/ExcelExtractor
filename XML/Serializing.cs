using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Xml.Serialization;
using System.Reflection;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelExtractor.XML
{
    /// <summary>
    /// Excel Serializing Execution
    /// </summary>
    class Serializing:IDisposable
    {
        /// <summary>
        /// Each connection for internal use
        /// </summary>
        private class ConnectClass : IDisposable
        {
            public ConnectClass(DbProviderFactory factory, string connstr)
            {
                Connection = factory.CreateConnection();
                Connection.ConnectionString = connstr;
                SheetConnection = factory.CreateConnection();
                SheetConnection.ConnectionString = connstr;
                RowConnection = factory.CreateConnection();
                RowConnection.ConnectionString = connstr;
                CellConnection = factory.CreateConnection();
                CellConnection.ConnectionString = connstr;
                
            }

            public ConnectClass Open()
            {
                if (IsOpen) throw new InvalidOperationException("All Connection is already opened.");

                Connection.Open();
                SheetConnection.Open();
                RowConnection.Open();
                CellConnection.Open();

                return this;
            }

            public ConnectClass Close()
            {
                if (!IsOpen) throw new InvalidOperationException("All Connection is already closed.");

                CellConnection.Close();
                RowConnection.Close();
                SheetConnection.Close();
                Connection.Close();
                
                return this;
            }

            public bool IsOpen
            {
                get
                {
                    return CellConnection.State.HasFlag(ConnectionState.Open) &&
                           RowConnection.State.HasFlag(ConnectionState.Open) &&
                           SheetConnection.State.HasFlag(ConnectionState.Open) &&
                           Connection.State.HasFlag(ConnectionState.Open);
                }
            }

            public DbConnection Connection { get; private set; }
            public DbConnection SheetConnection { get; private set; }
            public DbConnection RowConnection { get; private set; }
            public DbConnection CellConnection { get; private set; }
            public int CommandTimeout { get; set; }

            public void Dispose()
            {
                if (CellConnection != null && CellConnection.State.HasFlag(ConnectionState.Open)) CellConnection.Close();
                if (RowConnection != null && RowConnection.State.HasFlag(ConnectionState.Open)) RowConnection.Close();
                if (SheetConnection != null && SheetConnection.State.HasFlag(ConnectionState.Open)) SheetConnection.Close();
                if (Connection != null && Connection.State.HasFlag(ConnectionState.Open)) Connection.Close();
            }
        }

        /// <summary>
        /// XML Workbook
        /// </summary>
        private readonly ExcelWorkbook book = null;
        /// <summary>
        /// Each DB Connection
        /// </summary>

        private readonly ConnectClass conn = null;
        /// <summary>
        /// DB Connection String
        /// </summary>
        private readonly string connstr = null;
        /// <summary>
        /// Overwrite existing file Y/N
        /// </summary>
        private readonly bool isOverwrite = false;

        /// <summary>
        /// sheet idx for sheet name duplication.
        /// </summary>
        private int sheetidx = 1;

        private static readonly Regex cmdsplitter = new Regex("\\s+", RegexOptions.Compiled);

        /// <summary>
        /// Serialize constructor from a file
        /// </summary>
        /// <param name="file">save file as...</param>
        public Serializing(string file)
        {

            Console.WriteLine("Serialize activating...");

            if (string.IsNullOrEmpty(file)) throw new ArgumentNullException("file");

            using (var stream = File.Open(file, FileMode.Open))
            {
                var serial = new XmlSerializer(typeof(ExcelWorkbook));
                book = (ExcelWorkbook) serial.Deserialize(stream);
            }

            if (book == null) throw new NullReferenceException("template must not be null. may be failed to making excel template.");

            if (book.ConnectionString != null)
            {
                this.conn = new ConnectClass(DbProviderFactories.GetFactory(book.ConnectionString.Type), book.ConnectionString.Text);
                if (book.ConnectionString.Timeout < 0) conn.CommandTimeout = 0;
                else if (book.ConnectionString.Timeout < 0) conn.CommandTimeout = book.ConnectionString.Timeout;
            }

            Console.WriteLine("Serialize initialzed.");

        }
        /// <summary>
        /// Serialize constructor from a file
        /// </summary>
        /// <param name="file">save file as...</param>
        /// <param name="overwrite">overwrite exising file?</param>
        public Serializing(string file, bool overwrite) : this(file)
        {
            this.isOverwrite = overwrite;
        }

        public static bool IsHex(IEnumerable<char> chars)
        {
            return chars.Select(c => ((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))).All(isHex => isHex);
        }

        public static Color TranslateColor(string cs)
        {
            Color co = Color.White;
            if (cs != null)
            {
                KnownColor kco;
                if (cs.Length > 3 && cs[0] == '#' && IsHex(cs.Substring(1)))
                    co = ColorTranslator.FromHtml(cs);
                else if (Enum.TryParse(cs, out kco)) co = Color.FromKnownColor(kco);
            }
            return co;
        }

        private static void ApplyStyle(ExcelWorksheet sheet, ExcelStyle style)
        {
            Console.WriteLine("Apply style range : " + style.Range);
            try
            {
                var range = sheet.Cells[style.Range];
                if (style.Border != null)
                {
                    var prop = range.Style.Border;
                    Action<ExcelBorderBase, ExcelBorderItem> borfunc = (bord, item) =>
                    {
                        Color co = TranslateColor(bord.Color);

                        ExcelBorderStyle st = ExcelBorderStyle.None;

                        if (bord.Style != null) Enum.TryParse(bord.Style, out st);

                        if (st != ExcelBorderStyle.None)
                        {
                            if (item != null)
                            {
                                item.Style = st;
                                item.Color.SetColor(co);
                            }
                            else prop.BorderAround(st, co);
                        }
                        
                    };

                    var border = style.Border;
                    if (border.All != null)
                    {
                        borfunc.Invoke(border.All, null);
                    }
                    else
                    {
                        if (border.Top != null) borfunc.Invoke(border.Top, prop.Top);
                        if (border.Left != null) borfunc.Invoke(border.Left, prop.Left);
                        if (border.Right != null) borfunc.Invoke(border.Right, prop.Right);
                        if (border.Bottom != null) borfunc.Invoke(border.Bottom, prop.Bottom);
                    }
                }
                if (style.Font != null)
                {
                    var prop = range.Style.Font;
                    var font = style.Font;
                    prop.Color.SetColor(TranslateColor(font.Color));
                    prop.Name = font.Family;
                    prop.Size = font.Size;
                    prop.Bold = font.Bold;
                    prop.Italic = font.Italic;
                    prop.UnderLine = font.Underline;
                    prop.Strike = font.Strike;
                }
                if (style.Back != null)
                {
                    var ptt = ExcelFillStyle.None;
                    if (Enum.TryParse(style.Back.Pattern, out ptt) && ptt != ExcelFillStyle.None)
                    {
                        range.Style.Fill.PatternType = ptt;
                        range.Style.Fill.BackgroundColor.SetColor(TranslateColor(style.Back.Color));
                    }

                }

                /*if (style.Format != null)
                {
                    switch (style.Format.Out)
                    {
                        case ExcelOutType.DateTime:
                            //...?
                            break;
                        case ExcelOutType.Integer:
                            range.Style.Numberformat.Format = "#,##0";
                            break;
                        case ExcelOutType.Number:
                            range.Style.Numberformat.Format = "#,##0.00";
                            break;
                        case ExcelOutType.Money:
                            range.Style.Numberformat.Format = "￦ #,##0";
                            break;
                        case ExcelOutType.Normal:
                        default:
                            range.Style.Numberformat.Format = null;
                            break;
                    }
                }*/

            }
            catch (Exception E)
            {
                Console.Error.WriteLine(E.ToString());
            }
        }

        private static ExcelOutType DetermineOutType(Type type)
        {
            if (type == null) return ExcelOutType.Text;

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.DateTime:
                    return ExcelOutType.DateTime;
                
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    return ExcelOutType.Integer;

                case TypeCode.Double:
                case TypeCode.Single:
                case TypeCode.Decimal:
                    return ExcelOutType.Number;

                case TypeCode.Char:
                case TypeCode.String:
                    return ExcelOutType.Text;
                default:
                    return ExcelOutType.Normal;

            }
        }

        public Serializing Do(out string path)
        {
            Console.WriteLine("Serialize processing...");

            var exfile = book.File;
            string filepath;// = Path.Combine(exfile.Path ?? string.Empty, exfile.Name);
            object[] args = new object[0];
            object[] evargs = new object[0];

            conn.Open();

            if (exfile.SQL != ExcelSQLType.PlainText)
            {
                using (DbCommand cmd = conn.Connection.CreateCommand())
                {

                    cmd.CommandTimeout = conn.CommandTimeout;
                    cmd.CommandType = exfile.SQL == ExcelSQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = exfile.Text;

                    using (DbDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            args = new object[reader.FieldCount];
                            reader.GetValues(args);
                            filepath = Path.Combine(exfile.Path ?? string.Empty, String.Format(exfile.Name, args));
                        }
                        else
                        {
                            filepath = Path.Combine(exfile.Path ?? string.Empty, String.Format(exfile.Name, new object[0]));
                        }
                    }
                }
            }
            else
            {
                filepath = Path.Combine(exfile.Path ?? string.Empty, exfile.Name);
            }

            Console.WriteLine("Generating Excel file : {0}", filepath);

            FileInfo fileInfo = new FileInfo(filepath);
            

            if(fileInfo.Exists && !isOverwrite) throw new FileSkipException();
            if (!fileInfo.Directory.Exists) Directory.CreateDirectory(fileInfo.Directory.FullName);

            evargs = new object[args.Length + 1];
            evargs[0] = filepath;
            if(args.Length > 0) Array.Copy(args, 0, evargs, 1, args.Length);

            try
            {

                if (book.Before != null)
                {
                    var before = book.Before;
                    if (before.SQL != ExcelSQLType.PlainText)
                    {

                        using (DbCommand cmd = conn.Connection.CreateCommand())
                        {

                            cmd.CommandTimeout = conn.CommandTimeout;
                            cmd.CommandType = before.SQL == ExcelSQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                            cmd.CommandText = before.Text;

                            for (int i = 0; i < evargs.Length; i++)
                            {
                                DbParameter param = cmd.CreateParameter();
                                param.DbType = DbType.String;
                                param.ParameterName = cmd.Connection is System.Data.SqlClient.SqlConnection ? "@F" + i : "F" + i;
                                param.Value = evargs[i];
                                cmd.Parameters.Add(param);
                            }

                            cmd.ExecuteNonQuery();
                        }

                    }
                    else if (!string.IsNullOrEmpty(before.CMD))
                    {

                        var process = new Process {EnableRaisingEvents = false};
                        process.StartInfo.FileName = before.CMD;
                        if (before.Text != null) process.StartInfo.Arguments = before.Text;
                        process.Start();

                    }
                }

                if (book.Sheets != null)
                {
                    Console.WriteLine("Generating sheets...");
                    using (var excel = new ExcelPackage(fileInfo))
                    {
                        foreach (var sheet in book.Sheets) DoSheet(sheet, excel, args);
                        excel.Save();
                    }
                }
                else
                {
                    Console.WriteLine("No sheet specified. file will not created.");
                }

                Console.WriteLine("Saving file : {0}", filepath);

                if (book.After != null)
                {
                    var after = book.After;
                    if (after.SQL != ExcelSQLType.PlainText)
                    {

                        using (DbCommand cmd = conn.Connection.CreateCommand())
                        {

                            cmd.CommandTimeout = conn.CommandTimeout;
                            cmd.CommandType = after.SQL == ExcelSQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                            cmd.CommandText = after.Text;

                            for (int i = 0; i < evargs.Length; i++)
                            {
                                DbParameter param = cmd.CreateParameter();
                                param.DbType = DbType.String;
                                param.ParameterName = cmd.Connection is System.Data.SqlClient.SqlConnection ? "@F" + i : "F" + i;
                                param.Value = evargs[i];
                                cmd.Parameters.Add(param);
                            }

                            cmd.ExecuteNonQuery();
                        }

                    }
                    else if (!string.IsNullOrEmpty(after.CMD))
                    {

                        var process = new Process { EnableRaisingEvents = false };
                        process.StartInfo.FileName = after.CMD;
                        if (after.Text != null) process.StartInfo.Arguments = after.Text;
                        process.Start();

                    }
                }

                Console.WriteLine("DONE!!!");

            }
            catch (UnauthorizedAccessException e)
            {
                throw new SerializingException("Cannot create and write file. access denied.", e) {FilePath = filepath};
            }
            catch (IOException e)
            {
                throw new SerializingException("Cannot create and write file. file not exists or locked for other process.", e) {FilePath = filepath};
            }
            catch (Exception e)
            {
                throw new SerializingException(e) {FilePath = filepath};
            }
            finally
            {
                conn.Close();
            }

            path = filepath;
            return this;
        }

        public Serializing Do()
        {
            string dummy;
            return Do(out dummy);
        }

        private void DoSheet(ExcelSheet sheet, ExcelPackage excel, object[] fileargs = null)
        {

            if (sheet.SQL != ExcelSQLType.PlainText)
            {
                using (DbCommand cmd = conn.SheetConnection.CreateCommand())
                {

                    cmd.CommandTimeout = conn.CommandTimeout;
                    cmd.CommandType = sheet.SQL == ExcelSQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = sheet.Text;

                    if (fileargs != null)
                        for (int i = 0; i < fileargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is System.Data.SqlClient.SqlConnection ? "@F" + i : "F" + i;
                            param.Value = fileargs[i];
                            cmd.Parameters.Add(param);
                        }

                    using (DbDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            int currow = 1;
                            object[] args = new object[reader.FieldCount];
                            reader.GetValues(args);

                            var sh = excel.Workbook.Worksheets.Add(reader.FieldCount > 1 ? reader.GetValue(1).ToString() : "SQLSheet" + (sheetidx++));

                            Console.WriteLine("Creating sheet '{0}' with given SQL...", sh.Name);

                            var row = sheet.Rows != null ? sheet.Rows.FirstOrDefault(r => r.ColumnHeader) : null;
                            var idummy = 0;
                            if (row != null) DoRow(ref idummy, row, sheet, sh, fileargs, args);

                            var etcrow = sheet.Rows != null ? sheet.Rows.Where(r => !r.ColumnHeader).ToArray() : new ExcelRow[0];
                            for (int i = 0; i < etcrow.Length; i++) DoRow(ref currow, etcrow[i], sheet, sh, fileargs, args);

                            if (sheet.Style != null) foreach (var style in sheet.Style) ApplyStyle(sh, style);
                            if (sh.Dimension != null) sh.Cells[sh.Dimension.Address].AutoFitColumns();

                        }
                    }
                }

            }
            else
            {
                int currow = 1;
                var sh = excel.Workbook.Worksheets.Add(!string.IsNullOrEmpty(sheet.Name) ? sheet.Name : "OKSheet" + (sheetidx++));
                Console.WriteLine("Creating sheet '{0}'...", sh.Name);
                for (int i = 0; i < (sheet.Rows != null ? sheet.Rows.Count : 0); i++) DoRow(ref currow, sheet.Rows[i], sheet, sh, fileargs);

                if (sheet.Style != null) foreach (var style in sheet.Style) ApplyStyle(sh, style);
                if (sh.Dimension != null) sh.Cells[sh.Dimension.Address].AutoFitColumns();

            }

            Console.WriteLine("Sheet created!");
        }

        private void DoRow(ref int rownum, ExcelRow row, ExcelSheet sheet, ExcelWorksheet exSheet, object[] fileargs = null, object[] sheetargs = null)
        {
            //Process Cell
            
            if (row.SQL != ExcelSQLType.PlainText)
            {
                using (DbCommand cmd = conn.RowConnection.CreateCommand())
                {

                    cmd.CommandTimeout = conn.CommandTimeout;
                    cmd.CommandType = row.SQL == ExcelSQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = row.Text;

                    if (fileargs != null)
                        for (int i = 0; i < fileargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is System.Data.SqlClient.SqlConnection ? "@F" + i : "F" + i;
                            param.Value = fileargs[i];
                            cmd.Parameters.Add(param);
                        }

                    if (sheetargs != null)
                        for (int i = 0; i < sheetargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is System.Data.SqlClient.SqlConnection ? "@S" + i : "S" + i;
                            param.Value = sheetargs[i];
                            cmd.Parameters.Add(param);
                        }

                    using (DbDataReader reader = cmd.ExecuteReader())
                    {
                        
                        if (row.ColumnHeader)
                        {
                            int curcell = 1;
                            using (var rw = exSheet.Cells[rownum++ + ":" + rownum])
                            {
                                var cols = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray();
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    ExcelCell dc = new ExcelCell();
                                    dc.SQL = ExcelSQLType.PlainText;
                                    dc.Text = cols[i];

                                    //Console.WriteLine("Column '{0}' is have a type of '{1}'", cols[i], (reader.GetFieldType(i) ?? typeof(string)).FullName);

                                    DoCell(ref curcell, dc, row, sheet, rw, fileargs, sheetargs, cols);
                                }
                            }
                        }
                        
                        while (reader.Read())
                        {
                            int curcell = 1;
                            using (var rw = exSheet.Cells[rownum++ + ":" + rownum])
                            {
                                //Console.WriteLine("Writing row {0} with given SQL...", rw.Start.Row);
                                object[] args = new object[reader.FieldCount];
                                reader.GetValues(args);
                                for (int i = 0; i < reader.FieldCount; i++)
                                {

                                    if (row.Fetch != null)
                                    {
                                        using (DbCommand subcmd = conn.Connection.CreateCommand())
                                        {

                                            cmd.CommandTimeout = conn.CommandTimeout;
                                            subcmd.CommandType = row.Fetch.SQL == ExcelSQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                                            subcmd.CommandText = row.Fetch.Text;

                                            if (fileargs != null)
                                                for (int j = 0; j < fileargs.Length; j++)
                                                {
                                                    DbParameter param = subcmd.CreateParameter();
                                                    param.DbType = DbType.String;
                                                    param.ParameterName = subcmd.Connection is System.Data.SqlClient.SqlConnection ? "@F" + j : "F" + j;
                                                    param.Value = fileargs[j];
                                                    subcmd.Parameters.Add(param);
                                                }

                                            if (sheetargs != null)
                                                for (int j = 0; j < sheetargs.Length; j++)
                                                {
                                                    DbParameter param = subcmd.CreateParameter();
                                                    param.DbType = DbType.String;
                                                    param.ParameterName = subcmd.Connection is System.Data.SqlClient.SqlConnection ? "@S" + j : "S" + j;
                                                    param.Value = sheetargs[j];
                                                    subcmd.Parameters.Add(param);
                                                }

                                            for (int j = 0; j < args.Length; j++)
                                            {
                                                DbParameter param = subcmd.CreateParameter();
                                                param.DbType = DbType.String;
                                                param.ParameterName = subcmd.Connection is System.Data.SqlClient.SqlConnection ? "@R" + j : "R" + j;
                                                param.Value = args[j];
                                                subcmd.Parameters.Add(param);
                                            }

                                            using (DbDataReader subreader = subcmd.ExecuteReader())
                                            {
                                                if (row.Fetch.Type == ExcelFetchType.Single)
                                                {
                                                    if (!subreader.Read()) continue;
                                                    for (int j = 0; j < subreader.FieldCount; j++)
                                                    {
                                                        ExcelCell dc = new ExcelCell();
                                                        dc.SQL = ExcelSQLType.PlainText;
                                                        dc.Text = subreader.GetValue(j).ToString();
                                                        dc.Out = DetermineOutType(subreader.GetFieldType(j));

                                                        DoCell(ref curcell, dc, row, sheet, rw, fileargs, sheetargs, args);
                                                    }
                                                }
                                                else
                                                {
                                                    while (subreader.Read())
                                                    {
                                                        ExcelCell dc = new ExcelCell();
                                                        dc.SQL = ExcelSQLType.PlainText;
                                                        dc.Text = subreader.GetValue(0).ToString();
                                                        dc.Out = DetermineOutType(subreader.GetFieldType(0));

                                                        DoCell(ref curcell, dc, row, sheet, rw, fileargs, sheetargs, args);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        ExcelCell dc = new ExcelCell();
                                        dc.SQL = ExcelSQLType.PlainText;
                                        dc.Text = reader.GetValue(i).ToString();
                                        dc.Out = DetermineOutType(reader.GetFieldType(i));

                                        DoCell(ref curcell, dc, row, sheet, rw, fileargs, sheetargs, args);
                                    }

                                }
                            }
                                
                        }
                    }

                }
            }
            else
            {
                int curcell = 1;
                using (var rw = exSheet.Cells[rownum++ + ":" + rownum])
                {
                    //Console.WriteLine("Writing row {0}...", rw.Start.Row);
                    for (int i = 0; i < (row.Cells != null ? row.Cells.Count : 0); i++) DoCell(ref curcell, row.Cells[i], row, sheet, rw, fileargs, sheetargs);
                }
            }


        }

        private void DoCell(ref int cellnum, ExcelCell cell, ExcelRow row, ExcelSheet sheet, ExcelRange exRow, object[] fileargs = null, object[] sheetargs = null, object[] rowargs = null)
        {
            //Cell Value
            int cnt = cellnum;
            Action<string> cellfunc = val =>
            {
                using (var cl = exRow[exRow.Start.Row, cnt++])
                {
                    //Console.WriteLine("Writing cell {0}...", cl.Start.Address);

                    if (val != null)
                    {
                        double nval;
                        long lval;
                        DateTime dval;
                        bool isFomula = !string.IsNullOrEmpty(val) && val[0] == '=';
                        string rawval = isFomula ? val.Substring(1) : val;
                        switch (cell.Out)
                        {
                            case ExcelOutType.DateTime:
                                if (isFomula) cl.Formula = rawval; else if (DateTime.TryParse(rawval, out dval)) cl.Value = dval; else cl.Value = rawval;
                                break;
                            case ExcelOutType.Integer:
                                cl.Style.Numberformat.Format = "#,##0";
                                if (isFomula) cl.Formula = rawval; else if (long.TryParse(rawval, out lval)) cl.Value = lval; else cl.Value = rawval;
                                break;
                            case ExcelOutType.Number:
                                cl.Style.Numberformat.Format = "#,##0.00";
                                if (isFomula) cl.Formula = rawval; else if (double.TryParse(rawval, out nval)) cl.Value = nval; else cl.Value = rawval;
                                break;
                            case ExcelOutType.Money:
                                cl.Style.Numberformat.Format = "￦ #,##0";
                                if (isFomula) cl.Formula = rawval; else if (long.TryParse(rawval, out lval)) cl.Value = lval; else cl.Value = rawval;
                                break;
                            case ExcelOutType.Normal:
                            default:
                                if (isFomula) cl.Formula = rawval; else cl.Value = rawval;
                                break;
                        }
                    }
                }

            };

            if (cell.SQL != ExcelSQLType.PlainText)
            {
                using (DbCommand cmd = conn.CellConnection.CreateCommand())
                {

                    cmd.CommandTimeout = conn.CommandTimeout;
                    cmd.CommandType = cell.SQL == ExcelSQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = cell.Text;

                    if (fileargs != null)
                        for (int i = 0; i < fileargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is System.Data.SqlClient.SqlConnection ? "@F" + i : "F" + i;
                            param.Value = fileargs[i];
                            cmd.Parameters.Add(param);
                        }

                    if (sheetargs != null)
                        for (int i = 0; i < sheetargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is System.Data.SqlClient.SqlConnection ? "@S" + i : "S" + i;
                            param.Value = sheetargs[i];
                            cmd.Parameters.Add(param);
                        }

                    if (rowargs != null)
                        for (int i = 0; i < rowargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is System.Data.SqlClient.SqlConnection ? "@R" + i : "R" + i;
                            param.Value = rowargs[i];
                            cmd.Parameters.Add(param);
                        }

                    using (DbDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            cellfunc.Invoke(reader.GetValue(0).ToString());
                        }
                    }

                }
            }
            else if(cell.Text != null)
            {
                cellfunc.Invoke(cell.Text);
            }

            cellnum = cnt;
        }

        public void Dispose()
        {
            try
            {
                conn.Dispose();
            }
            catch
            {
                // ignored
            }
        }
    }

    public class FileSkipException : IOException
    {
        public FileSkipException() : base("File is already exists. skipped.") { }
    }

    public class SerializingException : Exception
    {
        public SerializingException() : base("An error occured while serializing.") { }
        public SerializingException(Exception e) : base("An error occured while serializing.", e) { }
        public SerializingException(string s) : base(s) { }
        public SerializingException(string s, Exception e) : base(s, e) { }
        public string FilePath { get; set; }
    }
}
