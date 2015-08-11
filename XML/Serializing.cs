using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Xml.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Reflection;

namespace ExcelExtractor.XML
{
    /// <summary>
    /// Excel Serializing Execution
    /// </summary>
    class Serializing
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
        private readonly Workbook book = null;
        /// <summary>
        /// Excel Workbook
        /// </summary>
        private readonly IWorkbook excel = null;
        /// <summary>
        /// Each DB Connection
        /// </summary>

        private ConnectClass conn = null;
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

        /// <summary>
        /// named styles
        /// </summary>
        private Dictionary<string, ICellStyle> styles = new Dictionary<string, ICellStyle>(0);

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
                var serial = new XmlSerializer(typeof(Workbook));
                book = (Workbook) serial.Deserialize(stream);
            }

            if (book == null) throw new NullReferenceException("template must not be null. may be failed to making excel template.");

            if (book.ConnectionString != null)
            {
                this.conn = new ConnectClass(DbProviderFactories.GetFactory(book.ConnectionString.Type), book.ConnectionString.Text);
            }

            excel = book.Xlsx ? (IWorkbook)new XSSFWorkbook() : new HSSFWorkbook();

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

        private short GetColor(string color)
        {
            
            if (color.Length == 6 && IsHex(color))
            {
                byte[] raw = new byte[3];
                for (int i = 0; i < raw.Length; i++)
                    raw[i] = byte.Parse(color.Substring(i*2, (i + 1)*2), NumberStyles.AllowHexSpecifier);

                return new XSSFColor(raw).Indexed;
            }
            else
            {
                Type named = typeof (HSSFColor)
                    .GetNestedTypes(BindingFlags.Instance | BindingFlags.Public)
                    .FirstOrDefault(t => t.Name.Equals(color, StringComparison.InvariantCultureIgnoreCase));

                if (named != null)
                {
                    return ((IColor) Activator.CreateInstance(named)).Indexed;
                }
                else return HSSFColor.COLOR_NORMAL;
            }

            
        }

        private static NPOI.SS.UserModel.BorderStyle GetBorderStyle(string bs)
        {
            foreach (var style in Enum.GetNames(typeof(NPOI.SS.UserModel.BorderStyle)))
            {
                if (style.Equals(bs, StringComparison.InvariantCultureIgnoreCase)) return (NPOI.SS.UserModel.BorderStyle) Enum.Parse(typeof (NPOI.SS.UserModel.BorderStyle), style);
            }
            return NPOI.SS.UserModel.BorderStyle.None;
        }

        private void ApplyCellStyle(ISheet origin, Sheet sheet)
        {
            HGroup heven = null, hodd = null;
            VGroup veven = null, vodd = null;

            if (sheet.HGroup != null)
            {
                heven = sheet.HGroup.FirstOrDefault(hg => "even".Equals(hg.IndexGroup, StringComparison.InvariantCultureIgnoreCase));
                hodd = sheet.HGroup.FirstOrDefault(hg => "odd".Equals(hg.IndexGroup, StringComparison.InvariantCultureIgnoreCase));
            }
            if (sheet.VGroup != null)
            {
                veven = sheet.VGroup.FirstOrDefault(vg => "even".Equals(vg.IndexGroup, StringComparison.InvariantCultureIgnoreCase));
                vodd = sheet.VGroup.FirstOrDefault(vg => "odd".Equals(vg.IndexGroup, StringComparison.InvariantCultureIgnoreCase));
            }

            int maxcol = 0, rc = 0;
            var iterate = origin.GetEnumerator();
            while (iterate.MoveNext())
            {
                IRow rw = (IRow) iterate.Current;
                maxcol = Math.Max(maxcol, rw.Cells.Count);

                if (sheet.VGroup != null)
                {
                    short height = (short) (sheet.HGroup.FirstOrDefault(hg => hg.Index == rc) ?? new HGroup()).Height;
                    if (height > 0) rw.Height = height;
                }

                foreach (var cell in rw.Cells)
                {
                    if (cell.CellStyle == null)
                    {
                        CellGroup cg = null;
                        if (sheet.HGroup != null) cg = sheet.HGroup.FirstOrDefault(hg => hg.Index == cell.ColumnIndex);
                        else if(sheet.VGroup != null) cg = sheet.VGroup.FirstOrDefault(vg => vg.Index == rc);
                        else if (hodd != null || heven != null) cg = cell.ColumnIndex%2 == 0 ? hodd : heven;
                        else if (vodd != null || veven != null) cg = rc%2 == 0 ? vodd : veven;

                        cell.CellStyle = excel.CreateCellStyle();
                        if (cg != null && styles.ContainsKey(cg.Style)) cell.CellStyle.CloneStyleFrom(styles[cg.Style]);
                        switch ((cg.HAlign ?? string.Empty).ToLower())
                        {
                            case "left":
                                cell.CellStyle.Alignment = HorizontalAlignment.Left;
                                break;
                            case "center":
                                cell.CellStyle.Alignment = HorizontalAlignment.Center;
                                break;
                            case "right":
                                cell.CellStyle.Alignment = HorizontalAlignment.Right;
                                break;
                        }
                        switch ((cg.VAlign ?? string.Empty).ToLower())
                        {
                            case "top":
                                cell.CellStyle.VerticalAlignment = VerticalAlignment.Top;
                                break;
                            case "middle":
                                cell.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                                break;
                            case "bottom":
                                cell.CellStyle.VerticalAlignment = VerticalAlignment.Bottom;
                                break;
                        }
                    }
                }

                rc++;
            }
            if(sheet.AutoSizeColumn)
                for(int i=0;i<maxcol;i++) origin.AutoSizeColumn(i);
            else
                for (int i = 0; i < maxcol; i++)
                {
                    int width = (sheet.VGroup.FirstOrDefault(vg => vg.Index == rc) ?? new VGroup()).Width;
                    if (width > 0) origin.SetColumnWidth(i, width);
                }
        }

        public Serializing Do(out string path)
        {
            Console.WriteLine("Serialize processing...");

            var exfile = book.File;
            string filepath;// = Path.Combine(exfile.Path ?? string.Empty, exfile.Name);

            conn.Open();

            if (exfile.Sql != SQLType.PlainText)
            {
                using (DbCommand cmd = conn.Connection.CreateCommand())
                {
                    
                    cmd.CommandType = exfile.Sql == SQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = exfile.Text;

                    using (DbDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            object[] args = new object[reader.FieldCount];
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

            try
            {


                if (book.Styles != null)
                    foreach (var style in book.Styles.Style)
                    {
                        var cellstyle = excel.CreateCellStyle();



                        if (style.Font != null)
                        {
                            var font = excel.CreateFont();
                            font.FontName = style.Font.Face;
                            font.FontHeight = style.Font.Size;
                            font.Color = GetColor(style.Font.Color);
                            font.Boldweight = (short) (style.Font.Bold ? 1 : 0);
                            font.IsItalic = style.Font.Italic;
                            font.IsStrikeout = style.Font.Strike;
                            font.Underline = style.Font.Underline ? FontUnderlineType.Single : FontUnderlineType.None;

                            cellstyle.SetFont(font);
                        }

                        if (style.Border != null || style.Back != null)
                        {

                            if (style.Back != null)
                            {
                                cellstyle.FillBackgroundColor = GetColor(style.Back.Color);
                            }

                            if (style.Border != null)
                                if (style.Border.All != null)
                                {
                                    cellstyle.BorderLeft = cellstyle.BorderRight = cellstyle.BorderTop = cellstyle.BorderBottom = GetBorderStyle(style.Border.All.Width);
                                    cellstyle.LeftBorderColor = cellstyle.RightBorderColor = cellstyle.TopBorderColor = cellstyle.BottomBorderColor = GetColor(style.Border.All.Color);
                                }
                                else
                                {
                                    if (style.Border.Left != null)
                                    {
                                        cellstyle.BorderLeft = GetBorderStyle(style.Border.Left.Width);
                                        cellstyle.LeftBorderColor = GetColor(style.Border.Left.Color);
                                    }
                                    if (style.Border.Right != null)
                                    {
                                        cellstyle.BorderLeft = GetBorderStyle(style.Border.Right.Width);
                                        cellstyle.LeftBorderColor = GetColor(style.Border.Right.Color);
                                    }
                                    if (style.Border.Top != null)
                                    {
                                        cellstyle.BorderLeft = GetBorderStyle(style.Border.Top.Width);
                                        cellstyle.LeftBorderColor = GetColor(style.Border.Top.Color);
                                    }
                                    if (style.Border.Bottom != null)
                                    {
                                        cellstyle.BorderLeft = GetBorderStyle(style.Border.Bottom.Width);
                                        cellstyle.LeftBorderColor = GetColor(style.Border.Bottom.Color);
                                    }
                                }
                        }

                        /*
                        switch ((style.HAlign ?? string.Empty).ToLower())
                        {
                            case "left":
                                cellstyle.Alignment = HorizontalAlignment.Left;
                                break;
                            case "center":
                                cellstyle.Alignment = HorizontalAlignment.Center;
                                break;
                            case "right":
                                cellstyle.Alignment = HorizontalAlignment.Right;
                                break;
                        }
                        switch ((style.VAlign ?? string.Empty).ToLower())
                        {
                            case "top":
                                cellstyle.VerticalAlignment = VerticalAlignment.Top;
                                break;
                            case "middle":
                                cellstyle.VerticalAlignment = VerticalAlignment.Center;
                                break;
                            case "bottom":
                                cellstyle.VerticalAlignment = VerticalAlignment.Bottom;
                                break;
                        }
                        */

                        styles.Add(style.Id, cellstyle);

                    }

                if (book.Sheet == null || book.Sheet.Length == 0) throw new SerializingException("No sheet specified. You must have least 1 <Sheet> element.");

                Console.WriteLine("Generating sheets...");

                foreach (var sheet in book.Sheet) DoSheet(sheet);

                Console.WriteLine("Saving file : {0}", filepath);

                //Wirte file and end.
                if (!fileInfo.Directory.Exists) Directory.CreateDirectory(fileInfo.Directory.FullName);
                using (var file = fileInfo.Create())
                {
                    excel.Write(file);
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

        private void DoSheet(Sheet sheet, object[] fileargs = null)
        {
            
            //Process Row
            if (sheet.Sql != SQLType.PlainText)
            {
                using (DbCommand cmd = conn.SheetConnection.CreateCommand())
                {

                    cmd.CommandType = sheet.Sql == SQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = sheet.Text;

                    if (fileargs != null)
                        for (int i = 0; i < fileargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is SqlConnection ? "@F" + i : "F" + i;
                            param.Value = fileargs[i];
                            cmd.Parameters.Add(param);
                        }

                    using (DbDataReader reader = cmd.ExecuteReader())
                    {
                        int idx = 0;
                        while (reader.Read())
                        {
                            object[] args = new object[reader.FieldCount];
                            reader.GetValues(args);

                            var sh = excel.CreateSheet(reader.FieldCount > 1 ? reader.GetValue(1).ToString() : "SQLSheet" + (sheetidx++));

                            Console.WriteLine("Creating sheet '{0}' with given SQL...", sh.SheetName);

                            var row = sheet.Row != null ? sheet.Row.FirstOrDefault(r => r.ColumnHeader) : null;
                            var idummy = 0;
                            if (row != null) DoRow(ref idummy, row, sheet, sh, fileargs, args);

                            var etcrow = sheet.Row != null ? sheet.Row.Where(r => !r.ColumnHeader).ToArray() : new Row[0];
                            for (int i=0; i < etcrow.Length; i++) DoRow(ref i, etcrow[i], sheet, sh, fileargs, args);

                            ApplyCellStyle(sh, sheet);

                            idx++;
                        }
                    }

                }
            }
            else
            {
                var sh = excel.CreateSheet(!string.IsNullOrEmpty(sheet.Name) ? sheet.Name : "OKSheet" + (sheetidx++));
                Console.WriteLine("Creating sheet '{0}'...", sh.SheetName);
                for (int i = 0; i < (sheet.Row != null ? sheet.Row.Length : 0); i++) DoRow(ref i, sheet.Row[i], sheet, sh, fileargs);
                ApplyCellStyle(sh, sheet);
            }

            Console.WriteLine("Sheet created!");
        }

        private void DoRow(ref int rownum, Row row, Sheet sheet, ISheet exSheet, object[] fileargs = null, object[] sheetargs = null)
        {
            //Process Cell
            int cnt = rownum - 1;
            if (row.Sql != SQLType.PlainText)
            {
                using (DbCommand cmd = conn.RowConnection.CreateCommand())
                {

                    cmd.CommandType = row.Sql == SQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = row.Text;

                    if (fileargs != null)
                        for (int i = 0; i < fileargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is SqlConnection ? "@F" + i : "F" + i;
                            param.Value = fileargs[i];
                            cmd.Parameters.Add(param);
                        }

                    if (sheetargs != null)
                        for (int i = 0; i < sheetargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is SqlConnection ? "@S" + i : "S" + i;
                            param.Value = sheetargs[i];
                            cmd.Parameters.Add(param);
                        }

                    using (DbDataReader reader = cmd.ExecuteReader())
                    {
                        var rw = exSheet.CreateRow(++cnt);
                        var cols = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            Cell dc = new Cell();
                            dc.Sql = SQLType.PlainText;
                            dc.Text = cols[i];

                            DoCell(ref i, dc, row, sheet, rw, fileargs, sheetargs, cols);
                        }
                        while (reader.Read())
                        {
                            rw = exSheet.CreateRow(++cnt);
                            Console.WriteLine("Writing row {0} with given SQL...", rw.RowNum);
                            object[] args = new object[reader.FieldCount];
                            reader.GetValues(args);
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Cell dc = new Cell();
                                dc.Sql = SQLType.PlainText;
                                dc.Text = reader.GetValue(i).ToString();

                                DoCell(ref i, dc, row, sheet, rw, fileargs, sheetargs, args);
                            }
                        }
                    }

                }
            }
            else
            {
                IRow rw = exSheet.CreateRow(++cnt);
                Console.WriteLine("Writing row {0}...", rw.RowNum);
                for (int i = 0; i < (row.Cell != null ? row.Cell.Length : 0); i++) DoCell(ref i, row.Cell[i], row, sheet, rw, fileargs, sheetargs);
            }

            rownum = cnt;


        }

        private void DoCell(ref int cellnum, Cell cell, Row row, Sheet sheet, IRow exRow, object[] fileargs = null, object[] sheetargs = null, object[] rowargs = null)
        {
            //Cell Value
            int cnt = cellnum - 1;
            Action<string> cellfunc = val =>
            {
                ICell cl = exRow.CreateCell(++cnt);
                Console.WriteLine("Writing cell {0}...", cl.ColumnIndex);
                cl.CellStyle = excel.CreateCellStyle();

                if (styles.ContainsKey(cell.Style)) cl.CellStyle.CloneStyleFrom(styles[cell.Style]);

                switch ((cell.HAlign ?? string.Empty).ToLower())
                {
                    case "left":
                        cl.CellStyle.Alignment = HorizontalAlignment.Left;
                        break;
                    case "center":
                        cl.CellStyle.Alignment = HorizontalAlignment.Center;
                        break;
                    case "right":
                        cl.CellStyle.Alignment = HorizontalAlignment.Right;
                        break;
                }
                switch ((cell.VAlign ?? string.Empty).ToLower())
                {
                    case "top":
                        cl.CellStyle.VerticalAlignment = VerticalAlignment.Top;
                        break;
                    case "middle":
                        cl.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        break;
                    case "bottom":
                        cl.CellStyle.VerticalAlignment = VerticalAlignment.Bottom;
                        break;
                }

                if (val != null)
                {
                    double nval;
                    DateTime dval;
                    bool isFomula = !string.IsNullOrEmpty(val) && val[0] == '=';
                    string rawval = isFomula ? val.Substring(1) : val;
                    switch (cell.Out)
                    {
                        case OutType.DateTime:
                            if (isFomula) cl.SetCellFormula(rawval); else if(DateTime.TryParse(rawval, out dval)) cl.SetCellValue(dval); else cl.SetCellValue(rawval);
                            break;
                        case OutType.Number:
                            cl.SetCellType(CellType.Numeric);
                            if (isFomula) cl.SetCellFormula(rawval); else if(double.TryParse(rawval, out nval)) cl.SetCellValue(nval); else cl.SetCellValue(rawval);
                            break;
                        case OutType.Normal:
                            cl.SetCellType(CellType.Unknown);
                            if (isFomula) cl.SetCellFormula(rawval); else cl.SetCellValue(rawval);
                            break;
                        default:
                            cl.SetCellType(CellType.String);
                            if (isFomula) cl.SetCellFormula(rawval); else cl.SetCellValue(rawval);
                            break;
                    }
                }

            };

            if (cell.Sql != SQLType.PlainText)
            {
                using (DbCommand cmd = conn.CellConnection.CreateCommand())
                {

                    cmd.CommandType = cell.Sql == SQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = cell.Text;

                    if (fileargs != null)
                        for (int i = 0; i < fileargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is SqlConnection ? "@F" + i : "F" + i;
                            param.Value = fileargs[i];
                            cmd.Parameters.Add(param);
                        }

                    if (sheetargs != null)
                        for (int i = 0; i < sheetargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is SqlConnection ? "@S" + i : "S" + i;
                            param.Value = sheetargs[i];
                            cmd.Parameters.Add(param);
                        }

                    if (rowargs != null)
                        for (int i = 0; i < rowargs.Length; i++)
                        {
                            DbParameter param = cmd.CreateParameter();
                            param.DbType = DbType.String;
                            param.ParameterName = cmd.Connection is SqlConnection ? "@R" + i : "R" + i;
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
