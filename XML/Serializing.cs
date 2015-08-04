using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Serialization;
using NPOI;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF;
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
        /// XML Workbook
        /// </summary>
        private readonly Workbook book = null;
        /// <summary>
        /// Excel Workbook
        /// </summary>
        private readonly IWorkbook excel = null;
        /// <summary>
        /// DB Provider for create DB connection
        /// </summary>
        private readonly DbProviderFactory factory = null;
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
                this.connstr = book.ConnectionString.Text;
                this.factory = DbProviderFactories.GetFactory(book.ConnectionString.Type);
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

        public Serializing Do(out string path)
        {
            Console.WriteLine("Serialize processing...");

            var exfile = book.File;
            string filepath;// = Path.Combine(exfile.Path ?? string.Empty, exfile.Name);

            if (exfile.Sql != SQLType.PlainText)
            {
                using (DbConnection conn = factory.CreateConnection())
                using (DbCommand cmd = conn.CreateCommand())
                {
                    conn.ConnectionString = this.connstr;
                    conn.Open();
                    
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
                Dictionary<string, ICellStyle> styles = new Dictionary<string, ICellStyle>(0);

                if(book.Styles != null)
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

                        styles.Add(style.Id, cellstyle);
                            
                    }

                if(book.Sheet == null || book.Sheet.Length == 0) throw new SerializingException("No sheet specified. You must have least 1 <Sheet> element.");

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
                throw new SerializingException(e) { FilePath = filepath };
            }

            path = filepath;
            return this;
        }

        public Serializing Do()
        {
            string dummy;
            return Do(out dummy);
        }

        private void DoSheet(Sheet sheet)
        {
            

            if (sheet.Sql != SQLType.PlainText)
            {
                using (DbConnection conn = factory.CreateConnection())
                using (DbCommand cmd = conn.CreateCommand())
                {
                    conn.ConnectionString = this.connstr;
                    conn.Open();

                    cmd.CommandType = sheet.Sql == SQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = sheet.Text;

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
                            var irow = 0;
                            if (row != null) DoRow(ref irow, row, sheet, sh, args);

                            var etcrow = sheet.Row != null ? sheet.Row.Where(r => !r.ColumnHeader).ToArray() : new Row[0];
                            for (; irow < etcrow.Length; irow++) DoRow(ref irow, etcrow[irow], sheet, sh, args);

                            idx++;
                        }
                    }

                }
            }
            else
            {
                var sh = excel.CreateSheet(!string.IsNullOrEmpty(sheet.Name) ? sheet.Name : "OKSheet" + (sheetidx++));
                Console.WriteLine("Creating sheet '{0}'...", sh.SheetName);
                for (int i = 0; i < (sheet.Row != null ? sheet.Row.Length : 0); i++) DoRow(ref i, sheet.Row[i], sheet, sh);
            }

            Console.WriteLine("Sheet created!");
        }

        private void DoRow(ref int rownum, Row row, Sheet sheet, ISheet exSheet, object[] sheetargs = null)
        {
            int cnt = rownum - 1;
            if (row.Sql != SQLType.PlainText)
            {
                using (DbConnection conn = factory.CreateConnection())
                using (DbCommand cmd = conn.CreateCommand())
                {
                    conn.ConnectionString = this.connstr;
                    conn.Open();

                    cmd.CommandType = row.Sql == SQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = row.Text;

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
                        while (reader.Read())
                        {
                            var rw = exSheet.CreateRow(++cnt);
                            Console.WriteLine("Writing row {0} with given SQL...", rw.RowNum);
                            object[] args = new object[reader.FieldCount];
                            reader.GetValues(args);
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Cell dc = new Cell();
                                dc.Sql = SQLType.PlainText;
                                dc.Text = reader.GetValue(i).ToString();

                                DoCell(ref i, dc, row, sheet, rw, sheetargs, args);
                            }
                        }
                    }

                }
            }
            else
            {
                IRow rw = exSheet.CreateRow(++cnt);
                Console.WriteLine("Writing row {0}...", rw.RowNum);
                for (int i = 0; i < (row.Cell != null ? row.Cell.Length : 0); i++) DoCell(ref i, row.Cell[i], row, sheet, rw, sheetargs);
            }

            rownum = cnt;


        }

        private void DoCell(ref int cellnum, Cell cell, Row row, Sheet sheet, IRow exRow, object[] sheetargs = null, object[] rowargs = null)
        {
            int cnt = cellnum - 1;
            Action<string> cellfunc = val =>
            {
                ICell cl = exRow.CreateCell(++cnt);
                Console.WriteLine("Writing cell {0}...", cl.ColumnIndex);

                if (val != null)
                {
                    bool isFomula = val[0] == '=';
                    string rawval = isFomula ? val.Substring(1) : val;
                    switch (cell.Out)
                    {
                        default:
                            cl.SetCellType(CellType.String);
                            if (isFomula) cl.SetCellFormula(rawval); else cl.SetCellValue(rawval);
                            break;
                    }
                }

            };

            if (cell.Sql != SQLType.PlainText)
            {
                using (DbConnection conn = factory.CreateConnection())
                using (DbCommand cmd = conn.CreateCommand())
                {
                    conn.ConnectionString = this.connstr;
                    conn.Open();

                    cmd.CommandType = cell.Sql == SQLType.Procedure ? CommandType.StoredProcedure : CommandType.Text;
                    cmd.CommandText = cell.Text;

                    if(sheetargs != null)
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
