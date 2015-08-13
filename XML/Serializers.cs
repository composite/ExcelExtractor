// ***********************************************************************
// Assembly         : ExcelExtractor
// Author           : User
// Created          : 07-16-2015
//
// Last Modified By : User
// Last Modified On : 07-16-2015
// ***********************************************************************
// <copyright file="Serializers.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace ExcelExtractor.XML
{
    /// <summary>
    /// Enum SQLType
    /// </summary>
    public enum ExcelSQLType
    {
        /// <summary>
        /// 일반 텍스트 출력
        /// </summary>
        [XmlEnum(Name = "")]
        PlainText = 0,

        /// <summary>
        /// 일반 쿼리 사용
        /// </summary>
        [XmlEnum(Name = "Text")]
        CommandText,
        /// <summary>
        /// 프로시저 사용
        /// </summary>
        [XmlEnum(Name = "Proc")]
        Procedure,
    }

    /// <summary>
    /// Enum OutType
    /// </summary>
    public enum ExcelOutType
    {
        /// <summary>
        /// 텍스트 그대로 출력, 기본값
        /// </summary>
        [XmlEnum(Name = "Text")]
        Text = 0,
        /// <summary>
        /// 텍스트를 엑셀 방식대로 출력
        /// </summary>
        [XmlEnum(Name = "Normal")]
        Normal,
        /// <summary>
        /// 엑셀 숫자 형식으로 출력
        /// </summary>
        [XmlEnum(Name = "Number")]
        Number,
        /// <summary>
        /// 엑셀 날짜 형식으로 출력
        /// </summary>
        [XmlEnum(Name = "Date")]
        Date,
        /// <summary>
        /// 엑셀 시각 형식으로 출력
        /// </summary>
        [XmlEnum(Name = "DateTime")]
        DateTime,
        /// <summary>
        /// 엑셀 시간 형식으로 출력
        /// </summary>
        [XmlEnum(Name = "Time")]
        Time,
        /// <summary>
        /// 엑셀 통화 형식으로 출력
        /// </summary>
        [XmlEnum(Name = "Money")]
        Money,
        /// <summary>
        /// 엑셀 백분율 형식으로 출력
        /// </summary>
        [XmlEnum(Name = "Percent")]
        Percent
    }

    /// <summary>
    /// 엑셀의 셀 나열 방법
    /// </summary>
    public enum ExcelCellFetch
    {
        /// <summary>
        /// 수평으로 나열
        /// </summary>
        [XmlEnum(Name = "Horizontal")]
        Horizontal = 0,
        /// <summary>
        /// 수직으로 나열
        /// </summary>
        [XmlEnum(Name = "Vertical")]
        Vertical
    }

    /// <summary>
    /// 나열 모드 방법
    /// </summary>
    public enum ExcelFetchType
    {
        /// <summary>
        /// 행을 나열
        /// </summary>
        [XmlEnum(Name = "Fetch")]
        Fetch = 0,
        /// <summary>
        /// 열대로 나열
        /// </summary>
        [XmlEnum(Name = "Single")]
        Single
    }

    /// <summary>
    /// 수평 정렬방법
    /// </summary>
    public enum ExcelHAlign
    {
        [XmlEnum(Name = "")]
        Default = 0,
        [XmlEnum(Name = "Left")]
        Left,
        [XmlEnum(Name = "Center")]
        Center,
        [XmlEnum(Name = "Right")]
        Right,
        [XmlEnum(Name = "Justify")]
        Justify
    }

    /// <summary>
    /// 수직 정렬방법
    /// </summary>
    public enum ExcelVAlign
    {
        [XmlEnum(Name = "")]
        Default = 0,
        [XmlEnum(Name = "Top")]
        Top,
        [XmlEnum(Name = "Middle")]
        Middle,
        [XmlEnum(Name = "Bottom")]
        Bottom,
        [XmlEnum(Name = "Justify")]
        Justify
    }

    [XmlRoot(ElementName = "File")]
    public class ExcelFile
    {
        [XmlAttribute(AttributeName = "Name")]
        public string Name { get; set; }
        [XmlAttribute(AttributeName = "Path")]
        public string Path { get; set; }
        [XmlAttribute(AttributeName = "SQL")]
        public ExcelSQLType SQL { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "ConnectionString")]
    public class ExcelConnectionString
    {
        [XmlAttribute(AttributeName = "Type")]
        public string Type { get; set; }
        [XmlAttribute(AttributeName = "Timeout")]
        public int Timeout { get; set; }
        [XmlAttribute(AttributeName = "EncryptFile")]
        public string EncryptFile { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "Cell")]
    public class ExcelCell
    {
        [XmlAttribute(AttributeName = "SQL")]
        public ExcelSQLType SQL { get; set; }
        [XmlAttribute(AttributeName = "Fetch")]
        public ExcelCellFetch Fetch { get; set; }
        [XmlText]
        public string Text { get; set; }
        [XmlAttribute(AttributeName = "Out")]
        public ExcelOutType Out { get; set; }
    }

    [XmlRoot(ElementName = "Row")]
    public class ExcelRow
    {
        [XmlElement(ElementName = "Cell")]
        public List<ExcelCell> Cells { get; set; }
        [XmlElement(ElementName = "Fetch")]
        public ExcelFetch Fetch { get; set; }
        [XmlAttribute(AttributeName = "SQL")]
        public ExcelSQLType SQL { get; set; }
        [XmlAttribute(AttributeName = "ColumnHeader")]
        public bool ColumnHeader { get; set; }
        [XmlAttribute(AttributeName = "DataSet")]
        public bool DataSet { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "Fetch")]
    public class ExcelFetch
    {
        [XmlAttribute(AttributeName = "Type")]
        public ExcelFetchType Type { get; set; }
        [XmlText]
        public string Text { get; set; }
        [XmlAttribute(AttributeName = "SQL")]
        public ExcelSQLType SQL { get; set; }
    }

    [XmlRoot(ElementName = "Sheet")]
    public class ExcelSheet
    {
        [XmlElement(ElementName = "Row")]
        public List<ExcelRow> Rows { get; set; }
        [XmlAttribute(AttributeName = "Name")]
        public string Name { get; set; }
        [XmlElement(ElementName = "Style")]
        public List<ExcelStyle> Style { get; set; }
        [XmlAttribute(AttributeName = "SQL")]
        public ExcelSQLType SQL { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "Align")]
    public class ExcelAlign
    {
        [XmlAttribute(AttributeName = "Horizontal")]
        public ExcelHAlign Horizontal { get; set; }
        [XmlAttribute(AttributeName = "Vertical")]
        public ExcelVAlign Vertical { get; set; }
    }

    [XmlRoot(ElementName = "Font")]
    public class ExcelFont
    {
        [XmlAttribute(AttributeName = "Family")]
        public string Family { get; set; }
        [XmlAttribute(AttributeName = "Size")]
        public float Size { get; set; }
        [XmlAttribute(AttributeName = "Color")]
        public string Color { get; set; }
        [XmlAttribute(AttributeName = "Bold")]
        public bool Bold { get; set; }
        [XmlAttribute(AttributeName = "Italic")]
        public bool Italic { get; set; }
        [XmlAttribute(AttributeName = "Underline")]
        public bool Underline { get; set; }
        [XmlAttribute(AttributeName = "Strike")]
        public bool Strike { get; set; }
    }

    [XmlRoot(ElementName = "Back")]
    public class ExcelBackColor
    {
        [XmlAttribute(AttributeName = "Color")]
        public string Color { get; set; }
        [XmlAttribute(AttributeName = "Pattern")]
        public string Pattern { get; set; }
    }

    public abstract class ExcelBorderBase
    {
        [XmlAttribute(AttributeName = "Color")]
        public string Color { get; set; }
        [XmlAttribute(AttributeName = "Style")]
        public string Style { get; set; }
    }

    [XmlRoot(ElementName = "All")]
    public class ExcelBorderAll : ExcelBorderBase
    {

    }

    [XmlRoot(ElementName = "Top")]
    public class ExcelBorderTop : ExcelBorderBase
    {

    }

    [XmlRoot(ElementName = "Left")]
    public class ExcelBorderLeft : ExcelBorderBase
    {

    }

    [XmlRoot(ElementName = "Right")]
    public class ExcelBorderRight : ExcelBorderBase
    {

    }

    [XmlRoot(ElementName = "Bottom")]
    public class ExcelBorderBottom : ExcelBorderBase
    {

    }

    [XmlRoot(ElementName = "Border")]
    public class ExcelBorder
    {
        [XmlElement(ElementName = "All")]
        public ExcelBorderAll All { get; set; }
        [XmlElement(ElementName = "Top")]
        public ExcelBorderTop Top { get; set; }
        [XmlElement(ElementName = "Left")]
        public ExcelBorderLeft Left { get; set; }
        [XmlElement(ElementName = "Right")]
        public ExcelBorderRight Right { get; set; }
        [XmlElement(ElementName = "Bottom")]
        public ExcelBorderBottom Bottom { get; set; }
    }

    [XmlRoot(ElementName = "Style")]
    public class ExcelStyle
    {
        [XmlElement(ElementName = "Align")]
        public ExcelAlign Align { get; set; }
        [XmlElement(ElementName = "Font")]
        public ExcelFont Font { get; set; }
        [XmlElement(ElementName = "Back")]
        public ExcelBackColor Back { get; set; }
        [XmlElement(ElementName = "Border")]
        public ExcelBorder Border { get; set; }
        [XmlAttribute(AttributeName = "Range")]
        public string Range { get; set; }
        [XmlAttribute(AttributeName = "ColGroup")]
        public string ColGroup { get; set; }
        [XmlAttribute(AttributeName = "RowGroup")]
        public string RowGroup { get; set; }
    }

    public abstract class ExcelEventBase
    {
        [XmlAttribute(AttributeName = "CMD")]
        public string CMD { get; set; }
        [XmlAttribute(AttributeName = "SQL")]
        public ExcelSQLType SQL { get; set; }
        [XmlAttribute(AttributeName = "Specific")]
        public bool Specific { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "Before")]
    public class ExcelBefore : ExcelEventBase
    {

    }

    [XmlRoot(ElementName = "After")]
    public class ExcelAfter : ExcelEventBase
    {

    }

    [XmlRoot(ElementName = "Workbook")]
    public class ExcelWorkbook
    {
        [XmlElement(ElementName = "File")]
        public ExcelFile File { get; set; }
        [XmlElement(ElementName = "ConnectionString")]
        public ExcelConnectionString ConnectionString { get; set; }
        [XmlElement(ElementName = "Sheet")]
        public List<ExcelSheet> Sheets { get; set; }
        [XmlElement(ElementName = "Before")]
        public ExcelBefore Before { get; set; }
        [XmlElement(ElementName = "After")]
        public ExcelAfter After { get; set; }
        [XmlAttribute(AttributeName = "Label")]
        public string Label { get; set; }
    }
}
