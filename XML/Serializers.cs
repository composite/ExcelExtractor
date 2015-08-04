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
    public enum SQLType
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
    public enum OutType
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
    /// Class Bit.
    /// </summary>
    public class Bit : IEquatable<Bit>, IXmlSerializable
    {
        /// <summary>
        /// The origin
        /// </summary>
        private string origin;
        /// <summary>
        /// The value
        /// </summary>
        private bool value;

        public Bit()
        {
            this.origin = null;
            this.value = false;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Bit"/> class.
        /// </summary>
        /// <param name="value">if set to <c>true</c> [value].</param>
        public Bit(bool value)
        {
            this.origin = (this.value = value).ToString();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Bit"/> class.
        /// </summary>
        /// <param name="value">The value.</param>
        public Bit(string value)
        {
            this.origin = value;
            this.value = Parse(value);
        }

        /// <summary>
        /// Parses the specified value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        private static bool Parse(string value)
        {
            return "true".Equals(value, StringComparison.OrdinalIgnoreCase) || "1".Equals(value);
        }

        /// <summary>
        /// 현재 개체가 동일한 형식의 다른 개체와 같은지 여부를 나타냅니다.
        /// </summary>
        /// <param name="other">이 개체와 비교할 개체입니다.</param>
        /// <returns>현재 개체가 
        /// <paramref name="other" /> 매개 변수와 같으면 true이고, 그렇지 않으면 false입니다.</returns>
        public bool Equals(Bit other)
        {
            return this.value == other.value;
        }

        /// <summary>
        /// Determines whether the specified <see cref="System.Object" /> is equal to this instance.
        /// </summary>
        /// <param name="obj">현재 <see cref="T:System.Object" />와 비교할 <see cref="T:System.Object" />입니다.</param>
        /// <returns><c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.</returns>
        public override bool Equals(object obj)
        {
            return obj is Bit && this.Equals((Bit) obj);
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table.</returns>
        public override int GetHashCode()
        {
            return this.value.GetHashCode();
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>A <see cref="System.String" /> that represents this instance.</returns>
        public override string ToString()
        {
            return this.origin;
        }

        /// <summary>
        /// 이 메서드는 예약되어 있으므로 사용해서는 안 됩니다.IXmlSerializable 인터페이스를 구현할 때 이 메서드에서 null(Visual Basic에서는 Nothing)을 반환해야 하지만 사용자 지정 스키마를 지정해야 하는 경우에는 <see cref="T:System.Xml.Serialization.XmlSchemaProviderAttribute" />를 클래스에 적용합니다.
        /// </summary>
        /// <returns><see cref="M:System.Xml.Serialization.IXmlSerializable.WriteXml(System.Xml.XmlWriter)" /> 메서드에 의해 생성되고 <see cref="M:System.Xml.Serialization.IXmlSerializable.ReadXml(System.Xml.XmlReader)" /> 메서드가 사용하는 개체의 XML 표현을 설명하는 <see cref="T:System.Xml.Schema.XmlSchema" />입니다.</returns>
        public XmlSchema GetSchema()
        {
            return null;
        }

        /// <summary>
        /// 개체의 XML 표현에서 개체를 생성합니다.
        /// </summary>
        /// <param name="reader">개체가 deserialize되는 <see cref="T:System.Xml.XmlReader" /> 스트림입니다.</param>
        public void ReadXml(XmlReader reader)
        {
            this.origin = reader.ReadContentAsString();
            this.value = Parse(this.origin);
        }

        /// <summary>
        /// 개체를 XML 표현으로 변환합니다.
        /// </summary>
        /// <param name="writer">개체가 serialize되는 <see cref="T:System.Xml.XmlWriter" /> 스트림입니다.</param>
        public void WriteXml(XmlWriter writer)
        {
            writer.WriteValue(this.origin);
        }

        /// <summary>
        /// Performs an implicit conversion from <see cref="Bit"/> to <see cref="System.Boolean"/>.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The result of the conversion.</returns>
        public static implicit operator bool(Bit value)
        {
            return value != null ? value.value : false;
        }
    }

    public class XMLNull<T> : IEquatable<T>, IXmlSerializable where T : struct
    {
        private T value;

        public XMLNull()
        {
            this.Value = default(T);
        } 

        public XMLNull(T value)
        {
            this.value = value;
        }

        public T Value { get { return this.value; } set { this.value = value; } }

        public bool Equals(T other)
        {
            return this.value.Equals(other);
        }

        public override bool Equals(object obj)
        {
            return this.value.Equals(obj);
        }

        public static implicit operator T(XMLNull<T> obj)
        {
            return obj.value;
        }

        public XmlSchema GetSchema()
        {
            return null;
        }

        public void ReadXml(XmlReader reader)
        {
            this.value = (T) reader.ReadElementContentAs(typeof (T), null);
        }

        public void WriteXml(XmlWriter writer)
        {
            writer.WriteValue(this.value);
        }

        
    }

    /// <summary>
    /// Class File.
    /// </summary>
    [XmlRoot(ElementName = "File"), Serializable]
    public class ExcelFile
    {
        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The name.</value>
        [XmlAttribute(AttributeName = "Name")]
        public string Name { get; set; }
        /// <summary>
        /// Gets or sets the path.
        /// </summary>
        /// <value>The path.</value>
        [XmlAttribute(AttributeName = "Path")]
        public string Path { get; set; }
        /// <summary>
        /// Gets or sets the SQL.
        /// </summary>
        /// <value>The SQL.</value>
        [XmlAttribute(AttributeName = "SQL")]
        public SQLType Sql { get; set; }
        /// <summary>
        /// Gets or sets the text.
        /// </summary>
        /// <value>The text.</value>
        [XmlText]
        public string Text { get; set; }
    }

    /// <summary>
    /// Class ConnectionString.
    /// </summary>
    [XmlRoot(ElementName = "ConnectionString"), Serializable]
    public class ConnectionString
    {
        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        /// <value>The type.</value>
        [XmlAttribute(AttributeName = "Type")]
        public string Type { get; set; }
        /// <summary>
        /// Gets or sets the encrypt file.
        /// </summary>
        /// <value>The encrypt file.</value>
        [XmlAttribute(AttributeName = "EncryptFile")]
        public string EncryptFile { get; set; }
        /// <summary>
        /// Gets or sets the text.
        /// </summary>
        /// <value>The text.</value>
        [XmlText]
        public string Text { get; set; }
    }

    /// <summary>
    /// Class Font.
    /// </summary>
    [XmlRoot(ElementName = "Font"), Serializable]
    public class ExcelFont
    {
        /// <summary>
        /// Gets or sets the color.
        /// </summary>
        /// <value>The color.</value>
        [XmlAttribute(AttributeName = "Color")]
        public string Color { get; set; }
        /// <summary>
        /// Gets or sets the face.
        /// </summary>
        /// <value>The face.</value>
        [XmlAttribute(AttributeName = "Face")]
        public string Face { get; set; }
        /// <summary>
        /// Gets or sets the face.
        /// </summary>
        /// <value>The face.</value>
        [XmlAttribute(AttributeName = "Size")]
        public double Size { get; set; }
        /// <summary>
        /// Gets or sets the bold.
        /// </summary>
        /// <value>The face.</value>
        [XmlAttribute(AttributeName = "Bold")]
        public bool Bold { get; set; }
        /// <summary>
        /// Gets or sets the italic.
        /// </summary>
        /// <value>The face.</value>
        [XmlAttribute(AttributeName = "Italic")]
        public bool Italic { get; set; }
        /// <summary>
        /// Gets or sets the underline.
        /// </summary>
        /// <value>The face.</value>
        [XmlAttribute(AttributeName = "Underline")]
        public bool Underline { get; set; }
        /// <summary>
        /// Gets or sets the strikethrough.
        /// </summary>
        /// <value>The face.</value>
        [XmlAttribute(AttributeName = "Strike")]
        public bool Strike { get; set; }
    }

    /// <summary>
    /// Class Back.
    /// </summary>
    [XmlRoot(ElementName = "Back"), Serializable]
    public class Back
    {
        /// <summary>
        /// Gets or sets the color.
        /// </summary>
        /// <value>The color.</value>
        [XmlAttribute(AttributeName = "Color")]
        public string Color { get; set; }
    }

    /// <summary>
    /// class BorderAttributes
    /// </summary>
    [Serializable]
    public abstract class BorderStyle
    {
        /// <summary>
        /// Gets or sets the width.
        /// </summary>
        /// <value>The width.</value>
        [XmlAttribute(AttributeName = "Width")]
        public string Width { get; set; }
        /// <summary>
        /// Gets or sets the color.
        /// </summary>
        /// <value>The color.</value>
        [XmlAttribute(AttributeName = "Color")]
        public string Color { get; set; }
    }

    /// <summary>
    /// Class All.
    /// </summary>
    [XmlRoot(ElementName = "All"), Serializable]
    public class All : BorderStyle
    {

    }

    /// <summary>
    /// Class Top.
    /// </summary>
    [XmlRoot(ElementName = "Top"), Serializable]
    public class Top : BorderStyle
    {
        
    }

    /// <summary>
    /// Class Left.
    /// </summary>
    [XmlRoot(ElementName = "Left"), Serializable]
    public class Left : BorderStyle
    {

    }

    /// <summary>
    /// Class Right.
    /// </summary>
    [XmlRoot(ElementName = "Right"), Serializable]
    public class Right : BorderStyle
    {

    }

    /// <summary>
    /// Class Bottom.
    /// </summary>
    [XmlRoot(ElementName = "Bottom"), Serializable]
    public class Bottom : BorderStyle
    {

    }

    /// <summary>
    /// Class Border.
    /// </summary>
    [XmlRoot(ElementName = "Border"), Serializable]
    public class Border
    {
        /// <summary>
        /// Gets or sets all.
        /// </summary>
        /// <value>All.</value>
        [XmlElement(ElementName = "All")]
        public All All { get; set; }
        /// <summary>
        /// Gets or sets the top.
        /// </summary>
        /// <value>The top.</value>
        [XmlElement(ElementName = "Top")]
        public Top Top { get; set; }
        /// <summary>
        /// Gets or sets the left.
        /// </summary>
        /// <value>The left.</value>
        [XmlElement(ElementName = "Left")]
        public Left Left { get; set; }
        /// <summary>
        /// Gets or sets the right.
        /// </summary>
        /// <value>The right.</value>
        [XmlElement(ElementName = "Right")]
        public Right Right { get; set; }
        /// <summary>
        /// Gets or sets the bottom.
        /// </summary>
        /// <value>The bottom.</value>
        [XmlElement(ElementName = "Bottom")]
        public Bottom Bottom { get; set; }
    }

    /// <summary>
    /// Class Style.
    /// </summary>
    [XmlRoot(ElementName = "Style"), Serializable]
    public class Style
    {
        /// <summary>
        /// Gets or sets the font.
        /// </summary>
        /// <value>The font.</value>
        [XmlElement(ElementName = "Font")]
        public ExcelFont Font { get; set; }
        /// <summary>
        /// Gets or sets the back.
        /// </summary>
        /// <value>The back.</value>
        [XmlElement(ElementName = "Back")]
        public Back Back { get; set; }
        /// <summary>
        /// Gets or sets the border.
        /// </summary>
        /// <value>The border.</value>
        [XmlElement(ElementName = "Border")]
        public Border Border { get; set; }
        /// <summary>
        /// Gets or sets the identifier.
        /// </summary>
        /// <value>The identifier.</value>
        [XmlAttribute(AttributeName = "ID")]
        public string Id { get; set; }
    }

    /// <summary>
    /// Class Styles.
    /// </summary>
    [XmlRoot(ElementName = "Styles"), Serializable]
    public class Styles
    {
        /// <summary>
        /// Gets or sets the style.
        /// </summary>
        /// <value>The style.</value>
        [XmlElement(ElementName = "Style")]
        public Style[] Style { get; set; }
    }

    /// <summary>
    /// Class Cell.
    /// </summary>
    [XmlRoot(ElementName = "Cell"), Serializable]
    public class Cell
    {
        /// <summary>
        /// Gets or sets the SQL.
        /// </summary>
        /// <value>The SQL.</value>
        [XmlAttribute(AttributeName = "SQL")]
        public SQLType Sql { get; set; }
        /// <summary>
        /// Gets or sets the text.
        /// </summary>
        /// <value>The text.</value>
        [XmlText]
        public string Text { get; set; }
        /// <summary>
        /// Gets or sets the h align.
        /// </summary>
        /// <value>The h align.</value>
        [XmlAttribute(AttributeName = "HAlign")]
        public string HAlign { get; set; }
        /// <summary>
        /// Gets or sets the v align.
        /// </summary>
        /// <value>The v align.</value>
        [XmlAttribute(AttributeName = "VAlign")]
        public string VAlign { get; set; }
        /// <summary>
        /// Gets or sets the style.
        /// </summary>
        /// <value>The style.</value>
        [XmlAttribute(AttributeName = "Style")]
        public string Style { get; set; }
        /// <summary>
        /// Gets or sets the out.
        /// </summary>
        /// <value>The out.</value>
        [XmlAttribute(AttributeName = "Out")]
        public OutType Out { get; set; }
    }

    /// <summary>
    /// Class Row.
    /// </summary>
    [XmlRoot(ElementName = "Row"), Serializable]
    public class Row
    {
        /// <summary>
        /// Gets or sets the cell.
        /// </summary>
        /// <value>The cell.</value>
        [XmlElement(ElementName = "Cell")]
        public Cell[] Cell { get; set; }
        /// <summary>
        /// Gets or sets the style.
        /// </summary>
        /// <value>The style.</value>
        [XmlAttribute(AttributeName = "Style")]
        public string Style { get; set; }
        /// <summary>
        /// Gets or sets the SQL.
        /// </summary>
        /// <value>The SQL.</value>
        [XmlAttribute(AttributeName = "SQL")]
        public SQLType Sql { get; set; }
        /// <summary>
        /// Gets or sets the column header.
        /// </summary>
        /// <value>The column header.</value>
        [XmlAttribute(AttributeName = "ColumnHeader")]
        public bool ColumnHeader { get; set; }
        /// <summary>
        /// Gets or sets the text.
        /// </summary>
        /// <value>The text.</value>
        [XmlText]
        public string Text { get; set; }
    }

    /// <summary>
    /// Class Sheet.
    /// </summary>
    [XmlRoot(ElementName = "Sheet"), Serializable]
    public class Sheet
    {
        /// <summary>
        /// Gets or sets the row.
        /// </summary>
        /// <value>The row.</value>
        [XmlElement(ElementName = "Row")]
        public Row[] Row { get; set; }
        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The name.</value>
        [XmlAttribute(AttributeName = "Name")]
        public string Name { get; set; }
        /// <summary>
        /// Gets or sets the header style.
        /// </summary>
        /// <value>The header style.</value>
        [XmlAttribute(AttributeName = "HeaderStyle")]
        public string HeaderStyle { get; set; }
        /// <summary>
        /// Gets or sets the even style.
        /// </summary>
        /// <value>The even style.</value>
        [XmlAttribute(AttributeName = "EvenStyle")]
        public string EvenStyle { get; set; }
        /// <summary>
        /// Gets or sets the odd style.
        /// </summary>
        /// <value>The odd style.</value>
        [XmlAttribute(AttributeName = "OddStyle")]
        public string OddStyle { get; set; }
        /// <summary>
        /// Gets or sets the footer style.
        /// </summary>
        /// <value>The footer style.</value>
        [XmlAttribute(AttributeName = "FooterStyle")]
        public string FooterStyle { get; set; }
        /// <summary>
        /// Gets or sets the h group.
        /// </summary>
        /// <value>The h group.</value>
        [XmlElement(ElementName = "HGroup")]
        public HGroup[] HGroup { get; set; }
        /// <summary>
        /// Gets or sets the v group.
        /// </summary>
        /// <value>The v group.</value>
        [XmlElement(ElementName = "VGroup")]
        public VGroup[] VGroup { get; set; }
        /// <summary>
        /// Gets or sets the SQL.
        /// </summary>
        /// <value>The SQL.</value>
        [XmlAttribute(AttributeName = "SQL")]
        public SQLType Sql { get; set; }
        /// <summary>
        /// Gets or sets the Text.
        /// </summary>
        [XmlText]
        public string Text { get; set; }
    }

    /// <summary>
    /// Class CellGroup
    /// </summary>
    [Serializable]
    public abstract class CellGroup
    {
        /// <summary>
        /// Gets or sets the index.
        /// </summary>
        /// <value>The index.</value>
        [XmlAttribute(AttributeName = "Index")]
        public string Index { get; set; }
        /// <summary>
        /// Gets or sets the style.
        /// </summary>
        /// <value>The style.</value>
        [XmlAttribute(AttributeName = "Style")]
        public string Style { get; set; }
        /// <summary>
        /// Gets or sets the h align.
        /// </summary>
        /// <value>The h align.</value>
        [XmlAttribute(AttributeName = "HAlign")]
        public string HAlign { get; set; }
        /// <summary>
        /// Gets or sets the v align.
        /// </summary>
        /// <value>The v align.</value>
        [XmlAttribute(AttributeName = "VAlign")]
        public string VAlign { get; set; }
        /// <summary>
        /// Gets or sets the index group.
        /// </summary>
        /// <value>The index group.</value>
        [XmlAttribute(AttributeName = "IndexGroup")]
        public string IndexGroup { get; set; }
    }

    /// <summary>
    /// Class HGroup.
    /// </summary>
    [XmlRoot(ElementName = "HGroup"), Serializable]
    public class HGroup : CellGroup
    {
        
    }

    /// <summary>
    /// Class VGroup.
    /// </summary>
    [XmlRoot(ElementName = "VGroup"), Serializable]
    public class VGroup : CellGroup
    {

    }

    /// <summary>
    /// Class PrePostProccessors
    /// </summary>
    [Serializable]
    public abstract class PrePostProccessor
    {
        /// <summary>
        /// Gets or sets the command.
        /// </summary>
        /// <value>The command.</value>
        [XmlAttribute(AttributeName = "CMD")]
        public string Cmd { get; set; }
        /// <summary>
        /// Gets or sets the SQL.
        /// </summary>
        /// <value>The SQL.</value>
        [XmlAttribute(AttributeName = "SQL")]
        public SQLType Sql { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="Before"/> is specific.
        /// </summary>
        /// <value><c>true</c> if specific; otherwise, <c>false</c>.</value>
        [XmlAttribute(AttributeName = "Specific")]
        public bool Specific { get; set; }
        /// <summary>
        /// Gets or sets the text.
        /// </summary>
        /// <value>The text.</value>
        [XmlText]
        public string Text { get; set; }
    }

    /// <summary>
    /// Class Before.
    /// </summary>
    [XmlRoot(ElementName = "Before"), Serializable]
    public class Before : PrePostProccessor
    {
        
    }

    /// <summary>
    /// Class After.
    /// </summary>
    [XmlRoot(ElementName = "After"), Serializable]
    public class After : PrePostProccessor
    {

    }

    /// <summary>
    /// 엑셀 최상위 요소 클래스.
    /// 반드시 이 클래스로 Serialize/Deserialize 할 것.
    /// </summary>
    [XmlRoot(ElementName = "Workbook"), Serializable]
    public class Workbook
    {
        /// <summary>
        /// Gets or sets the file.
        /// </summary>
        /// <value>The file.</value>
        [XmlElement(ElementName = "File")]
        public ExcelFile File { get; set; }
        /// <summary>
        /// Gets or sets the connection string.
        /// </summary>
        /// <value>The connection string.</value>
        [XmlElement(ElementName = "ConnectionString")]
        public ConnectionString ConnectionString { get; set; }
        /// <summary>
        /// Gets or sets the styles.
        /// </summary>
        /// <value>The styles.</value>
        [XmlElement(ElementName = "Styles")]
        public Styles Styles { get; set; }
        /// <summary>
        /// Gets or sets the sheet.
        /// </summary>
        /// <value>The sheet.</value>
        [XmlElement(ElementName = "Sheet")]
        public Sheet[] Sheet { get; set; }
        /// <summary>
        /// Gets or sets the before.
        /// </summary>
        /// <value>The before.</value>
        [XmlElement(ElementName = "Before")]
        public Before Before { get; set; }
        /// <summary>
        /// Gets or sets the after.
        /// </summary>
        /// <value>The after.</value>
        [XmlElement(ElementName = "After")]
        public After After { get; set; }
        /// <summary>
        /// Gets or sets the XLSX.
        /// </summary>
        /// <value>The XLSX.</value>
        [XmlAttribute(AttributeName = "XLSX")]
        public bool Xlsx { get; set; }
    }
}
