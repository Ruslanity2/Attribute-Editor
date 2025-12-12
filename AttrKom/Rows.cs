using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace AttrKom
{
    [XmlRoot("Rows")]
    public class Rows
    {
        [XmlElement("Row")]
        public List<AttributeRow> Items { get; set; } = new List<AttributeRow>();

        /// <summary>
        /// Загрузить XML-файл
        /// </summary>
        public static Rows Load(string xmlFilePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Rows));

            if (!File.Exists(xmlFilePath))
                return new Rows();

            using (StreamReader reader = new StreamReader(xmlFilePath))
            {
                return (Rows)serializer.Deserialize(reader);
            }
        }

        /// <summary>
        /// Сохранить в XML-файл
        /// </summary>
        public void Save(string filePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Rows));
            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                serializer.Serialize(fs, this);
            }
        }
    }

    public class AttributeRow
    {
        [XmlAttribute("AttrName")]
        public string AttrName { get; set; }
    }
}
