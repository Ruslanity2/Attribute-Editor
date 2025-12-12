using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace AttrKom
{
    [XmlRoot("Operations")]
    public class Operations
    {
        [XmlElement("Operation")]
        public List<OperationItem> Items { get; set; } = new List<OperationItem>();

        /// <summary>
        /// Загрузить XML-файл
        /// </summary>
        public static Operations Load(string xmlFilePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Operations));

            if (!File.Exists(xmlFilePath))
                return new Operations();

            using (StreamReader reader = new StreamReader(xmlFilePath))
            {
                return (Operations)serializer.Deserialize(reader);
            }
        }

        /// <summary>
        /// Сохранить в XML-файл
        /// </summary>
        public void Save(string filePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Operations));
            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                serializer.Serialize(fs, this);
            }
        }
    }

    public class OperationItem
    {
        [XmlAttribute("Name")]
        public string Name { get; set; }

        [XmlAttribute("Code")]
        public string Code { get; set; }
    }
}
