namespace PMSImporter
{
    public class Field
    {
        [System.Xml.Serialization.XmlAttribute(AttributeName = "Source" )]
        public string Source { get; set; }
        [System.Xml.Serialization.XmlAttribute(AttributeName = "Target")]
        public string Target { get; set; }
         [System.Xml.Serialization.XmlAttribute(AttributeName = "IsResourceColumn")]
        public bool IsResourceColumn { get; set; }
         [System.Xml.Serialization.XmlAttribute(AttributeName = "MapStringToGuid")]
         public bool MapStringToGuid { get; set; }

        public Field()
        {
            IsResourceColumn = false;
            MapStringToGuid = false;
        }
    }
}
