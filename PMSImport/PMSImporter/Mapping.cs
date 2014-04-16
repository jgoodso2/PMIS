using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace PMSImporter
{
    public class Mapping
    {
        public Project Project { get; set; }
        public Task Task { get; set; }
        public Assignment Assignment { get; set; }
        public ProjectResource ProjectResource { get; set; }
        private Dictionary<string, string> projectMap;
        private Dictionary<string, string> projectCustomFieldsMap;
        private Dictionary<string, string> taskCustomFieldsMap;
        private Dictionary<string, string> taskMap;
        private Dictionary<string, string> assignmentMap;
        private Dictionary<string, string> projectResourceMap;

        [XmlIgnore]
        public Dictionary<string, string> ProjectCustomFieldsMap
        {
            get
            {
                if (projectCustomFieldsMap == null)
                {
                    return projectCustomFieldsMap = Project.CustomFields.ToDictionary(t => t.Target, t => t.Source);
                }
                return projectCustomFieldsMap;
            }
        }

        [XmlIgnore]
        public Dictionary<string, string> TaskCustomFieldsMap
        {
            get
            {
                if (taskCustomFieldsMap == null)
                {
                    return taskCustomFieldsMap = Task.CustomFields.ToDictionary(t => t.Target, t => t.Source);
                }
                return taskCustomFieldsMap;
            }
        }

        [XmlIgnore]
        public Dictionary<string, string> ProjectMap
        {
            get
            {
                if (projectMap == null)
                {
                    return projectMap = Project.Fields.ToDictionary(t => t.Target, t => t.Source);
                }
                return projectMap;
            }
        }

        [XmlIgnore]
        public Dictionary<string, string> AssignmentMap
        {
            get
            {
                if (assignmentMap == null)
                {
                    return assignmentMap = Assignment.Fields.ToDictionary(t => t.Target, t => t.Source);
                }
                return assignmentMap;
            }
        }

        [XmlIgnore]
        public Dictionary<string, string> TaskMap
        {
            get
            {
                if (taskMap == null)
                {
                    return taskMap = Task.Fields.ToDictionary(t => t.Target, t => t.Source);
                }
                return taskMap;
            }
        }
        [XmlIgnore]
        public Dictionary<string, string> ProjectResourceMap
        {
            get
            {
                if (projectResourceMap == null)
                {
                    return projectResourceMap = ProjectResource.Fields.ToDictionary(t => t.Source, t => t.Target);
                }
                return projectResourceMap;
            }
        }

        public static Mapping Load(string filename = "FieldMapping.xml")
        {
            Console.WriteLine("Load mapping file started");
            if (File.Exists(filename))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Mapping));
                FileStream stream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                if (stream != null)
                {
                    try
                    {
                        Mapping res = serializer.Deserialize(stream) as Mapping;
                        Console.WriteLine("Load mapping file done successfully");
                        return res;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Mapping  File  Error=" + ex.Message);
                    }
                    finally
                    {
                        stream.Close();
                    }

                }
            }
            
                throw new ArgumentException("Mapping File is missing");
        }
    }
}
