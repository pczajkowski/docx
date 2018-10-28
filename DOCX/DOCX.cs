using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;
using Newtonsoft.Json;

namespace DOCX
{
    public class Docx : IDisposable
    {
        private readonly ZipArchive _zip;
        private readonly string _zipPath;
        private readonly XmlNamespaceManager _ns = new XmlNamespaceManager(new NameTable());

        private readonly Dictionary<string, string> _namespaces = new Dictionary<string, string>
        {
            {"w",  "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        };

        private void LoadNamespaces()
        {
            foreach (var item in _namespaces)
            {
                _ns.AddNamespace(item.Key, item.Value);
            }
        }

        public Docx(string path)
        {
            _zip = ZipFile.Open(path, ZipArchiveMode.Update);
            _zipPath = path;
            LoadNamespaces();
        }

        public Docx(ZipArchive zipArchive)
        {
            _zip = zipArchive;
            LoadNamespaces();
        }

        private (XmlDocument doc, string message) GetXML(ZipArchiveEntry entry)
        {
            XmlDocument doc = new XmlDocument()
            {
                PreserveWhitespace = true // Disables auto-indent
            };

            try
            {
                using (StreamReader sr = new StreamReader(entry.Open()))
                {
                    doc.Load(sr);
                }
            }
            catch (Exception e)
            {
                return (null, $"Error reading {entry.Name}!\n{e}");
            }

            return (doc, "OK");
        }

        private (bool status, string message) SaveXML(XmlDocument doc, ZipArchiveEntry entry)
        {
            try
            {
                using (var sr = entry.Open())
                {
                    sr.SetLength(doc.OuterXml.Length);
                    using (StreamWriter sw = new StreamWriter(sr))
                    {
                        doc.Save(sw);
                    }
                }
            }
            catch (Exception e)
            {
                return (false, $"Error saving {entry.Name}!\n{e}");
            }

            return (true, "OK");
        }

        private (bool status, string message) AddTrackRevisions(ZipArchiveEntry settings)
        {
            var loadResult = GetXML(settings);
            if (loadResult.doc == null)
                return (false, loadResult.message);

            XmlDocument doc = loadResult.doc;

            if (doc.SelectSingleNode("//w:trackRevisions", _ns) != null) return (true, "No change needed.");

            XmlElement trackRevisions = doc.CreateElement("w", "trackRevisions", _namespaces["w"]);
            if (doc.DocumentElement == null)
                return (false, "No root element in settings.xml!");

            doc.DocumentElement.AppendChild(trackRevisions);

            return SaveXML(doc, settings);
        }

        public (bool status, string message) EnableTrackedChanges()
        {
            ZipArchiveEntry settings = _zip.GetEntry(@"word/settings.xml");
            if (settings == null)
                return (false,
            "Can't access settings.xml!");

            var result = AddTrackRevisions(settings);
            return !result.status ? (false, result.message) : (true, "OK");
        }

        private Dictionary<string, string> _authors = new Dictionary<string, string>();

        private string AnonymizeName(string name)
        {
            if (_authors.TryGetValue(name, out var anonymousName))
                return anonymousName;

            anonymousName = $"Author{_authors.Count + 1}";
            _authors.Add(name, anonymousName);
            return anonymousName;
        }
        
        private (bool status, string message) AnonymizeAuthors(ZipArchiveEntry comments)
        {
            var loadResult = GetXML(comments);
            if (loadResult.doc == null)
                return (false, loadResult.message);

            XmlDocument doc = loadResult.doc;

            var commentNodes = doc.SelectNodes("//w:comment", _ns);
            if (commentNodes == null)
                return (false, "There are no comments!");

            foreach (XmlNode node in commentNodes)
            {
                var author = node.Attributes["w:author"];
                author.Value = AnonymizeName(author.Value);
            }

            return SaveXML(doc, comments);
        }
        
        private void SaveAuthors()
        {
            string path = Path.ChangeExtension(_zipPath, "json");
            
            using (StreamWriter sw = new StreamWriter(path))
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                JsonSerializer serializer = new JsonSerializer
                {
                    NullValueHandling = NullValueHandling.Ignore
                };

                serializer.Serialize(writer, _authors);
            }
        }

        
        public (bool status, string message) AnonymizeComments()
        {
            ZipArchiveEntry comments = _zip.GetEntry(@"word/comments.xml");
            if (comments == null)
                return (false,
                    "Can't access comments.xml!");

            var result = AnonymizeAuthors(comments);
            if (!result.status) return (false, result.message);
            
            SaveAuthors();
            return (true, "OK");
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~Docx()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                _zip.Dispose();
            }
        }
    }
}