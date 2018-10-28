using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace DOCX
{
    public class Docx : IDisposable
    {
        private readonly ZipArchive _zip;
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
                using (StreamWriter sw = new StreamWriter(entry.Open()))
                {
                    doc.Save(sw);
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