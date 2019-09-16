using System;
using System.IO;
using System.Xml.Serialization;

namespace ASTA
{
    class MakerOfUpdateAppXML
    {
        string _appVersion;
        string _appXmlLocalPath;
        string _appUpdateFileURL;
        string _appUpdateChangeLogURL;
        string _appMD5;
        public string Status { get; private set; }

        public MakerOfUpdateAppXML(string appVersion, string appXmlLocalPath, string appUpdateFileURL, string appUpdateChangeLogURL, string appMD5)
        {
            _appVersion = appVersion;
            _appXmlLocalPath = appXmlLocalPath;
            _appUpdateFileURL = appUpdateFileURL;
            _appUpdateChangeLogURL = appUpdateChangeLogURL;
            _appMD5 = appMD5;
        }

        public void SaveXML()
        {
            Status = null;
            if (_appXmlLocalPath == null)
            {
                throw new NullReferenceException();
            }

            //https://stackoverflow.com/questions/44477727/writing-xml-and-reading-it-back-c-sharp

            XMLDocument document = new XMLDocument();

            //clear any xmlns attributes from the root element
            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("","");//clear any xmlns attributes from the root element

            if (_appVersion != null)
                document.version = _appVersion;
            if (_appUpdateFileURL != null)
                document.url = _appUpdateFileURL;
            if (_appUpdateChangeLogURL != null)
                document.changelogUrl = _appUpdateChangeLogURL;
            if (_appMD5 != null)
            {
                var checksum = new XMLElementChecksum();
                checksum.value = _appMD5;
                checksum.algorithm = "MD5";
                document.checksum = checksum;
            }

            //  var nodesToStore = new List<XMLDocument> { document };

            using (FileStream fs = new FileStream(_appXmlLocalPath, FileMode.Create))
            {
                

                // XmlSerializer serializer = new XmlSerializer(nodesToStore.GetType());
                // serializer.Serialize(fs, nodesToStore);
                XmlSerializer serializer = new XmlSerializer(document.GetType());//, atribXmlOver
                serializer.Serialize(fs, document, ns); //clear any xmlns attributes from the root element
                Status = "XML файл сохранен как " + _appXmlLocalPath;
            }

            /* var readNodes = new List<document>();
             using (FileStream fs2 = new FileStream(filepath, FileMode.Open))
             {
                 XmlSerializer serializer = new XmlSerializer(nodesToStore.GetType());
                 readNodes = serializer.Deserialize(fs2) as List<document>;
             }*/
        }
        
        /*
            private void CreateAppXMLFile()
        {
            //calculate app's MD5
            appFileMD5 = CalculateMD5(appFilePath);

            //write ver of application on the disk
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            System.Xml.XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);   //the xml declaration is recommended, but not mandatory
            System.Xml.XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            System.Xml.XmlElement item = doc.CreateElement(string.Empty, "item", string.Empty);     //string.Empty makes cleaner code
            doc.AppendChild(item);

            System.Xml.XmlElement versionAssemblyInXML = doc.CreateElement(string.Empty, "version", string.Empty);
            System.Xml.XmlText versionInXML = doc.CreateTextNode(appVersionAssembly);
            versionAssemblyInXML.AppendChild(versionInXML);
            item.AppendChild(versionAssemblyInXML);

            System.Xml.XmlElement urlZip = doc.CreateElement(string.Empty, "url", string.Empty);
            System.Xml.XmlText urlText = doc.CreateTextNode(appUpdateFolderURL + appNameZIP); //
            urlZip.AppendChild(urlText);
            item.AppendChild(urlZip);

            System.Xml.XmlElement changelog = doc.CreateElement(string.Empty, "changelog", string.Empty);
            System.Xml.XmlText changelogText = doc.CreateTextNode(appUpdateFolderURL + "urlLog");
            changelog.AppendChild(changelogText);
            item.AppendChild(changelog);

            System.Xml.XmlElement checksum = doc.CreateElement(string.Empty, "checksum", string.Empty);
            System.Xml.XmlText checksumText = doc.CreateTextNode(appFileMD5); //

            System.Xml.XmlAttribute algorithm = doc.CreateAttribute("algorithm");
            algorithm.Value = "MD5";

            checksum.Attributes.Append(algorithm);
            checksum.AppendChild(checksumText);
            item.AppendChild(checksum);

            
          // <changelog>https://github.com/ravibpatel/AutoUpdater.NET/releases</changelog>
          // <checksum algorithm="MD5">Update file Checksum</checksum>
           

            doc.Save(appNameXML);
        }
         */
    }

    [XmlRoot(ElementName = "item", IsNullable = false)]
    public class XMLDocument //Класс должен иметь модификатор public 
    {
        //Класс для сериализации должен иметь стандартный конструктор без параметров. 
        //поля или свойства с модификатором private, при сериализации будут игнорироваться. 
        [XmlElement]
        public string version { get; set; }
        public string url { get; set; }
        public string changelogUrl { get; set; }
        public XMLElementChecksum checksum { get; set; }
    }

    public class XMLElementChecksum //Класс должен иметь модификатор public 
    {
        //Класс для сериализации должен иметь стандартный конструктор без параметров. 
        //поля или свойства с модификатором private, при сериализации будут игнорироваться. 

        [XmlText]
        public string value { get; set; }

        [XmlAttribute]
        public string algorithm { get; set; }
    }
}
