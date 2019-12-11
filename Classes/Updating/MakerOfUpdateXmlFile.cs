using System;
using System.Diagnostics.Contracts;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace ASTA.Classes.Updating
{
    public class MakerOfUpdateXmlFile : IMakeable
    {
        private IParameters _parameters { get; set; }

        public delegate void Status(object sender, TextEventArgs e);

        public event Status status;

        public MakerOfUpdateXmlFile()
        {
        }

        public MakerOfUpdateXmlFile(UpdatingParameters parameters)
        {
            _parameters = parameters.Get();
        }

        public void Set(UpdatingParameters parameters)
        {
            _parameters = parameters.Get();
        }

        public UpdatingParameters Get()
        {
            return _parameters.Get();
        }

        public void Make()
        {
            status?.Invoke(this, new TextEventArgs("MakeFile "));

            Contract.Requires(_parameters != null,
                    "Не создан экземпляр UpdatingParameters!");

            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters.Get().appFileXml),
                    "Отсутствует параметр appFileXml или ссылка пустая!");

            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters.Get().appVersion),
                    "Отсутствует параметр appVersion или ссылка пустая!");

            //https://stackoverflow.com/questions/44477727/writing-xml-and-reading-it-back-c-sharp
            //clear any xmlns attributes from the root element
            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add("", "");//clear any xmlns attributes from the root element

            XMLDocument document = new XMLDocument
            {
                version = _parameters.Get().appVersion,
                url = _parameters.Get().appUpdateFolderURL + _parameters.Get().appFileZip
            };

            if (_parameters.Get().appUpdateMD5 != null)
            {
                var checksum = new XMLElementChecksum
                {
                    value = _parameters.Get().appUpdateMD5,
                    algorithm = "MD5"
                };
                document.checksum = checksum;
            }
            status?.Invoke(this, new TextEventArgs($"XML файл: {_parameters.Get().appFileXml}"));

            //  var nodesToStore = new List<XMLDocument> { document };
            try
            {
                using (FileStream fs = new FileStream(_parameters.Get().appFileXml, FileMode.Create))
                {
                    // XmlSerializer serializer = new XmlSerializer(nodesToStore.GetType());
                    // serializer.Serialize(fs, nodesToStore);

                    XmlSerializer serializer = new XmlSerializer(document.GetType());//, atribXmlOver
                    serializer.Serialize(fs, document, ns); //clear any xmlns attributes from the root element
                }
                status?.Invoke(this, new TextEventArgs($"XML файл сохранен как {Path.GetFullPath(_parameters.Get().appFileXml)}"));
            }
            catch
            {
                status?.Invoke(this, new TextEventArgs("Ошибка сохранения XML файла"));
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

        internal XmlDeclaration CreateXmlDeclaration(string v1, string v2, object p)
        {
            throw new NotImplementedException();
        }
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