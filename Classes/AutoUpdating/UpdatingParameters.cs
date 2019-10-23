namespace ASTA.Classes.AutoUpdating
{
    public class UpdatingParameters
    {
        public static string Name { get { return nameof(UpdatingParameters); } }
        public string appUpdateMD5 { get; set; }
        public string appFileXml { get; set; }
        public string appVersion { get; set; }
        public string serverUpdateURL { get; set; }
        public string appUpdateFolderURL { get; set; }
        public string appUpdateURL { get; set; }
        public string appUpdateFolderURI { get; set; }
        public string appUpdateChangeLogURL { get; set; }
    }

}
