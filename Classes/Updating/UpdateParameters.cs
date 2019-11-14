namespace ASTA.Classes.Updating
{
    public class UpdatingParameters: IParameters
    {
        public static string Name { get { return nameof(UpdatingParameters); } }
        public string appFileZip { get; set; }
        public string appUpdateMD5 { get; set; }
        public string appFileXml { get; set; }
        public string appVersion { get; set; }
        public string remoteFolderUpdatingURL { get; set; }
        public string localFolderUpdatingURL { get; set; }
        public string appUpdateFolderURL { get; set; }
        public string appUpdateURL { get; set; }
        public string appUpdateFolderURI { get; set; }
        public string appUpdateChangeLogURL { get; set; }

        public UpdatingParameters() { }

        public UpdatingParameters(UpdatingParameters parameters)
        {
            remoteFolderUpdatingURL = parameters.remoteFolderUpdatingURL;
            localFolderUpdatingURL = parameters.localFolderUpdatingURL;
            appUpdateFolderURL = parameters.appUpdateFolderURL;
            appUpdateFolderURI = parameters.appUpdateFolderURI;
            appUpdateURL = parameters.appUpdateURL;
            appFileXml = parameters.appFileXml;
            appUpdateChangeLogURL = parameters.appUpdateChangeLogURL;
            appUpdateMD5 = parameters.appUpdateMD5;
            appVersion = parameters.appVersion;
            appFileZip = parameters.appFileZip;
        }

        public void SetParameters(UpdatingParameters parameters)
        {
            remoteFolderUpdatingURL = parameters.remoteFolderUpdatingURL;
            localFolderUpdatingURL = parameters.localFolderUpdatingURL;
            appUpdateFolderURL = parameters.appUpdateFolderURL;
            appUpdateFolderURI = parameters.appUpdateFolderURI;
            appUpdateURL = parameters.appUpdateURL;
            appFileXml = parameters.appFileXml;
            appUpdateChangeLogURL = parameters.appUpdateChangeLogURL;
            appUpdateMD5 = parameters.appUpdateMD5;
            appVersion = parameters.appVersion;
            appFileZip = parameters.appFileZip;
        }


        public UpdatingParameters GetParameters()
        {
            return new UpdatingParameters
            {
                remoteFolderUpdatingURL = remoteFolderUpdatingURL,
                localFolderUpdatingURL = localFolderUpdatingURL,
                appUpdateFolderURL = appUpdateFolderURL,
                appUpdateFolderURI = appUpdateFolderURI,
                appUpdateURL = appUpdateURL,
                appFileXml = appFileXml,
                appUpdateChangeLogURL = appUpdateChangeLogURL,
                appUpdateMD5 = appUpdateMD5,
                appVersion = appVersion,
                appFileZip = appFileZip
            };
        }
    }

  public  interface IParameters
    {
        UpdatingParameters GetParameters();
        void SetParameters(UpdatingParameters parameters);
    }
}
