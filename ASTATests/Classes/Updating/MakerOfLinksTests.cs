using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Shouldly;

namespace ASTA.Classes.AutoUpdating.Tests
{
    [TestClass()]
    public class MakerOfLinksTests
    {
        [TestMethod()]
        public void MakeShouldThrowIfServerUpdateURLIsEmpty()
        {
            UpdatingParameters _parameters = new UpdatingParameters();

            Should.Throw<Exception>(() => new MakerOfLinks(_parameters).Make())
            .Message.ShouldBe("Отсутствует параметр remoteFolderUpdatingURL или ссылка пустая!");
        }
        
        [TestMethod()]
        public void MakeShouldThrowIfUpdatingParametersIsNull()
        {
            UpdatingParameters _parameters=null;

            Should.Throw<Exception>(() => new MakerOfLinks(_parameters).Make())
            .Message.ShouldBe("Не создан экземпляр UpdatingParameters!");
        }

        [TestMethod()]
        public void MakeShouldThrowIfAppXmlIsNullOrEmpty()
        {
            UpdatingParameters _parameters = new UpdatingParameters();
            _parameters.remoteFolderUpdatingURL = @"server.com.ua/common";

            Should.Throw<Exception>(() => new MakerOfLinks(_parameters).Make())
            .Message.ShouldBe("Отсутствует параметр appFileXml или ссылка пустая!");
        }

        [TestMethod()]
        public void MakeShouldWriteCorrectRemoteFolderUpdatingURL()
        {
            UpdatingParameters _parameters = new UpdatingParameters();
            _parameters.remoteFolderUpdatingURL = @"server.com.ua/common";
            _parameters.appFileXml = @"app.xml";

            MakerOfLinks makerOfLinks = new MakerOfLinks(_parameters);
            makerOfLinks.Make();

            _parameters.appUpdateFolderURL.ShouldBe(@"file://server.com.ua/common/");
        }

        [TestMethod()]
        public void MakeShouldWriteCorrectAppUpdateFolderURI()
        {
            UpdatingParameters _parameters = new UpdatingParameters();
            _parameters.remoteFolderUpdatingURL = @"server.com.ua/common";
            _parameters.appFileXml = @"app.xml";

            MakerOfLinks makerOfLinks = new MakerOfLinks(_parameters);
            makerOfLinks.Make();

            _parameters.appUpdateFolderURI.ShouldBe(@"\\server.com.ua\common\");
        }
        
        [TestMethod()]
        public void MakeShouldWriteCorrectAppUpdateURL()
        {
            UpdatingParameters _parameters = new UpdatingParameters();
            _parameters.remoteFolderUpdatingURL = @"server.com.ua/common";
            _parameters.appFileXml = @"app.xml";

            MakerOfLinks makerOfLinks = new MakerOfLinks(_parameters);
            makerOfLinks.Make();

            _parameters.appUpdateURL.ShouldBe(@"file://server.com.ua/common/app.xml");
        }
    
    
    }
}