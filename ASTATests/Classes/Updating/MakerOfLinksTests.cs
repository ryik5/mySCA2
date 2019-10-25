using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Shouldly;

namespace ASTA.Classes.Updating.Tests
{
    [TestClass()]
    public class MakerOfLinksTests
    {
        [TestMethod()]
        public void MakeShouldThrowIfServerUpdateURLIsEmpty()
        {
            UpdatingParameters _parameters = new UpdatingParameters();
            MakerOfLinks maker = new MakerOfLinks();
            maker.SetParameters(_parameters);
            Should.Throw<Exception>(() => maker.Make())
            .Message.ShouldBe("Отсутствует параметр remoteFolderUpdatingURL или ссылка пустая!");
        }
        
        [TestMethod()]
        public void MakeShouldThrowIfUpdatingParametersIsNull()
        {
            UpdatingParameters _parameters=null;

            MakerOfLinks maker = new MakerOfLinks();
            maker.SetParameters(_parameters);
            Should.Throw<Exception>(() => maker.Make())
            .Message.ShouldBe("Не создан экземпляр UpdatingParameters!");
        }

        [TestMethod()]
        public void MakeShouldThrowIfAppXmlIsNullOrEmpty()
        {
            UpdatingParameters _parameters = new UpdatingParameters();
            _parameters.remoteFolderUpdatingURL = @"server.com.ua/common";

            MakerOfLinks maker = new MakerOfLinks();
            maker.SetParameters(_parameters);
            
            Should.Throw<Exception>(() => maker.Make())
            .Message.ShouldBe("Отсутствует параметр appFileXml или ссылка пустая!");
        }

        [TestMethod()]
        public void MakeShouldWriteCorrectRemoteFolderUpdatingURL()
        {
            UpdatingParameters _parameters = new UpdatingParameters
            {
                remoteFolderUpdatingURL = @"server.com.ua/common",
                appFileXml = @"app.xml"
            };

            MakerOfLinks maker = new MakerOfLinks();
            maker.SetParameters(_parameters);
            
            maker.GetParameters().appUpdateFolderURL.ShouldBe(@"file://server.com.ua/common/");
        }

        [TestMethod()]
        public void MakeShouldWriteCorrectAppUpdateFolderURI()
        {
            UpdatingParameters _parameters = new UpdatingParameters
            {
                remoteFolderUpdatingURL = @"server.com.ua/common",
                appFileXml = @"app.xml"
            };

            MakerOfLinks maker = new MakerOfLinks();
            maker.SetParameters(_parameters);
            
            maker.GetParameters().appUpdateFolderURI.ShouldBe(@"\\server.com.ua\common\");
        }
        
        [TestMethod()]
        public void MakeShouldWriteCorrectAppUpdateURL()
        {
            UpdatingParameters _parameters = new UpdatingParameters
            {
                remoteFolderUpdatingURL = @"server.com.ua/common",
                appFileXml = @"app.xml"
            };

            MakerOfLinks maker = new MakerOfLinks();
            maker.SetParameters(_parameters);

            maker.GetParameters().appUpdateURL.ShouldBe(@"file://server.com.ua/common/app.xml");
        }    
    }
}