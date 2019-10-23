using Microsoft.VisualStudio.TestTools.UnitTesting;
using ASTA.Classes.AutoUpdating;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Shouldly;

namespace ASTA.Classes.AutoUpdating.Tests
{
    [TestClass()]
    public class AccountEventArgsTests
    {
        [TestMethod()]
        public void AccountEventArgsShouldRememberCorrectMessageIfGotMessage()
        {
            AccountEventArgs accountEventArgs = new AccountEventArgs("test");
            accountEventArgs.Message.ShouldBe("test");
        }

        [TestMethod()]
        public void AccountEventArgsShouldNotBeEmptyIfGotMessage()
        {
            AccountEventArgs accountEventArgs = new AccountEventArgs("test");
            accountEventArgs.Message.ShouldNotBeEmpty();
        }
    }
}