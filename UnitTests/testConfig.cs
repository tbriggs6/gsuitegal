using Microsoft.VisualStudio.TestTools.UnitTesting;
using GSuiteGAL;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace UnitTests
{
    [TestClass]
    public class testConfig
    {
        [TestMethod]
        public void TestFlatConfig()
        {
            Registry.CurrentUser.DeleteSubKey("Software\\GoogleGAL", false);
            
            // this should not throw any errors
            Config config = new Config();
            Assert.AreEqual(0,config.lastSync);
            Assert.AreEqual(86400, config.syncPeriod);
            Assert.AreEqual("credentials.json", config.credentialFileName);
            Assert.AreEqual("tokens.json", config.tokenFileName);

        }

        [TestMethod]
        public void TetRestoredConfig()
        {

            Config config = new Config();
            Assert.AreEqual(100, config.lastSync);
            Assert.AreEqual(86400, config.syncPeriod);
            Assert.AreEqual("credentials.json", config.credentialFileName);
            Assert.AreEqual("tokens.json", config.tokenFileName);

        }


        [TestMethod]
        public void TestSetLastSyncTime()
        {

            Config config = new Config();
            Assert.AreEqual(0, config.lastSync);
            config.SetLastSync(100);
            Assert.AreEqual(100, config.lastSync);

            // reopen the registry and read it again
            Config config2 = new Config();
            Assert.AreEqual(100, config2.lastSync);

        }

    }
}
