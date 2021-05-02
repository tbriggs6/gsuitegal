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
    public class testListMerger
    {
        [TestMethod]
        public void TestGoogleOnly()
        {
            ListMerger merger = new ListMerger();
            merger.addGoogleAddress("a", "a");

            var changes = merger.processLists();
            Assert.AreEqual(1, changes.toadd.Count);
            Assert.AreEqual(0, changes.toremove.Count);
            Assert.AreEqual(0, changes.tochange.Count);
        }

        [TestMethod]
        public void TestOutlookOnly()
        {
            ListMerger merger = new ListMerger();
            merger.addOutlookAddress("a", "a");

            var changes = merger.processLists();
            Assert.AreEqual(0, changes.toadd.Count);
            Assert.AreEqual(1, changes.toremove.Count);
            Assert.AreEqual(0, changes.tochange.Count);
        }

        [TestMethod]
        public void TestBoth()
        {
            ListMerger merger = new ListMerger();
            merger.addGoogleAddress("a", "a");
            merger.addOutlookAddress("a", "a");
            var changes = merger.processLists();
            Assert.AreEqual(0, changes.toadd.Count);
            Assert.AreEqual(0, changes.toremove.Count);
            Assert.AreEqual(0, changes.tochange.Count);
        }

        [TestMethod]
        public void TestChange()
        {
            ListMerger merger = new ListMerger();
            merger.addGoogleAddress("a", "a");
            merger.addOutlookAddress("a", "b");

            var changes = merger.processLists();
            Assert.AreEqual(0, changes.toadd.Count);
            Assert.AreEqual(0, changes.toremove.Count);
            Assert.AreEqual(1, changes.tochange.Count);
        }
    }
}
