using System;
using System.IO;
using System.Linq;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VibrationVIEW_GUS;

namespace VibrationVIEW_GUS.Tests
{
    [TestClass]
    public class GUS_GetTestProfilesTests
    {
        private const string ProfilesPath = @"c:\vibrationview\profiles";

        // Unique prefix to avoid collisions with real profile files
        private const string TestPrefix = "_UNITTEST_";

        private GUS _gus;

        [TestInitialize]
        public void Setup()
        {
            _gus = new GUS();

            // Ensure the profiles directory exists
            if (!Directory.Exists(ProfilesPath))
            {
                Directory.CreateDirectory(ProfilesPath);
            }

            // Create test files for each type
            CreateTestFile("sine_test1.vsp");
            CreateTestFile("sine_test2.vsp");
            CreateTestFile("random_test1.vrp");
            CreateTestFile("shock_test1.vkp");
            CreateTestFile("datareplay_test1.vfp");
        }

        [TestCleanup]
        public void Cleanup()
        {
            // Remove only our test files
            foreach (string file in Directory.GetFiles(ProfilesPath, TestPrefix + "*"))
            {
                try { File.Delete(file); }
                catch { /* best effort cleanup */ }
            }
        }

        private void CreateTestFile(string name)
        {
            string path = Path.Combine(ProfilesPath, TestPrefix + name);
            File.WriteAllText(path, "test content");
        }

        private static string GetProfileName(XmlNode profileNode)
        {
            var nameNode = profileNode.SelectSingleNode("Name");
            return nameNode != null ? nameNode.InnerText : "";
        }

        [TestMethod]
        public void FilterSine_ReturnsVspFiles()
        {
            string result = _gus.GUS_GetTestProfiles("sine");

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);

            var testProfiles = doc.GetElementsByTagName("Profile").Cast<XmlNode>()
                .Where(n => GetProfileName(n).StartsWith(TestPrefix))
                .ToList();

            Assert.AreEqual(2, testProfiles.Count, "Should find 2 test .vsp files");
            foreach (var profile in testProfiles)
            {
                Assert.IsTrue(GetProfileName(profile).EndsWith(".vsp"),
                    "Sine filter should only return .vsp files");
            }
        }

        [TestMethod]
        public void FilterRandom_ReturnsVrpFiles()
        {
            string result = _gus.GUS_GetTestProfiles("random");

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);

            var testProfiles = doc.GetElementsByTagName("Profile").Cast<XmlNode>()
                .Where(n => GetProfileName(n).StartsWith(TestPrefix))
                .ToList();

            Assert.AreEqual(1, testProfiles.Count, "Should find 1 test .vrp file");
            Assert.IsTrue(GetProfileName(testProfiles[0]).EndsWith(".vrp"),
                "Random filter should only return .vrp files");
        }

        [TestMethod]
        public void FilterShock_ReturnsVkpFiles()
        {
            string result = _gus.GUS_GetTestProfiles("shock");

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);

            var testProfiles = doc.GetElementsByTagName("Profile").Cast<XmlNode>()
                .Where(n => GetProfileName(n).StartsWith(TestPrefix))
                .ToList();

            Assert.AreEqual(1, testProfiles.Count, "Should find 1 test .vkp file");
            Assert.IsTrue(GetProfileName(testProfiles[0]).EndsWith(".vkp"),
                "Shock filter should only return .vkp files");
        }

        [TestMethod]
        public void FilterDataReplay_ReturnsVfpFiles()
        {
            string result = _gus.GUS_GetTestProfiles("datareplay");

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);

            var testProfiles = doc.GetElementsByTagName("Profile").Cast<XmlNode>()
                .Where(n => GetProfileName(n).StartsWith(TestPrefix))
                .ToList();

            Assert.AreEqual(1, testProfiles.Count, "Should find 1 test .vfp file");
            Assert.IsTrue(GetProfileName(testProfiles[0]).EndsWith(".vfp"),
                "DataReplay filter should only return .vfp files");
        }

        [TestMethod]
        public void FilterIsCaseInsensitive()
        {
            string[] variants = { "SINE", "Sine", "sInE" };

            foreach (string variant in variants)
            {
                string result = _gus.GUS_GetTestProfiles(variant);

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(result);

                var testProfiles = doc.GetElementsByTagName("Profile").Cast<XmlNode>()
                    .Where(n => GetProfileName(n).StartsWith(TestPrefix))
                    .ToList();

                Assert.AreEqual(2, testProfiles.Count,
                    string.Format("Filter '{0}' should find 2 test .vsp files", variant));
            }
        }

        [TestMethod]
        public void DirectFileFilter_Works()
        {
            string result = _gus.GUS_GetTestProfiles("*.vsp");

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);

            var testProfiles = doc.GetElementsByTagName("Profile").Cast<XmlNode>()
                .Where(n => GetProfileName(n).StartsWith(TestPrefix))
                .ToList();

            Assert.AreEqual(2, testProfiles.Count, "Direct *.vsp filter should find 2 test files");
        }

        [TestMethod]
        public void ResultIsValidXml()
        {
            string result = _gus.GUS_GetTestProfiles("sine");

            // Should not throw
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);

            Assert.IsNotNull(doc.DocumentElement, "Result should be valid XML with a root element");
        }

        [TestMethod]
        public void XmlContainsExpectedStructure()
        {
            string result = _gus.GUS_GetTestProfiles("sine");

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);

            // Root element should be TestProfiles
            Assert.AreEqual("TestProfiles", doc.DocumentElement.Name,
                "Root element should be TestProfiles");

            // Check a test Profile element has a Name child element
            var testProfile = doc.GetElementsByTagName("Profile").Cast<XmlNode>()
                .FirstOrDefault(n => GetProfileName(n).StartsWith(TestPrefix));

            Assert.IsNotNull(testProfile, "Should have at least one test Profile element");

            var nameElement = testProfile.SelectSingleNode("Name");
            Assert.IsNotNull(nameElement, "Profile should have Name element");
            Assert.IsTrue(nameElement.InnerText.EndsWith(".vsp"),
                "Name element should contain the file name");
        }

        [TestMethod]
        public void NoMatchingFiles_ReturnsEmptyXml()
        {
            // Use a filter that won't match any files
            string result = _gus.GUS_GetTestProfiles("*.zzz_nonexistent");

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);

            Assert.AreEqual("TestProfiles", doc.DocumentElement.Name,
                "Root element should still be TestProfiles");

            XmlNodeList profiles = doc.GetElementsByTagName("Profile");
            Assert.AreEqual(0, profiles.Count,
                "Should have zero Profile elements for non-matching filter");
        }
    }
}
