using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VSSDK.Tools.VsIdeTesting;
using Microsoft.VisualStudio.Shell.Interop;

namespace SoftwareNinjas.TestOriented.VisualStudio.IntegrationTests
{
    /// <summary>
    /// Integration test for package validation
    /// </summary>
    [TestClass]
    public class PackageTest
    {
        private delegate void ThreadInvoker();

        [TestMethod]
        [HostType("VS IDE")]
        public void PackageLoadTest()
        {
            UIThreadInvoker.Invoke((ThreadInvoker) delegate
            {
                //Get the Shell Service
                var shellService = VsIdeTestHostContext.ServiceProvider.GetService(typeof(SVsShell)) as IVsShell;
                Assert.IsNotNull(shellService);

                //Validate package load
                IVsPackage package;
                var packageGuid = new Guid(SoftwareNinjas_TestOriented_VisualStudio.GuidList.guidSoftwareNinjas_TestOriented_VisualStudioPkgString);
                Assert.IsTrue(0 == shellService.LoadPackage(ref packageGuid, out package));
                Assert.IsNotNull(package, "Package failed to load");

            });
        }
    }
}
