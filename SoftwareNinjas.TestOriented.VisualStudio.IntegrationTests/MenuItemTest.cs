using System.Globalization;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VsSDK.IntegrationTestLibrary;
using Microsoft.VSSDK.Tools.VsIdeTesting;

namespace SoftwareNinjas.TestOriented.VisualStudio.IntegrationTests
{
    [TestClass]
    public class MenuItemTest
    {
        private delegate void ThreadInvoker();

        /// <summary>
        ///A test for lauching the command and closing the associated dialogbox
        ///</summary>
        [TestMethod]
        [HostType("VS IDE")]
        public void LaunchCommand()
        {
            UIThreadInvoker.Invoke((ThreadInvoker) delegate
            {
                var menuItemCmd = new CommandID(SoftwareNinjas_TestOriented_VisualStudio.GuidList.guidSoftwareNinjas_TestOriented_VisualStudioCmdSet, (int) SoftwareNinjas_TestOriented_VisualStudio.PkgCmdIDList.cmdidGenerateTestStub);

                // Create the DialogBoxListener Thread.
                var expectedDialogBoxText = string.Format(CultureInfo.CurrentCulture, "{0}\n\nInside {1}.MenuItemCallback()", "SoftwareNinjas.TestOriented.VisualStudio", "SoftwareNinjas.SoftwareNinjas_TestOriented_VisualStudio.SoftwareNinjas_TestOriented_VisualStudioPackage");
                var purger = new DialogBoxPurger(NativeMethods.IDOK, expectedDialogBoxText);
                
                try
                {
                    purger.Start();

                    var testUtils = new TestUtils();
                    testUtils.ExecuteCommand(menuItemCmd);
                }
                finally
                {
                    Assert.IsTrue(purger.WaitForDialogThreadToTerminate(), "The dialog box has not shown");
                }
            });
        }

    }
}
