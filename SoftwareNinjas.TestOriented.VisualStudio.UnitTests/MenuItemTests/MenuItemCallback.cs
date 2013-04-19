/***************************************************************************

Copyright (c) Microsoft Corporation. All rights reserved.
This code is licensed under the Visual Studio SDK license terms.
THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.

***************************************************************************/

using System.Reflection;
using System.ComponentModel.Design;
using Microsoft.VsSDK.UnitTestLibrary;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.Shell;
using SoftwareNinjas.TestOriented.VisualStudio;

namespace SoftwareNinjas.TestOriented.VisualStudio.UnitTests.MenuItemTests
{
    [TestClass]
    public class MenuItemTest
    {
        /// <summary>
        /// Verify that a new menu command object gets added to the OleMenuCommandService. 
        /// This action takes place In the Initialize method of the Package object
        /// </summary>
        [TestMethod]
        public void InitializeMenuCommand()
        {
            // Create the package
            var package = new Package() as IVsPackage;
            Assert.IsNotNull(package, "The object does not implement IVsPackage");

            // Create a basic service provider
            var serviceProvider = OleServiceProvider.CreateOleServiceProviderWithBasicServices();

            // Site the package
            Assert.AreEqual(0, package.SetSite(serviceProvider), "SetSite did not return S_OK");

            //Verify that the menu command can be found
            var menuCommandId = new CommandID(GuidList.CmdSet, (int) PkgCmdIDList.cmdidGenerateTestStub);
            var info = typeof(Microsoft.VisualStudio.Shell.Package).GetMethod("GetService", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.IsNotNull(info);
            var mcs = info.Invoke(package, new object[] { ( typeof(IMenuCommandService) ) }) as OleMenuCommandService;
            Assert.IsNotNull(mcs.FindCommand(menuCommandId));
        }

        [TestMethod]
        public void MenuItemCallback()
        {
            // Create the package
            var package = new Package() as IVsPackage;
            Assert.IsNotNull(package, "The object does not implement IVsPackage");

            // Create a basic service provider
            var serviceProvider = OleServiceProvider.CreateOleServiceProviderWithBasicServices();

            // Create a UIShell service mock and proffer the service so that it can called from the MenuItemCallback method
            var uishellMock = UIShellServiceMock.GetUiShellInstance();
            serviceProvider.AddService(typeof(SVsUIShell), uishellMock, true);

            // Site the package
            Assert.AreEqual(0, package.SetSite(serviceProvider), "SetSite did not return S_OK");

            //Invoke private method on package class and observe that the method does not throw
            var info = package.GetType().GetMethod("MenuItemCallback", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.IsNotNull(info, "Failed to get the private method MenuItemCallback throug refplection");
            info.Invoke(package, new object[] { null, null });

            //Clean up services
            serviceProvider.RemoveService(typeof(SVsUIShell));

        }
    }
}
