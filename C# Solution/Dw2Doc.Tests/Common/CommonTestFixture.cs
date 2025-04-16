using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Appeon.DotnetDemo.Dw2Doc.Tests.Common
{
    [TestClass]
    public class CommonTestFixture
    {
        [AssemblyInitialize]
        public static void AssemblyInit(TestContext context)
        {
            // Setup code that runs once for all tests in the assembly
            Console.WriteLine("Initializing Common tests...");
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            // Cleanup code that runs once after all tests in the assembly have executed
            Console.WriteLine("Cleaning up after Common tests...");
        }
    }
} 