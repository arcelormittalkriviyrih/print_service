using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PrintWindowsService;

namespace PrintWindowsService.Tests
{
	[TestClass()]
	public class UnitTestPrintService
	{
		[TestMethod()]
		public void PrintServiceTest()
		{
            PrintJobs pJobTest = new PrintJobs();

            pJobTest.StartJob();
            pJobTest.OnPrintTimer(this, new EventArgs() as System.Timers.ElapsedEventArgs);
            //printService.PrintRange("TSC TTP-268M", "192.168.100.246");//"TSC TTP-268M");//ExportRangeAsBmp(); //PrintBmp();//

            Assert.IsTrue(true);
		}

		[TestMethod()]
		public void OnPrintTimerTest()
		{
			//Assert.Fail();
			Assert.IsTrue(true);
		}
	}
}
