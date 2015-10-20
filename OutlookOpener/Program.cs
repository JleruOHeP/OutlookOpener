using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOpener
{
	class Program
	{
		static void Main(string[] args)
		{
			Outlook.Application app;
			if (Process.GetProcessesByName("OUTLOOK").Any())
				app = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
			else
				app = new Outlook.Application();
			
			Outlook.MailItem mailItem = app.CreateItemFromTemplate(Path.Combine(Directory.GetCurrentDirectory(), "template.oft"));

			//it is throwing on the next line.
			//how to fix it?
			//and to confirm - I`m using office 2010
			var body = mailItem.HTMLBody;
			mailItem.HTMLBody = body.Replace("@firstname", "Test Testy");

			mailItem.Display();

			Console.ReadKey();
		}
	}
}
