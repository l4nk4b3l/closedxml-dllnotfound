using System;
using System.Diagnostics;
using System.IO;
using ClosedXML.Excel;
using Foundation;
using UIKit;

namespace DllNotFoundApp
{
    // The UIApplicationDelegate for the application. This class is responsible for launching the
    // User Interface of the application, as well as listening (and optionally responding) to application events from iOS.
    [Register("AppDelegate")]
    public class AppDelegate : UIApplicationDelegate
    {
        // class-level declarations

        public override UIWindow Window { get; set; }

        public override bool FinishedLaunching(UIApplication application, NSDictionary launchOptions)
        {
            // create a new window instance based on the screen size
            Window = new UIWindow(UIScreen.MainScreen.Bounds);
            Window.RootViewController = new UIViewController();

            // make the window visible
            Window.MakeKeyAndVisible();
            
            // Creating an Excel 
            var documents = Environment.GetFolderPath (Environment.SpecialFolder.MyDocuments);
            string path = Path.Combine(documents, "HelloWorld.xlsx");
                
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell("A1").Value = "Hello World!";
                // Uncomment this line for crash
                // worksheet.Column(1).AdjustToContents();
                worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                workbook.SaveAs(path);
            }

            var loadOptions = new LoadOptions();
            // Workaround: This makes it possible to get the data but probably breaks the size calculation
            loadOptions.GraphicEngine = new MyGraphicsEngine();
            
            // Reading the Excel crashes since >= 101.0.0
            using (var workbook = new XLWorkbook(File.OpenRead(path), loadOptions))
            {
                Debug.WriteLine(workbook.Worksheet(1).Cell("A1").Value.GetText());
            }

            return true;
        }

        public override void OnResignActivation(UIApplication application)
        {
            // Invoked when the application is about to move from active to inactive state.
            // This can occur for certain types of temporary interruptions (such as an incoming phone call or SMS message) 
            // or when the user quits the application and it begins the transition to the background state.
            // Games should use this method to pause the game.
        }

        public override void DidEnterBackground(UIApplication application)
        {
            // Use this method to release shared resources, save user data, invalidate timers and store the application state.
            // If your application supports background exection this method is called instead of WillTerminate when the user quits.
        }

        public override void WillEnterForeground(UIApplication application)
        {
            // Called as part of the transiton from background to active state.
            // Here you can undo many of the changes made on entering the background.
        }

        public override void OnActivated(UIApplication application)
        {
            // Restart any tasks that were paused (or not yet started) while the application was inactive. 
            // If the application was previously in the background, optionally refresh the user interface.
        }

        public override void WillTerminate(UIApplication application)
        {
            // Called when the application is about to terminate. Save data, if needed. See also DidEnterBackground.
        }
    }
}