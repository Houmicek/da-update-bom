﻿using System;
using Inventor;
using System.IO;

namespace DebugPluginLocally
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var inv = new InventorConnector())
            {
                InventorServer server = inv.GetInventorServer();

                try
                {
                    Console.WriteLine("Running locally...");
                    // run the plugin
                    DebugSamplePlugin(server);
                }
                catch (Exception e)
                {
                    string message = $"Exception: {e.Message}";
                    if (e.InnerException != null)
                        message += $"{System.Environment.NewLine}    Inner exception: {e.InnerException.Message}";

                    Console.WriteLine(message);
                }
                finally
                {
                    if (System.Diagnostics.Debugger.IsAttached)
                    {
                        Console.WriteLine("Press any key to exit. All documents will be closed.");
                        Console.ReadKey();
                    }
                }
            }
        }

        /// <summary>
        /// Opens box.ipt and runs SamplePlugin
        /// </summary>
        /// <param name="app"></param>
        private static void DebugSamplePlugin(InventorServer app)
        {
            //// get project directory
            //string projectdir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            //// get box.ipt absolute path
            //string boxPath = System.IO.Path.Combine(projectdir, @"Files\Workspaces\Workspace\Assemblies\Suspension", "Suspension.iam");

            //string boxPathCopy = System.IO.Path.Combine(projectdir, @"Files\Workspaces\Workspace\Assemblies\Suspension", "Suspension_copy.iam");

            //try
            //{
            //    // delete an existing file
            //    System.IO.File.Delete(boxPathCopy);
            //}
            //catch (IOException)
            //{
            //    Console.WriteLine("The specified file is in use. It might be open by Inventor");
            //    return;
            //}

            //// create a copy
            //System.IO.File.Copy(boxPath, boxPathCopy);

            //// open box.ipt by Inventor
            //Document doc = app.Documents.Open(boxPathCopy);

            //// get params.json absolute path
            //string paramsPath = System.IO.Path.Combine(projectdir, @"inputFiles\", "params.json");

            //// create a name value map
            //Inventor.NameValueMap map = app.TransientObjects.CreateNameValueMap();

            //// add parameters into the map, do not change "_1". You may add more parameters "_2", "_3"...
            //map.Add("_1", paramsPath);

            // create an instance of UpdateBomPlugin
            UpdateBomPlugin.SampleAutomation plugin = new UpdateBomPlugin.SampleAutomation(app);

            /// IMPORTANT ///
            /// You need to copy your dataset to this loocation: .\UpdateBom\DebugPluginLocally\bin\Debug
            /// All Inventor documents must be closed before this action
            /// 

            // run the plugin
            plugin.Run(null, "Suspension");
        }
    }
}
