/////////////////////////////////////////////////////////////////////
// Copyright (c) Autodesk, Inc. All rights reserved
// Written by Forge Partner Development
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
/////////////////////////////////////////////////////////////////////

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections.Generic;

using Autodesk.Forge.DesignAutomation.Inventor.Utils;

using Newtonsoft.Json;

using File = System.IO.File;
using Path = System.IO.Path;
using Directory = System.IO.Directory;
using Newtonsoft.Json.Linq;
using Inventor;
using System.IO;

namespace UpdateBomPlugin
{
    [ComVisible(true)]
    public class SampleAutomation
    {
        private readonly InventorServer inventorApplication;
        private string currentDirectory;

        public SampleAutomation(InventorServer inventorApp)
        {
            inventorApplication = inventorApp;
            currentDirectory = Directory.GetCurrentDirectory();
        }

        public void Run(Document doc, string documentName)
        {
            LogTrace("Run ...");
           
            //Local Debug
            NameValueMap map = inventorApplication.TransientObjects.CreateNameValueMap();
            map.Add("_1", documentName);
            RunWithArguments(doc, map);
        }

        public void RunWithArguments(Document doc, NameValueMap map)
        {
            LogTrace("RunWithArguments ...");
            
            using (new HeartBeat())
            {
                string documentName = map.Value["_1"].ToString();

                // find location of IPJ file
                string projectFile = FindFileOnCurrentDirectoryPath(documentName + ".ipj");
                LogTrace("Start processing this parameter: " + projectFile);

                if (string.IsNullOrEmpty(projectFile))
                    return;

                // we need to activate Inventor Project first to get an Assembly ready for BOM
                if (!ActivateInventorProject(projectFile))
                    return;

                // Find location of the Assembly open it
                string documentPath = FindFileOnCurrentDirectoryPath(documentName + ".iam");
                if ((doc = inventorApplication.Documents.Open(documentPath)) == null)
                    return;

                // Generate BOM with BOMViewTypeEnum
                GetBom(doc, BOMViewTypeEnum.kStructuredBOMViewType);
            }
        }

        /// <summary>
        /// Activate Inventor Project
        /// </summary>
        /// <param name="projectName"></param>
        /// <returns></returns>
        public bool ActivateInventorProject(string projectName)
        {
            try
            {
                //string currentDirectory = Directory.GetCurrentDirectory();
                string fullProjectPath = Path.GetFullPath(Path.Combine(currentDirectory, projectName));
                DesignProject project = inventorApplication.DesignProjectManager.DesignProjects.AddExisting(fullProjectPath);
                project.Activate();
            }
            catch (Exception ex)
            {
                LogTrace(ex.ToString());
                return false;
            }

            return true;
        }

        /// <summary>
        /// it create a JSON file with BOM Data
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="eBomViewType"></param>
        public void GetBom(Document doc, BOMViewTypeEnum eBomViewType)
        {
            doc.Update(); 

            try
            {
                string viewType = "Structured";
                AssemblyDocument assemblyDoc = doc as AssemblyDocument;
                AssemblyComponentDefinition componentDef = assemblyDoc.ComponentDefinition;

                LogTrace("Create BOM Object");
                BOM bom = componentDef.BOM;

                if (eBomViewType == BOMViewTypeEnum.kStructuredBOMViewType)
                {
                    bom.StructuredViewEnabled = true;
                    bom.StructuredViewFirstLevelOnly = false;
                }
                else
                {
                    bom.PartsOnlyViewEnabled = true;
                    viewType = "Parts Only";
                }

                LogTrace("Create BOM Views");
                BOMViews bomViews = bom.BOMViews;

                LogTrace("Create BOM View");
                BOMView structureView = bomViews[viewType];

                JArray bomRows = new JArray();

                LogTrace("Get BOM Rows to Json");
                BOMRowsEnumerator rows = structureView.BOMRows;
                
                LogTrace("Start to generate BOM data ...");
                GetBomRowProperties(structureView.BOMRows, bomRows);

                LogTrace("Writing out bomRows.json");
                File.WriteAllText(currentDirectory + "/bomRows.json", bomRows.ToString());
                GetListOfDirectory(currentDirectory);
            }
            catch (Exception e)
            {
                LogError("Bom failed: " + e.ToString());
            }
        }

        public void GetBomRowProperties(BOMRowsEnumerator rows, JArray bomRows)
        {
            const string TRACKING = "Design Tracking Properties";
            foreach (BOMRow row in rows)
            {
                ComponentDefinition componentDef = row.ComponentDefinitions[1];

                // Assumes not virtual component (if so add conditional for that here)
                Property partNum = componentDef.Document.PropertySets[TRACKING]["Part Number"];
                Property descr = componentDef.Document.PropertySets[TRACKING]["Description"];
                Property material = componentDef.Document.PropertySets[TRACKING]["Material"];

                JObject bomRow = new JObject(
                    new JProperty("row_number", row.ItemNumber),
                    new JProperty("part_number", partNum.Value),
                    new JProperty("quantity", row.ItemQuantity),
                    new JProperty("description", descr.Value),
                    new JProperty("material", material.Value)
                    );

                // LogTrace("Add BOM Row #" + row.ItemNumber); 
                bomRows.Add(bomRow);

                // iterate through child rows
                if (row.ChildRows != null)
                {
                    GetBomRowProperties(row.ChildRows, bomRows);
                }
            }
        }
        #region Logging utilities

        private void DirPrint(string sDir)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    foreach (string f in Directory.GetFiles(d))
                    {
                        LogTrace("file: " + f);
                    }
                    DirPrint(d);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
        }

        /// <summary>
        /// Find a file on Current Directory and in all Subdirectories
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="extension"></param>
        /// <returns></returns>
        private string FindFileOnCurrentDirectoryPath (string fileName, string extension = "*.*")
        {
            string ret = "";
            
            string[] files = Directory.GetFiles(Directory.GetCurrentDirectory(), extension, SearchOption.AllDirectories);
            foreach (string s in files)
            {
                if (s.Contains(fileName))
                    return s;
            }

            return ret;
        }

        /// <summary>
        /// List all items on a path
        /// </summary>
        /// <param name="dirPath"></param>
        private void GetListOfDirectory(string dirPath)
        {
            string[] directories = System.IO.Directory.GetDirectories(dirPath);
            string[] files = System.IO.Directory.GetFiles(dirPath);
            LogTrace(" List Of Directory : " + dirPath);
            foreach (string directory in directories)
            {
                LogTrace("Dir: " + directory);
            }

            foreach (string file in files)
            {
                LogTrace("File: " + file);
            }
            LogTrace("-----------------------------------------------------------");
        }

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string format, params object[] args)
        {
            Trace.TraceInformation(format, args);
        }

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string message)
        {
            Trace.TraceInformation(message);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string format, params object[] args)
        {
            Trace.TraceError(format, args);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string message)
        {
            Trace.TraceError(message);
        }

        #endregion
    }
}