using System.Collections.Generic;
using Autodesk.Forge.DesignAutomation.Model;

namespace Interaction
{
    /// <summary>
    /// Customizable part of Publisher class.
    /// </summary>
    internal partial class Publisher
    {
        /// <summary>
        /// Constants.
        /// </summary>
        private static class Constants
        {
            private const int EngineVersion = 24;
            public static readonly string Engine = $"Autodesk.Inventor+{EngineVersion}";

            public const string Description = "Outputs a BOM json file";

            internal static class Bundle
            {
                public static readonly string Id = "UpdateBom";
                public const string Label = "alpha";

                public static readonly AppBundle appBundle = new AppBundle
                {
                    Engine = Engine,
                    Id = Id,
                    Description = Description
                };
            }

            internal static class Activity
            {
                public static readonly string Id = Bundle.Id;
                public const string Label = Bundle.Label;
            }

            internal static class Parameters
            {
                public const string inputFile = nameof(inputFile);
                public const string inputParams = nameof(inputParams);
                public const string outputFile = nameof(outputFile);
            }
        }

        /// <summary>
        /// Get command line for activity. 
        /// Parameter is the name of IPJ file and must be the same as Assembly Name
        /// </summary>
        private static List<string> GetActivityCommandLine()
        {
            return new List<string> { $"$(engine.path)\\InventorCoreConsole.exe /al $(appbundles[{Constants.Bundle.Id}].path) Suspension" };
        }

        /// <summary>
        /// Get activity parameters.
        /// </summary>
        private static Dictionary<string, Parameter> GetActivityParams()
        {
            return new Dictionary<string, Parameter>
                    {
                         {
                            Constants.Parameters.inputFile,
                            new Parameter
                            {
                                Verb = Verb.Get,
                                Description = "Input assembly to extract parameters",
                            }
                        },
                       {
                            Constants.Parameters.outputFile,
                            new Parameter
                            {
                                Verb = Verb.Put,
                                LocalName = "bomRows.json",
                                Description = "Resulting BOM Rows data",
                            }
                        }
                    };
        }

        /// <summary>
        /// Get arguments for workitem.
        /// </summary>
        private static Dictionary<string, IArgument> GetWorkItemArgs()
        {
            // TODO: update the URLs below with real values
            return new Dictionary<string, IArgument>
                    {
                    {
                        Constants.Parameters.inputFile,
                        new XrefTreeArgument
                        {
                            Url = "https://inventor-io-samples.s3.us-west-2.amazonaws.com/holecep/bom/Suspension.zip?AWSAccessKeyId=AKIAINFJUJXZQ3REAW2A&Expires=1586530080&Signature=c3IkIyq4XnL1D0QZeVcit2thhfE%3D",
                            PathInZip = "Workspaces/Workspace/Assemblies/Suspension/Suspension.iam",
                            LocalName = "Asm"

                        }
                    },
                    {
                        Constants.Parameters.outputFile,
                        new XrefTreeArgument
                        {
                            Verb = Verb.Put,
                            Url = "https://inventor-io-samples.s3.us-west-2.amazonaws.com/holecep/bom/bomRows.json?AWSAccessKeyId=AKIAINFJUJXZQ3REAW2A&Expires=1582975080&Signature=wTGEsiFEsmAbQLRHXzwg2SkC6gs%3D"
                        }
                    }
            };
        }
    }
}
