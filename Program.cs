using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using System.Linq;

namespace GostCode
{
    class Program
    {
        /// <summary>
        /// Add file to Word document
        /// </summary>
        /// <param name="cfg">Configuration object</param>
        /// <param name="fileName">Target file with code</param>
        /// <param name="selection">Word current selection (the position where
        /// to add code)</param>
        static void addFile(Configuration cfg, String fileName,
                            Selection selection)
        {
            var shortFileName = fileName;
            if (shortFileName.StartsWith(cfg.BaseDirectory + @"\"))
                shortFileName = shortFileName.Substring(
                    cfg.BaseDirectory.Length + 1);

            Console.WriteLine("Processing '" + shortFileName + "'");

            /* Add file name */
            selection.set_Style(cfg.FileNameStyle);            
            selection.TypeText(shortFileName);

            selection.TypeParagraph();

            /* Add code */
            selection.set_Style(cfg.CodeStyle);
            var fileText = System.IO.File.ReadAllText(fileName);
            selection.TypeText(fileText);
            selection.TypeParagraph();
        }

        /// <summary>
        /// Add group to Word document
        /// </summary>
        /// <param name="cfg">Configuration</param>
        /// <param name="groupName">Group name (language-specific)</param>
        /// <param name="annotation">Group annotation</param>
        /// <param name="selection">Word current selection
        /// (where to put group)</param>
        static void addGroup(Configuration cfg, String groupName,
                             String annotation, Selection selection)
        {
            /* Add group header */
            selection.set_Style(cfg.GroupStyle);
            selection.TypeText(groupName);

            selection.TypeParagraph();

            /* Add group annotation */
            selection.set_Style(cfg.AnnotationStyle);
            selection.TypeText(annotation);
        }

        /// <summary>
        /// Process group completely
        /// </summary>
        /// <param name="cfg">Configuration object</param>
        /// <param name="group">Target group object</param>
        /// <param name="selection">Word current selection
        /// (where to put group and code)</param>
        static void processGroup(Configuration cfg, Group group,
                                 Selection selection)
        {
            var allFiles = new List<string>();

            Console.WriteLine("Processing group '" + group.Name + "'...");
            foreach (var folder in group.Folders)
            {
                foreach (var filter in group.Filters)
                {
                    var files = System.IO.Directory.GetFiles(
                        System.IO.Path.Combine(cfg.BaseDirectory, folder),
                        filter,
                        group.Recursive
                            ? System.IO.SearchOption.AllDirectories
                            : System.IO.SearchOption.TopDirectoryOnly);
                    allFiles.AddRange(files);
                }
            }

            /* We got many files, let's remove duplicates and sort the rest */
            allFiles = allFiles.Distinct().ToList();
            allFiles.Sort();

            addGroup(cfg, group.Name, group.Annotation, selection);
            foreach (var file in allFiles)
            {
                addFile(cfg, file, selection);
            }
            Console.WriteLine("Processing group '" + group.Name + "' ok");
        }

        /// <summary>
        /// Display usage information
        /// </summary>
        static void Usage()
        {
            Console.WriteLine(
                "GostCode.exe template_file output_file baseDirectory " +
                "fileList");
        }

        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Word.Application w = null;

            if (args.Length != 4)
            {
                Usage();
                return;
            }

            /* Retrieve arguments */
            var template = args[0];
            var outFile = args[1];
            var baseDirectory = args[2];
            var list = args[3];

            try
            {
                /* Create Word instance */
                w = new Microsoft.Office.Interop.Word.Application();

                /* Check files existence */
                if (!System.IO.File.Exists(template))
                {
                    throw new System.IO.FileNotFoundException(
                        "Template is not found: " + template);
                }

                if (!System.IO.File.Exists(list))
                {
                    throw new System.IO.FileNotFoundException(
                        "File list is not found: " + list);
                }

                /* Read file list (Configuration) */
                Console.WriteLine("Read file list '" + list + "'");
                var convention = new UnderscoredNamingConvention();
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(convention)
                    .Build();

                /*
                 * Read base directory from configuration. If it is not
                 * available, use the one from parameters
                 */
                var cfg = deserializer.Deserialize<Configuration>(
                    System.IO.File.ReadAllText(list));
                if (cfg.BaseDirectory == "")
                {
                    cfg.BaseDirectory = baseDirectory;
                }

                if (!System.IO.Directory.Exists(cfg.BaseDirectory))
                {
                    throw new System.IO.DirectoryNotFoundException(
                        "Base directory is not found: '" +
                        cfg.BaseDirectory + "'");
                }

                /* Reduce the load on Word */
                w.Options.CheckGrammarWithSpelling = false;
                w.Options.CheckSpellingAsYouType = false;
                w.Options.CheckGrammarAsYouType = false;
                w.Options.AllowFastSave = true;
                w.Documents.Open(System.IO.Path.GetFullPath(template));
                var doc = w.ActiveDocument;
                
                var range = doc.Range();
                if (range.Find.Execute("<Insert>"))
                {
                    w.Selection.SetRange(range.Start, range.End);
                    foreach (var group in cfg.Groups)
                    {
                        processGroup(cfg, group, w.Selection);
                        /*
                         * Clearing undo cache is FAIRLY important to let Word
                         * cope with very large files
                         */
                        doc.UndoClear();
                    }
                }
                else
                {
                    throw new Exception("<Insert> is not found");
                }
                
                Console.WriteLine("Updating ToC...");
                for (int i = 0; i < doc.TablesOfContents.Count; ++i)
                {
                    doc.TablesOfContents[1 + i].Update();
                }

                Console.WriteLine("Updating Page Count...");
                range = doc.Range();
                if (range.Find.Execute("<ListCount>"))
                {
                    range.Text = doc.ComputeStatistics(
                        WdStatistic.wdStatisticPages).ToString();
                    range.Select();
                }
                
                Console.WriteLine("Writing to '" +
                    System.IO.Path.GetFullPath(outFile) + "'...");
                w.ActiveDocument.SaveAs2(System.IO.Path.GetFullPath(outFile));
                Console.WriteLine("Done");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.ToString());
            }
            finally
            {
                if (w != null && w.Documents.Count != 0)
                    w.ActiveDocument.Close(false);
            }
        }
    }
}
