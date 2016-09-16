using Accord;
using Accord.MachineLearning.DecisionTrees;
using Accord.MachineLearning.DecisionTrees.Learning;
using Accord.MachineLearning.DecisionTrees.Rules;
using Accord.Math;
using Accord.Statistics.Filters;
using com.corticon.eclipse.studio.rule.rulesheet.core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using CorticonRules;
using org.eclipse.emf.common;
using org.eclipse.emf.common.util;
using com.corticon.eclipse.studio.rule.rulesheet.table.core;
using System.Data.Entity.Design.PluralizationServices;
using System.Globalization;
using OfficeOpenXml;
using System.IO;
using System.Reflection;

namespace CorticonRulesheetBuilder
{
    class Program
    {
        static IRulesheetTableModelAPI rulesheetTableModelAPI;
        static IRulesheetDialogAPI rulesheetDialogAPI;

        static string basePath = @"C:\Corticon\Generated\";

        static string rulesheetOutputPath = Path.Combine(basePath, @"ExampleOne.ers");

        // we're only generating the rules here. (Gotta create the vocabulary separately and link it up.)
        static string vocabularyFilePath = Path.Combine(basePath, @"Generated.ecore");

        static string dataFilePath = Path.Combine(basePath, @"Skydiver.xlsx");

        static void Main(string[] args)
        {
            #region Data
            FileInfo fileInfo = new FileInfo(dataFilePath);
            ExcelPackage pck = new ExcelPackage(fileInfo);

            var worksheet = pck.Workbook.Worksheets[1];
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;

            // create the table we're going to use to generate the rules
            var table = worksheet.ToDataTable("Applicants");

            // list the input columns
            List<string> inputColumns = new List<string>();

            for (int col = start.Column; col <= end.Column - 1; col++)
            {
                inputColumns.Add(worksheet.Cells[1, col].Text);
            }

            // list the output column
            string outputColumn = worksheet.Cells[1, end.Column].Text;
            
            var codeColumns = new List<string>();

            codeColumns.AddRange(inputColumns);
            codeColumns.Add(outputColumn);

            #endregion

            #region Decision Tree

            // convert strings to int symbols
            Codification codebook = new Codification(table, codeColumns.ToArray());

            // translate the training data into int symbols (using the codebook)
            var symbols = codebook.Apply(table);
            int[][] inputs = symbols.ToArray<int>(inputColumns.ToArray());
            int[] outputs = symbols.ToArray<int>(outputColumn);
            
            // decision variables
            List<DecisionVariable> attributes = new List<DecisionVariable>();
            for (var i = 0; i < inputColumns.Count(); i++)
            {
                attributes.Add(new DecisionVariable(inputColumns[i], codebook.Columns.First(c => c.ColumnName == inputColumns[i]).Values.Count()));
            }

            int classCount = codebook.Columns.First(c => c.ColumnName == outputColumn).Values.Count();
            
            // create decision tree
            DecisionTree tree = new DecisionTree(attributes, classCount);

            #endregion

            #region Machine Learning

            ID3Learning id3learning = new ID3Learning(tree);
            id3learning.Run(inputs, outputs);

            #endregion

            #region Convert Decision Tree to Rules

            // i wish this had been a bit more difficult to accomplish...lol
            var decisionSet = tree.ToRules();

            #endregion

            #region Corticon

            var corticonHome = ConfigurationManager.AppSettings["CORTICON_HOME"];
            var corticonWorkDir = ConfigurationManager.AppSettings["CORTICON_WORK_DIR"];
            var corticonConfiguration = new CorticonConfiguration();
            corticonConfiguration.readConfiguration(corticonHome, corticonWorkDir);

            rulesheetTableModelAPI = RulesheetTableModelAPIFactory.getInstance();
            rulesheetDialogAPI = RulesheetDialogAPIFactory.getInstance();

            var rulesheet = rulesheetDialogAPI.createRulesheet(rulesheetTableModelAPI, URI.createFileURI(rulesheetOutputPath), URI.createFileURI(vocabularyFilePath));
            
            // singularize the table name
            var entityName = PluralizationService.CreateService(CultureInfo.CurrentCulture).Singularize(table.TableName);

            // add possible conditions to rulesheet
            for (var attr = 0; attr < attributes.Count(); attr++)
            {
                rulesheetTableModelAPI.setCellValue(IRulesheetTableModelAPI.MATRIX_ID_CONDITIONS, 0, attr, String.Format("{0}.{1}", entityName, attributes[attr].Name));
            }

            // add condition values for each rule
            for (var d = 1; d <= decisionSet.Count; d++)
            {
                var rule = decisionSet.ElementAt(d - 1);
                
                for (var a = 0; a < rule.Antecedents.Count; a++)
                {
                    rulesheetTableModelAPI.setCellValue(IRulesheetTableModelAPI.MATRIX_ID_IF, d, rule.Antecedents[a].Index, String.Format("'{0}'", codebook.Translate(attributes[rule.Antecedents[a].Index].Name, (int)rule.Antecedents[a].Value)));
                }
            }

            // add action
            rulesheetTableModelAPI.setCellValue(IRulesheetTableModelAPI.MATRIX_ID_ACTIONS, 0, 0, String.Format("{0}.{1}", entityName, outputColumn));

            // add action values for each rule
            for (var d = 1; d <= decisionSet.Count; d++)
            {
                var rule = decisionSet.ElementAt(d - 1);

                rulesheetTableModelAPI.setCellValue(IRulesheetTableModelAPI.MATRIX_ID_THEN, d, 0, String.Format("'{0}'", codebook.Translate(outputColumn, (int)rule.Output)));
            }

            // save rulesheet
            rulesheetTableModelAPI.saveResource(rulesheetTableModelAPI.getPrimaryResource());
            rulesheetTableModelAPI.dispose();

            #endregion
        }
    }
}
