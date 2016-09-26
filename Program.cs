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

        // base path to the assets folder
        static string basePath = @"C:\Corticon\Generated\";
        static string rulesheetOutputPath = Path.Combine(basePath, @"ExampleOne.ers");
        // we're only generating the rules here. (Gotta create the vocabulary separately and link it up.)
        static string vocabularyFilePath = Path.Combine(basePath, @"Generated.ecore");
        static string dataFilePath = Path.Combine(basePath, @"Applicants.xlsx");

        static void Main(string[] args)
        {
            #region Data
            var fileInfo = new FileInfo(dataFilePath);
            var pck = new ExcelPackage(fileInfo);

            var worksheet = pck.Workbook.Worksheets[1];
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;

            // create the table we're going to use to generate the rules
            var table = worksheet.ToDataTable(Path.GetFileNameWithoutExtension(dataFilePath));

            // list the input columns
            var inputColumns = new List<string>();

            for (int col = start.Column; col <= end.Column - 1; col++)
            {
                inputColumns.Add(worksheet.Cells[1, col].Text);
            }

            // list the output column
            var outputColumn = worksheet.Cells[1, end.Column].Text;
            
            var codeColumns = new List<string>();

            codeColumns.AddRange(inputColumns);
            codeColumns.Add(outputColumn);

            #endregion

            #region Decision Tree

            // convert strings to int symbols
            var codebook = new Codification(table, codeColumns.ToArray());

            // translate the training data into int symbols (using the codebook)
            var symbols = codebook.Apply(table);
            var inputs = symbols.ToArray<double>(inputColumns.ToArray());
            var outputs = symbols.ToArray<int>(outputColumn);

            // decision variables
            var attributes = new List<DecisionVariable>();
            for (var i = 0; i < inputColumns.Count(); i++)
            {
                attributes.Add(new DecisionVariable(inputColumns[i], codebook.Columns.First(c => c.ColumnName == inputColumns[i]).Values.Count()));
            }

            var classCount = codebook.Columns.First(c => c.ColumnName == outputColumn).Values.Count();

            // create decision tree
            var tree = new DecisionTree(attributes, classCount);

            #endregion

            #region Machine Learning
            
            // ID3 can be used here to increase accuracy against training data.
            // I'm using C4.5 to avoid overfitting the training data and to improve accuracy against unseen data.

            // create the C4.5 algorithm
            var c45 = new C45Learning(tree);
            c45.Run(inputs, outputs);

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
