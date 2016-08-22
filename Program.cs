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

namespace CorticonRulesheetBuilder
{
    class Program
    {
        static IRulesheetTableModelAPI rulesheetTableModelAPI;
        static IRulesheetDialogAPI rulesheetDialogAPI;

        static string rulesheetOutputPath = @"C:\Corticon\ExampleOne.ers";

        // we're only generating the rules here. (Gotta create the vocabulary seperately and link it up.)
        static string vocabularyFilePath = @"C:\Corticon\Generated.ecore";

        static void Main(string[] args)
        {
            #region Data

            // create the data we're going to use to generate the rules
            DataTable table = new DataTable("Applicants");

            table.Columns.Add("Age", "Skydiver", "Weight", "Gender", "Risk");

            table.Rows.Add("young", "yes", "heavy", "male", "high");
            table.Rows.Add("young", "yes", "light", "female", "high");
            table.Rows.Add("old", "yes", "heavy", "male", "high");
            table.Rows.Add("old", "yes", "light", "female", "high");
            table.Rows.Add("young", "no", "light", "female", "low");
            table.Rows.Add("young", "no", "heavy", "female", "medium");
            table.Rows.Add("young", "no", "heavy", "male", "medium");
            table.Rows.Add("old", "no", "light", "male", "medium");
            table.Rows.Add("old", "no", "heavy", "female", "medium");

            // list the input columns
            List<string> inputColumns = new List<string>()
            {
                "Age",
                "Skydiver",
                "Weight",
                "Gender"
            };


            // list the output column
            string outputColumn = "Risk";
            
            var codeColumns = new List<string>();

            codeColumns.AddRange(inputColumns);
            codeColumns.Add(outputColumn);

            #endregion

            #region Decision Tree

            // convert strings to int symbols
            Codification codebook = new Codification(table, codeColumns.ToArray());
            
            // translate the training data into int symbols (using the codebook)
            DataTable symbols = codebook.Apply(table);
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
