using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation.Language;
using Microsoft.Windows.PowerShell.ScriptAnalyzer.Generic;
using System.ComponentModel.Composition;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Text;
using System.Management.Automation.Runspaces;
using System.Management.Automation;
using System.Collections.ObjectModel;

namespace Microsoft.Windows.PowerShell.ScriptAnalyzer.BuiltinRules
{
    /// <summary>
    /// MissingParameterDocumentation: Checks for functions that have parameters that don't have any comment documentation
    /// </summary>
    [Export(typeof(IScriptRule))]
    public class MissingParameterDocumentation : IScriptRule
    {
        /// <summary>
        /// MissingParameterDocumentation: Checks for functions that have parameters that don't have any comment documentation
        /// </summary>
        public IEnumerable<DiagnosticRecord> AnalyzeScript(Ast ast, string fileName)
        {
            if (ast == null) throw new ArgumentNullException(Strings.NullAstErrorMessage);

            List<DiagnosticRecord> diagnosticRecords = new List<DiagnosticRecord>();

            IEnumerable<Ast> foundAsts = ast.FindAll(testAst => testAst is FunctionDefinitionAst, true);

            // Iterates all FunctionDefinitionAst and check the help vs parameters name.
            foreach (FunctionDefinitionAst functionAst in foundAsts)
            {
                CommentHelpInfo commentHelp = functionAst.GetHelpContent();
                ICollection<string> parametersInHelp = commentHelp.Parameters.Keys;

                List<string> parametersInFunction = new List<string>();

                IEnumerable<Ast> paramBlockAsts = functionAst.FindAll(testAst => testAst is ParamBlockAst, true);
                foreach (ParamBlockAst paramBlockAst in paramBlockAsts)
                { 
                    foreach (ParameterAst paramAst in paramBlockAst.Parameters)
                    {
                         parametersInFunction.Add(paramAst.Name.VariablePath.UserPath);
                    }
                }

                IEnumerable<Ast> dynamicParamAsts = functionAst.FindAll(testAst => testAst is NamedBlockAst && ((NamedBlockAst)testAst).BlockKind == TokenKind.Dynamicparam, true);

                foreach (NamedBlockAst dynamicParamAst in dynamicParamAsts)
                {
                    StringBuilder script = new StringBuilder();
                    foreach (StatementAst statement in dynamicParamAst.Statements)
                    {
                        script.AppendLine(statement.ToString());
                    }
                    Runspace runSpace = RunspaceFactory.CreateRunspace();
                    runSpace.Open();
                    Pipeline pipeline = runSpace.CreatePipeline();

                    Command invokeScript = new Command("Invoke-Command");
                    RunspaceInvoke invoke = new RunspaceInvoke();
                    ScriptBlock sb = invoke.Invoke(String.Format("{{ $rtd = .{{ {0} }} ;  $rtd[0].Keys | Write-Output }}", script.ToString()))[0].BaseObject as ScriptBlock;
                    invokeScript.Parameters.Add("scriptBlock", sb);

                    pipeline.Commands.Add(invokeScript);
                    Collection<PSObject> output = pipeline.Invoke();
                    foreach (PSObject psObject in output)
                    {
                        parametersInFunction.Add(psObject.ToString());
                    }
                }

                foreach (string functionParam in parametersInFunction)
                {
                    if (parametersInHelp.Contains(functionParam) == false)
                    {
                        diagnosticRecords.Add(new DiagnosticRecord(
                            string.Format(CultureInfo.CurrentCulture, Strings.MissingParameterDocumentationError, functionParam),
                            functionAst.Extent,
                            GetName(),
                            DiagnosticSeverity.Warning,
                            fileName,
                            functionAst.Name,
                            suggestedCorrections: GetCorrectionExtent(functionAst.Extent, functionParam)));
                    }
                }

            }

            return diagnosticRecords;
        }



        /// <summary>
        /// Creates a list containing suggested correction
        /// </summary>
        /// <param name="violationExtent"></param>
        /// <param name="paramName"></param>
        /// <returns>Returns a list of suggested corrections</returns>
        private List<CorrectionExtent> GetCorrectionExtent(IScriptExtent violationExtent, string paramName)
        {
            var corrections = new List<CorrectionExtent>();
            string description = "Add parameter documentation";
            corrections.Add(new CorrectionExtent(
                violationExtent.StartLineNumber,
                violationExtent.EndLineNumber,
                violationExtent.StartColumnNumber + 1,
                violationExtent.EndColumnNumber,
                String.Empty,
                violationExtent.File,
                description));
            return corrections;
        }

        /// <summary>
        /// GetName: Retrieves the name of this rule.
        /// </summary>
        /// <returns>The name of this rule</returns>
        public string GetName()
        {
            return string.Format(CultureInfo.CurrentCulture, Strings.NameSpaceFormat, GetSourceName(), Strings.MissingParameterDocumentationName);
        }

        /// <summary>
        /// GetCommonName: Retrieves the common name of this rule.
        /// </summary>
        /// <returns>The common name of this rule</returns>
        public string GetCommonName()
        {
            return string.Format(CultureInfo.CurrentCulture, Strings.MissingParameterDocumentationCommonName);
        }

        /// <summary>
        /// GetDescription: Retrieves the description of this rule.
        /// </summary>
        /// <returns>The description of this rule</returns>
        public string GetDescription()
        {
            return string.Format(CultureInfo.CurrentCulture, Strings.MissingParameterDocumentationDescription);
        }

        /// <summary>
        /// GetSourceType: Retrieves the type of the rule: builtin, managed or module.
        /// </summary>
        public SourceType GetSourceType()
        {
            return SourceType.Builtin;
        }

        /// <summary>
        /// GetSeverity: Retrieves the severity of the rule: error, warning of information.
        /// </summary>
        /// <returns></returns>
        public RuleSeverity GetSeverity()
        {
            return RuleSeverity.Warning;
        }

        /// <summary>
        /// GetSourceName: Retrieves the module/assembly name the rule is from.
        /// </summary>
        public string GetSourceName()
        {
            return string.Format(CultureInfo.CurrentCulture, Strings.SourceName);
        }
    }
}
