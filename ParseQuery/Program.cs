using Microsoft.SqlServer.TransactSql.ScriptDom;
using System.Diagnostics;
using System.Text;

namespace Sample.ParseQuery
{
    class Program
    {
        static int Main(string[] args)
        {
            #if DEBUG
            TextWriterTraceListener myWriter = new TextWriterTraceListener(System.Console.Out);
            Trace.Listeners.Add(myWriter);
            Trace.Indent();
            #endif

            string filePath = "sample.sql";
            int outputTypeInt = 1;

            using (StreamReader rdr = new StreamReader(filePath))
            {
                // parse file with scriptdom
                IList<ParseError> errors = new List<ParseError>();
                var parser = new TSql160Parser(true, SqlEngineType.All);
                TSqlFragment tree = parser.Parse(rdr, out errors);
                // https://learn.microsoft.com/en-us/dotnet/api/microsoft.sqlserver.transactsql.scriptdom.tsqlfragment?view=sql-transactsql-161

                if (errors.Count == 0)
                {
                    Console.WriteLine("✅ Script parsed successfully, no errors found!");
                }

                foreach (ParseError err in errors)
                {
                    // print errors
                    Console.Error.WriteLine(err.Message);
                    return -1;
                }

                // parse the script for variables (@p0001, @p0002, etc.)
                var variableFinder = new VariableFinderVisitor();
                tree.Accept(variableFinder);
                if (variableFinder.Variables.Any()) {
                    foreach (VariableContents vc in variableFinder.Variables) {
                        Trace.WriteLine($"ℹ️  Variable {vc.VariableName} contains \"{vc.Contents.Script()}\" ");
                    }
                }

                var spExecuteFinder = new spExecuteVisitor();
                tree.Accept(spExecuteFinder);
                if (spExecuteFinder.ExecuteStatements.Any()) {
                    foreach (ExecuteSpecification execSpec in spExecuteFinder.ExecuteStatements) {
                        // output the sp_executesql statement
                        Console.WriteLine("✏️  sp_executesql statement:");
                        Console.WriteLine($"   {execSpec.Script()}");

                        // get the first variable from the exec sp_executesql statement
                        Trace.WriteLine($"ℹ️  The 1st param for sp_executesql is {execSpec.GetFirstParameter()}");
                    }
                }

                string baseQueryToExecute = "";
                foreach (VariableContents vc in variableFinder.Variables) {
                    // pick just the first call to sp_executesql
                    if (vc.VariableName == spExecuteFinder.ExecuteStatements.First().GetFirstParameter())
                    {
                        baseQueryToExecute = vc.Contents.ScriptStringValue();
                    }
                }

                if (outputTypeInt == 1)
                {
                    // output the initial sp_executesql query
                    Console.WriteLine("✏️  Initial sp_executesql query:");
                    Console.WriteLine($"   {baseQueryToExecute}");
                    
                    return 0;
                }
                else if (outputTypeInt == 2)
                {
                    List<(string, string)> pairs = spExecuteFinder.ExecuteStatements.First().GetParameterPairs();

                    // we need to parse the initial sp_executesql query by itself
                    // to get the variable names and and set their values
                    using (TextReader tr = new StringReader(baseQueryToExecute))
                    {
                        TSqlFragment initialQuery = parser.Parse(tr, out errors);

                        var replacementNeededVisitor = new VariableReferenceVisitor();
                        initialQuery.Accept(replacementNeededVisitor);
                        string newQuery = "";
                        int i = 0;

                        List<VariableReference> variablesToReplace = replacementNeededVisitor.VariablesToReplace;
                        // sort by startoffset
                        variablesToReplace.Sort((a, b) => a.StartOffset.CompareTo(b.StartOffset));

                        // get the variables that need to be replaced
                        foreach (VariableReference vr in variablesToReplace)
                        {
                            newQuery += baseQueryToExecute.Substring(i, vr.StartOffset - i);
                            i = newQuery.Length;

                            // find the value to replace the variable with
                            string lookupVariable = pairs.Find(x => x.Item1 == vr.Name).Item2;
                            Trace.WriteLine($"ℹ️  Variable {vr.Name} needs to be replaced with {lookupVariable}");
                            if (lookupVariable == null)
                            {
                                Console.Error.WriteLine($"Variable {vr.Name} not found in sp_executesql statement.");
                                return -1;
                            }
                            VariableContents variableContents = variableFinder.Variables.Find(x => x.VariableName == lookupVariable) ?? throw new Exception($"Variable {lookupVariable} not found.");
                            string newValue = variableContents.Contents.Script();
                            newQuery += newValue;

                            i += vr.FragmentLength;
                        }

                        // output the substituted query
                        Console.WriteLine("✏️  Substituted query:");
                        Console.WriteLine($"   {newQuery}");
                    }

                    return 0;
                }
                else
                {
                    Console.Error.WriteLine("Invalid output type.");
                    return -1;
                }
            }
        }
    }



    // structure to hold a variable by name and its contents for easier handling
    class VariableContents
    {
        public string VariableName { get; set; }
        public ScalarExpression Contents { get; set; }

        public VariableContents(string variableName, ScalarExpression contents)
        {
            VariableName = variableName;
            Contents = contents;
        }
    }

    // visitor to fragments that creates a list of variables (eg @p0001) and their contents
    class VariableFinderVisitor : TSqlFragmentVisitor
    {
        public List<VariableContents> Variables { get; } = new List<VariableContents>();

        // Grab "SET @VariableName = ..." nodes
        public override void Visit(SetVariableStatement node)
        {
            VariableContents vc = new VariableContents(node.Variable.Name, node.Expression);
            Variables.Add(vc);
        }
    }

    class spExecuteVisitor : TSqlFragmentVisitor
    {
        public List<ExecuteSpecification> ExecuteStatements { get; } = new List<ExecuteSpecification>();

        // grab the exec sp_executesql statements
        public override void Visit(ExecuteStatement node)
        {
            if (node.ExecuteSpecification.ExecutableEntity is ExecutableProcedureReference procRef)
            {
                if (procRef.ProcedureReference.ProcedureReference.Name.BaseIdentifier.Value.Equals("sp_executesql", StringComparison.OrdinalIgnoreCase))
                {
                    ExecuteStatements.Add(node.ExecuteSpecification);
                }
            }
        }
    }

    // visitor to variable references that returns a list of variables so that they can be replaced
    // finds @LastUsed and @sequence in "UPDATE session SET ss_last_used = @LastUsed WHERE ss_sequence = @sequence"
    class VariableReferenceVisitor : TSqlFragmentVisitor
    {
        public List<VariableReference> VariablesToReplace { get; } = new List<VariableReference>();

        public override void Visit(VariableReference node)
        {
            VariablesToReplace.Add(node);
        }
    }


    // utility class to extend the functionality of TSqlFragment
    public static class TSqlFragmentExtensions
    {
        // write out the contents of a fragment
        public static string Script(this TSqlFragment fragment)
        {
            return String.Join("", fragment.ScriptTokenStream
                .Skip(fragment.FirstTokenIndex)
                .Take(fragment.LastTokenIndex - fragment.FirstTokenIndex + 1)
                .Select(t => t.Text)
            );
        }

        // write out the contents of a scalar expression
        // removing the leading and trailing apostrophes and N prefix
        // from  N'UPDATE session SET ss_last_used = @LastUsed WHERE ss_sequence = @sequence'
        // to  UPDATE session SET ss_last_used = @LastUsed WHERE ss_sequence = @sequence
        public static string ScriptStringValue(this TSqlFragment fragment)
        {
            string script = Script(fragment);
            if (script.StartsWith("N'"))
            {
                script = script.Substring(2);
            }
            else if (script.StartsWith("'"))
            {
                script = script.Substring(1);
            }
            if (script.EndsWith("'"))
            {
                script = script.Substring(0, script.Length - 1);
            }
            return script;
        }
    }

    // utility class to extend the functionality of ExecuteSpecification
    public static class ExecuteSpecificationExtensions
    {
        // get the first parameter of an exec sp_executesql statement
        public static string GetFirstParameter(this ExecuteSpecification execSpec)
        {
            return execSpec.ExecutableEntity.Parameters[0].ParameterValue.Script();
        }

        // get a list of parameter pairs from an exec sp_executesql statement (eg @LastUsed, @P0003)
        public static List<(string, string)> GetParameterPairs(this ExecuteSpecification execSpec)
        {
            List<(string, string)> pairs = new List<(string, string)>();
            for (int i = 2; i < execSpec.ExecutableEntity.Parameters.Count; i++)
            {
                Trace.WriteLine($"Variable {execSpec.ExecutableEntity.Parameters[i].Variable.Name} is in {execSpec.ExecutableEntity.Parameters[i].ParameterValue.Script()}");

                pairs.Add((execSpec.ExecutableEntity.Parameters[i].Variable.Name, execSpec.ExecutableEntity.Parameters[i].ParameterValue.Script()));
            }
            return pairs;
        }
    }
}