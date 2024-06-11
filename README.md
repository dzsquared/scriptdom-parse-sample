# ScriptDOM sample for handling SQL statements with sp_executesql

This sample assumes the input is a T-SQL script that uses sp_executesql once to execute a SQL statement. The SQL statement may contain variables that are passed as parameters to sp_executesql.

Example input T-SQL content:
```sql
declare @p0001 nvarchar(73)
set @p0001 = N'UPDATE session SET ss_last_used = @LastUsed WHERE ss_sequence = @sequence'
declare @p0002 nvarchar(32)
set @p0002 = N'@LastUsed datetime,@sequence int'
declare @P0003 datetime set @P0003 = '20240305 15:53:09.093'
declare @P0004 int
set @P0004 = 9482

exec sp_executesql @p0001, @p0002, @LastUsed = @P0003, @sequence = @P0004
```
(line breaks added for readability)

Using [ScriptDOM](https://learn.microsoft.com/en-us/dotnet/api/microsoft.sqlserver.transactsql.scriptdom?view=sql-transactsql-161) we can parse the T-SQL script and extract different versions of the SQL statement executed by sp_executesql.

> [!INFO]
> ScriptDOM is open source and available on [GitHub](https://github.com/microsoft/sqlscriptdom).

## Option 1: what is the SQL statement passed to sp_executesql?

This returns a string based on the content of the variable passed as the first parameter to sp_executesql.

option 1 should output:
    
```sql
UPDATE session SET ss_last_used = @LastUsed WHERE ss_sequence = @sequence
```

## Option 2: what is the SQL statement executed?

This parses variable substitution such that we get the T-SQL from option 1 after replacing the variables with their values.

option 2 should output:

```sql
UPDATE session SET ss_last_used = '20240305 15:53:09.093' WHERE ss_sequence = 9482
```

## Try it yourself

> [!NOTE]
> The sample code is written in C# and uses the .NET Core SDK. You can download the .NET Core SDK from [here](https://dotnet.microsoft.com/download).

To run the sample, you can use the following command:

```bash
dotnet run
```

To run the sample without the additional output provided to improve your understanding of the process, you can use the following command:

```bash
dotnet run -c Release
```

### Change the code to dynamically accept input

To change the code to accept input from the command line, you can modify the `Program.cs` file.

Replace these 2 lines:
```cs
string filePath = "sample.sql";
int outputTypeInt = 1;
```

with:
```cs
//  get file path from user
Console.WriteLine("Enter SQL file path:");
string? filePath = Console.ReadLine();
while (filePath == null)
{
    filePath = Console.ReadLine();
}
// get the output type from the user
Console.WriteLine("Enter output type (1 for initial sp_executesql query, 2 for substituted query):");
string? outputType = Console.ReadLine();
int outputTypeInt = 0;
while (Int32.TryParse(outputType, out outputTypeInt) == false)
{
    outputType = Console.ReadLine();
}
```

### What are these visitors?

The `TSqlFragmentVisitor` is a class that allows you to visit each node in the parsed T-SQL script. The `TSqlFragmentVisitor` class is abstract, so you need to create a class that inherits from it and override the methods you want to use.

In this sample, we have three visitors:
- **VariableFinderVisitor**: gives us access to all of the `SetVariableStatement` nodes, eg `SET @P0003 = '20240305 15:53:09.093'`
- **spExecuteVisitor**: gives us access to all of the `ExecuteStatement` nodes and we filter it to when the procedure called is `sp_executesql`, eg `exec sp_executesql @p0001, @p0002, @LastUsed = @P0003, @sequence = @P0004`
- **VariableReferenceVisitor**: gives us access to all of the `VariableReference` nodes, eg `@LastUsed` out of `UPDATE session SET ss_last_used = @LastUsed WHERE ss_sequence = @sequence`

When we create an instance of a visitor and pass it to the script (as in the `Accept` method), the visitor will traverse the script and call the appropriate method for each node it encounters.

## References
https://www.sqlservercentral.com/steps/stairway-to-scriptdom-level-1-an-introduction-to-scriptdom
https://deep.data.blog/2013/04/04/using-the-transactsql-scriptdom-parser-to-get-statement-counts/
https://devblogs.microsoft.com/azure-sql/programmatically-parsing-transact-sql-t-sql-with-the-scriptdom-parser/

## License

This sample is available under the MIT license. For more information, see the [LICENSE](LICENSE) file. No support is implied.