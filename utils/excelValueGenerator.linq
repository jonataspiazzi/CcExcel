<Query Kind="Program">
  <Namespace>Microsoft.CSharp</Namespace>
  <Namespace>System.CodeDom</Namespace>
</Query>

void Main()
{
	var types = GetTypes()
		.Select(type => new
		{
			Name = type.Name,
			Alias = GetAlias(type)
		});

	var tt = File.ReadAllText(Path.Combine(Environment.CurrentDirectory, "excelValueGenerator.tt"));
	
	foreach (var type in types)
	{
		var text = tt;
		
		text = Regex.Replace(text, @"[<]#=\s*type[.]Name\s*#[>]", type.Name);
		text = Regex.Replace(text, @"[<]#=\s*type[.]Alias\s*#[>]", type.Alias);
		
		text.Dump();
	}
}

private IEnumerable<Type> GetTypes()
{
	foreach (var type in typeof(byte)
		.Assembly
		.GetTypes()
		.Where(w => w.IsPrimitive))
	{
		if (type.Name.Contains("Ptr")) continue;
		
		yield return type;
	}
	
	yield return typeof(decimal);
	yield return typeof(DateTime);
	yield return typeof(TimeSpan);
	yield return typeof(string);
}

private string GetAlias(Type type)
{
	using (var provider = new CSharpCodeProvider())
	{
		var typeName = provider.GetTypeOutput(new CodeTypeReference(type));

		typeName = Regex.Replace(typeName, @"(\w+[.])*(?<name>\w+)", "${name}");
		typeName = Regex.Replace(typeName, @"Nullable[<](?<name>\w+)[>]", "${name}?");

		return typeName;
	}
}