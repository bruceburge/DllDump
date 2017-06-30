<Query Kind="Program">
  <Reference Relative="..\..\Downloads\Microsoft.Search.Interop\Microsoft.Search.Interop.dll">C:\Users\bruce\Downloads\Microsoft.Search.Interop\Microsoft.Search.Interop.dll</Reference>
  <Namespace>Microsoft.Search.Interop</Namespace>
  <Namespace>System.Runtime.InteropServices</Namespace>
</Query>

/*add reference to Microsoft.Search.Interop.dll*/

void Main()
{
	string path = (System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
	string path2 = @"C:\Users\bruce\AppData\Local\_Company_\_Product_\_Version_\IndexableDirectory"; //Non indexed folder
	string path3 = @"C:\Users\bruce\Desktop\DAs_Notebook\test";

	IsLocationIndexed(path).Dump(path);
	IsLocationIndexed(path2).Dump(path2);

	if (Directory.Exists(path2))
	{
		if (!IsLocationIndexed(path2))
		{
			IndexLocations(new string[] { path2 });
		}
	}
	else
	{
		$"{path2} doesn't exist, won't be added to index".Dump();		
	}
	
	IsLocationIndexed(path2).Dump(path2);
}

//Methods from http://blogs.microsoft.co.il/sasha/2009/02/08/enable-windows-search-indexing-on-folders/


/// <summary>
/// Adds the specified locations to the system index.
/// </summary>
/// <param name="locations">The locations to add.</param>
public static void IndexLocations(params string[] locations)
{
	CSearchManager searchManager = new CSearchManager();
	CSearchCatalogManager catalogManager = searchManager.GetCatalog("SystemIndex");
	CSearchCrawlScopeManager scopeManager = catalogManager.GetCrawlScopeManager();
// fully URL-decoded and without URL control codes. For example, file:///c:\My Documents is fully URL-decoded, whereas file:///c:\My%20Documents is not.
	foreach (string location in locations)
	{
		string url = location;
		if (url[url.Length - 1] != '\\')
            url += '\\';
		if (!url.StartsWith("file:///"))
			url = "file:///" + url;

		scopeManager.AddUserScopeRule(url, 1, 0, 0);
	}
	scopeManager.SaveAll();

	Marshal.ReleaseComObject(scopeManager);
	Marshal.ReleaseComObject(catalogManager);
	Marshal.ReleaseComObject(searchManager);
}

public static bool IsLocationIndexed(string location)
{
	CSearchManager searchManager = new CSearchManager();
	CSearchCatalogManager catalogManager = searchManager.GetCatalog("SystemIndex");
	CSearchCrawlScopeManager scopeManager = catalogManager.GetCrawlScopeManager();

	int result = scopeManager.IncludedInCrawlScope(location);

	Marshal.ReleaseComObject(scopeManager);
	Marshal.ReleaseComObject(catalogManager);
	Marshal.ReleaseComObject(searchManager);

	return result != 0;
}
// Define other methods and classes here
