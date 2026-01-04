using ReportGenerator.Business.Models;

namespace ReportGenerator.Business.Extensions
{
    public static class FolderBlueprintExtensions
    {
        public static IEnumerable<string> ToPaths(this FolderNode node, string parent = "")
        {
            string current = string.IsNullOrEmpty(parent)
                ? node.Name
                : $"{parent}/{node.Name}".Trim('/');

            if (!string.IsNullOrWhiteSpace(current))
                yield return current;

            foreach (var child in node.Children)
            {
                foreach (var path in child.ToPaths(current))
                    yield return path;
            }
        }
    }
}
