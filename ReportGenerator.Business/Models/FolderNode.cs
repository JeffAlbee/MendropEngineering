namespace ReportGenerator.Business.Models
{
    public class FolderNode
    {
        public string Name { get; }

        public List<FolderNode> Children { get; }

        public FolderNode(string name, IEnumerable<FolderNode>? children = null)
        {
            Name = name;
            Children = children?.ToList() ?? new List<FolderNode>();
        }
    }
}
