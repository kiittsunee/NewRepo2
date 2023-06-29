namespace DotNetApi.Models
{
    public class TodoItemDTO
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public bool IsComplete { get; set; }
    }
    public class TodoItem
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public bool IsComplete { get; set; }
        public string? Secret { get; set; }

        public static implicit operator List<object>(TodoItem? v)
        {
            throw new NotImplementedException();
        }
    }
}
