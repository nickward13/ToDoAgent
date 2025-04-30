namespace ToDoAgent
{
    class Program
    {
        async static Task Main(string[] args)
        {
            Console.WriteLine($"Current UTC Date and Time: {DateTime.UtcNow}");

            // Initialize the MicrosoftToDoService
            var todoService = new MicrosoftToDoService();
            // list all task lists
            await todoService.GetAllActiveTasksAsync();
        }
    }
}
