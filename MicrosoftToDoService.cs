using System;
using Microsoft.Graph;
using Azure.Identity;
using Microsoft.Graph.Models;

namespace ToDoAgent
{
    public class MicrosoftToDoService
    {
        private GraphServiceClient _graphClient;

        public List<TodoTaskList> TodoTaskLists { get; set; } = new List<TodoTaskList>();

        public MicrosoftToDoService()
        {
            // Initialize the service, e.g., authenticate with Microsoft Graph API
            var scopes = new[] { "User.Read" };
            var clientId = Environment.GetEnvironmentVariable("OUTLOOK_AGENT_CLIENT_ID");
            var tenantId = "common";
            var clientSecret = Environment.GetEnvironmentVariable("OUTLOOK_AGENT_CLIENT_SECRET");
            var options = new TokenCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud };
            var interactiveCredential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
            {
                ClientId = clientId,
                TenantId = tenantId
            });
            _graphClient = new GraphServiceClient(interactiveCredential, scopes);
        }

        public void AddTask(string taskName)
        {
            // Logic to add a task to Microsoft To Do
        }

        public void RemoveTask(string taskName)
        {
            // Logic to remove a task from Microsoft To Do
        }

        public void ListTasks()
        {
            // Logic to list all tasks in Microsoft To Do
        }

        public async Task<List<TodoTaskList>> GetAllActiveTasksAsync()
        {
            var allActiveTasks = new List<TodoTaskList>();
            
            var todoTaskLists = await GetListsAsync();

            if (todoTaskLists == null || todoTaskLists.Count == 0)
            {
                Console.WriteLine("No task lists found. Cannot retrieve My Day tasks.");
            } else {
                foreach (var list in todoTaskLists)
                {
                    var todoTaskList = new TodoTaskList
                    {
                        Id = list.Id,
                        DisplayName = list.DisplayName,
                        Tasks = await GetActiveTasksForListAsync(list.Id)
                    };

                    if (todoTaskList.Tasks.Count > 0)
                    {
                        allActiveTasks.Add(todoTaskList);
                    }
                }
            }          
            
            todoTaskLists = allActiveTasks;
            return allActiveTasks;
        }

        public async Task<List<TodoTask>> GetActiveTasksForListAsync(string listId)
        {
            var taskList = new List<TodoTask>();

            var todoTasksResponse = await _graphClient.Me.Todo.Lists[listId].Tasks
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = "status ne 'completed'";
                });
            var todoTasks = todoTasksResponse?.Value;
            if (todoTasks == null || todoTasks.Count == 0)
            {
                Console.WriteLine($"No tasks found in list: {listId}");
            } else {
                foreach (var task in todoTasks)
                {
                    taskList.Add(task);
                    Console.WriteLine($"Task: {task.Title}, Status: {task.Status}, Due: {task.DueDateTime?.DateTime}");
                }
            }

            return taskList;
        }

        public async Task<List<TodoTaskList>> GetListsAsync()
        {
            // Logic to get all task lists in Microsoft To Do
            var todoTaskListsResponse = await _graphClient.Me.Todo.Lists.GetAsync();
            var todoTaskLists = todoTaskListsResponse?.Value;
            if (todoTaskLists != null)
            {
                foreach (var list in todoTaskLists)
                {
                    Console.WriteLine($"List: {list.DisplayName}");                    
                }
            }
            else
            {
                Console.WriteLine("No task lists found.");
            }
            return todoTaskLists;
        }
    }
}