using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Azure.Cosmos;
using Newtonsoft.Json;

namespace TRVautomator
{
    public partial class LoginForm : Form
    {    // ADD THIS PART TO YOUR CODE

        // The Azure Cosmos DB endpoint for running this sample.
        private static readonly string EndpointUri = "https://logincosmos.documents.azure.com:443/";
        // The primary key for the Azure Cosmos account.
        private static readonly string PrimaryKey = "0TpTW6p6ONP8QU1Y8UAkMlCRCbfAN9sa1JAGakgtIE80tpLqdu6gOWAciGIxrM0ZhSJDhnL6HQKpr2i3wiC1zQ==";

        // The Cosmos client instance
        private CosmosClient cosmosClient;

        // The database we will create
        private Database database;

        // The container we will create.
        private Microsoft.Azure.Cosmos.Container container;

        public bool showAutomator = false;

        public bool IsTraining = false;
        public bool IsAdmin = false;

        // The name of the database and container we will create
        private string databaseId = "UserData";
        private string containerId = "LoginData";
        public LoginForm()
        {
            ConnectToCosmos();
            InitializeComponent();
            this.BackgroundImageLayout = ImageLayout.Zoom;
        }

        private async void loginButton_Click(object sender, EventArgs e)
        {
            string originalText = loginButton.Text;
            loginButton.Text = "Loading...";
            loginButton.Enabled = false;
            if(await QueryItemsAsync())
            {
                showAutomator = true;
                this.Close();
            }
            else
            {
                loginErrorLabel.Text = "Wrong credentials!";
            }
            loginButton.Enabled = true;
            loginButton.Text = originalText;
        }
        private void ConnectToCosmos()
        {
            cosmosClient = new CosmosClient(EndpointUri, PrimaryKey);
            database = cosmosClient.GetDatabase(databaseId);
            container = database.GetContainer(containerId);
        }
        private async Task<bool> QueryItemsAsync()
        {
            var sqlQueryText = "SELECT * FROM c WHERE c.id = '"+usernameTextBox.Text+"' AND c.password = '"+passwordTextBox.Text+"'";

            QueryDefinition queryDefinition = new(sqlQueryText);
            using FeedIterator<UserModel> queryResultSetIterator = this.container.GetItemQueryIterator<UserModel>(queryDefinition);

            List<UserModel> users = new();

            while (queryResultSetIterator.HasMoreResults)
            {
                FeedResponse<UserModel> currentResultSet = await queryResultSetIterator.ReadNextAsync();
                foreach (UserModel user in currentResultSet)
                {
                    users.Add(user);
                }
            }

            if(users.Count > 0)
            {
                IsTraining= users[0].IsTraining;
                IsAdmin= users[0].IsAdmin;
                return true;
            }
            else
            {
                return false;
            }
        }

        private void passwordTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                loginButton_Click(this, new EventArgs());
            }
        }
    }
}
