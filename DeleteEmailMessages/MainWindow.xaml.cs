using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Configuration;

namespace DeleteEmailMessages
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public AppConfiguration AppConfig { get; set; }

        public ObservableCollection<FolderListItem> FolderListItems { get; set; }

        public Progress<string> TaskProgress { get; set; }


        public MainWindow()
        {
            InitializeComponent();

            AppConfig = ReadConfig();
            gridConfiguration.DataContext = AppConfig;
            passwordBox.Password = "";

            FolderListItems = new ObservableCollection<FolderListItem>();
            listViewFolderListItems.ItemsSource = FolderListItems;

            datePickerDate.SelectedDate = DateTime.Today;

            TaskProgress = new Progress<string>();
            TaskProgress.ProgressChanged += TaskProgress_ProgressChanged;
        }

        private AppConfiguration ReadConfig()
        {
            try
            {
                var appConfig = new AppConfiguration();

                var Host = ConfigurationManager.AppSettings[appConfig.HostKeyName];
                var Port = ConfigurationManager.AppSettings[appConfig.PortKeyName];
                var UseSsl = ConfigurationManager.AppSettings[appConfig.UseSslKeyName];
                var Username = ConfigurationManager.AppSettings[appConfig.UsernameKeyName];

                if (Host != "")
                    appConfig.Host = Host;

                int portOut;
                if (int.TryParse(Port, out portOut))
                    appConfig.Port = portOut;

                bool useSslOut;
                if (bool.TryParse(UseSsl, out useSslOut))
                    appConfig.UseSsl = useSslOut;

                if (Username != "")
                    appConfig.Username = Username;

                return appConfig;
            }
            catch (Exception ex)
            {
                string message = "There was an erro while reading configuration file. Using default settings";
                MessageBox.Show(message);
                logMessage($"Error: {ex.ToString()}");

                return new AppConfiguration()
                {
                    Host = "",
                    Port = 993,
                    UseSsl = true,
                    Username = ""
                };
            }

        }

        private Task WriteConfig(AppConfiguration appConfig, IProgress<string> progress)
        {
            return Task.Run(() =>
            {
                try
                {
                    Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                    // Host
                    config.AppSettings.Settings.Remove(appConfig.HostKeyName);
                    config.AppSettings.Settings.Add(appConfig.HostKeyName, appConfig.Host);

                    // Port
                    config.AppSettings.Settings.Remove(appConfig.PortKeyName);
                    config.AppSettings.Settings.Add(appConfig.PortKeyName, appConfig.Port.ToString());

                    // UseSsl
                    config.AppSettings.Settings.Remove(appConfig.UseSslKeyName);
                    config.AppSettings.Settings.Add(appConfig.UseSslKeyName, appConfig.UseSsl.ToString());

                    // Username
                    config.AppSettings.Settings.Remove(appConfig.UsernameKeyName);
                    config.AppSettings.Settings.Add(appConfig.UsernameKeyName, appConfig.Username);

                    // Save to file
                    config.Save(ConfigurationSaveMode.Modified);

                    var confirmationMessage = "Configuration saved successfully!";

                    progress.Report(confirmationMessage);
                    MessageBox.Show(confirmationMessage);
                }
                catch (Exception ex)
                {

                    string message = "The configuration could not be saved. See logs for more information.";
                    MessageBox.Show(message);
                    progress.Report($"Error: {ex.ToString()}");
                }
                
            });
        }

        private void TaskProgress_ProgressChanged(object sender, string e)
        {
            logMessage(e);
        }

        private async void btnDeleteEmails_Click(object sender, RoutedEventArgs e)
        {
            if (FolderListItems.FirstOrDefault(x => x.IsSelected) == null)
            {
                MessageBox.Show("You should select at least 1 folder.");
                return;
            }

            var totalMessagesToBeDeleted = FolderListItems
                .Where(x => x.IsSelected)
                .Select(x => x.TotalFilteredMessages)
                .Aggregate((a, b) => a + b);

            var messageBoxResult = MessageBox.Show(
                $"Do you really want to delete {totalMessagesToBeDeleted} messages that are older than {datePickerDate.SelectedDate.Value.ToShortDateString()}?"
                , "Are You Sure?"
                , MessageBoxButton.YesNo
                , MessageBoxImage.Question
                , MessageBoxResult.No);

            if (messageBoxResult == MessageBoxResult.No)
                return;

            try
            {
                buttonDeleteEmails.IsEnabled = false;
                datePickerDate.IsEnabled = false;
                using (var client = await ConnectToServerAsync(AppConfig, passwordBox.Password, TaskProgress))
                {
                    var selectedFolderListItems = FolderListItems.Where(x => x.IsSelected).ToList();

                    logMessage($"Start attempt to delete {totalMessagesToBeDeleted} messages.");

                    await DeleteSelectedMessagesAsync(client, selectedFolderListItems, TaskProgress);

                    client.Disconnect(true);

                    logMessage($"Disconnected.");

                    // Update filtered messages
                    btnFilter_Click(sender, e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
                logMessage($"Error message: {ex.Message}");
                logMessage($"Error stack trace: {ex.StackTrace}");
            }
            finally
            {
                buttonDeleteEmails.IsEnabled = true;
                datePickerDate.IsEnabled = true;
            }
        }

        private Task DeleteSelectedMessagesAsync(ImapClient client, List<FolderListItem> folderListItems, IProgress<string> progress)
        {
            return Task.Run(async () =>
            {
                var totalMessagesDeleted = 0;
                foreach (var folderListItem in folderListItems)
                {
                    var totalMessagesToBeDeleted = folderListItem.MessageUniqueIds.Count;
                    var folderName = folderListItem.FolderName;

                    progress.Report($"Deleting {totalMessagesToBeDeleted} messages in folder '{folderName}'.");

                    var folder = await client.GetFolderAsync(folderName);
                    await folder.OpenAsync(FolderAccess.ReadWrite);
                    await folder.AddFlagsAsync(folderListItem.MessageUniqueIds, MessageFlags.Deleted, true);
                    await folder.CloseAsync(true);

                    totalMessagesDeleted += totalMessagesToBeDeleted;
                    progress.Report($"{totalMessagesToBeDeleted} messages deleted successfully in folder '{folderName}'.");
                }

                progress.Report($"Total of {totalMessagesDeleted} messages were deleted successfully.");
            });
        }

        private Task<ImapClient> ConnectToServerAsync(AppConfiguration config, string password, IProgress<string> progress)
        {
            return Task.Run(async () =>
            {
                var client = new ImapClient();
                // Accept all SSL certificates
                client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                progress.Report($"Connecting to host '{config.Host}:{config.Port}'...");
                await client.ConnectAsync(config.Host, config.Port, config.UseSsl);
                progress.Report($"Connected succesfully.");
                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                progress.Report($"Authenticating with username '{config.Username}'...");
                await client.AuthenticateAsync(config.Username, password);
                progress.Report($"Authenticated successfully");

                return client;
            });
        }

        private Task<List<string>> GetFolderNamesAsync(ImapClient client, IMailFolder folder, IProgress<string> progress)
        {
            return Task.Run(async () =>
            {
                progress.Report($"Getting folder names.");

                List<string> folderNames = new List<string>();
                folderNames.Add(folder.FullName);

                var subfolders = await folder.GetSubfoldersAsync(StatusItems.None);
                foreach (var subfolder in subfolders)
                {
                    //progress.Report($"Folder name: {subfolder.FullName}");
                    if (!folderNames.Contains(subfolder.FullName))
                        folderNames.Add(subfolder.FullName);
                }

                progress.Report($"Finished getting folder names.");

                return folderNames;
            });
        }

        private Task<ObservableCollection<FolderListItem>> GetFolderListItemsAsync(
            ImapClient client,
            List<string> folderNames,
            DateTime deliveredBeforeDate,
            IProgress<string> progress)
        {
            return Task.Run(async () =>
            {
                progress.Report($"Getting folder list items for messages delivered before {deliveredBeforeDate.ToShortDateString()}.");
                progress.Report($"Showing only folders that the number of filtered messages are greater than 0.");
                var folderListItems = new ObservableCollection<FolderListItem>();
                foreach (var folderName in folderNames)
                {
                    var folder = await client.GetFolderAsync(folderName);
                    await folder.OpenAsync(FolderAccess.ReadOnly);
                    var query = SearchQuery.DeliveredBefore(deliveredBeforeDate);
                    var result = await folder.SearchAsync(SearchOptions.All, query);

                    var numberOfMessages = result.Count;
                    if (numberOfMessages != 0)
                    {
                        folderListItems.Add(
                            new FolderListItem()
                            {
                                IsSelected = true,
                                FolderName = folderName,
                                TotalFilteredMessages = numberOfMessages,
                                MessageUniqueIds = result.UniqueIds
                            });

                        progress.Report($"Folder: {folderName}, Number of messages: {numberOfMessages}");
                    }

                    await folder.CloseAsync();
                }

                if (folderListItems.Count == 0)
                    progress.Report($"Zero messages were delivered before {deliveredBeforeDate.ToShortDateString()}");

                return folderListItems;
            });
        }

        private async void btnFilter_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                buttonFilter.IsEnabled = false;
                datePickerDate.IsEnabled = false;
                listViewFolderListItems.IsEnabled = false;
                using (var client = await ConnectToServerAsync(AppConfig, passwordBox.Password, TaskProgress))
                {
                    var folderNames = await GetFolderNamesAsync(client, client.Inbox, TaskProgress);
                    var folderListItems = await GetFolderListItemsAsync(client, folderNames, datePickerDate.SelectedDate.Value, TaskProgress);

                    FolderListItems.Clear();
                    foreach (var folderListItem in folderListItems)
                    {
                        FolderListItems.Add(folderListItem);
                    }

                    client.Disconnect(true);

                    logMessage($"Disconnected.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
                logMessage($"Error message: {ex.Message}");
                logMessage($"Error stack trace: {ex.StackTrace}");
            }
            finally
            {
                buttonFilter.IsEnabled = true;
                datePickerDate.IsEnabled = true;
                listViewFolderListItems.IsEnabled = true;
            }

        }

        private string getFormattedCurrentDate()
        {
            return DateTime.Now.ToString();
        }

        private void logMessage(string message)
        {
            richTextBoxOutput.AppendText($"{getFormattedCurrentDate()} - {message}");
            richTextBoxOutput.AppendText(Environment.NewLine);
            richTextBoxOutput.ScrollToEnd();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private async void buttonSaveConfiguration_Click(object sender, RoutedEventArgs e)
        {
            await WriteConfig(AppConfig, TaskProgress);
        }
    }
}
