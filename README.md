# DeleteEmailMessages
DeleteEmailMessages - A C# WPF windows application to bulk delete email messages through IMAP protocol using https://github.com/jstedfast/MailKit lib.

## How to Build
Open the project with Visual Studio 2015+ and build the solution.

## Usage
- Open 'Configuration' tab and set Host, Port, Username and Password options;
- Select a date to filter messages in all email folders that are older than it;
- Click 'Filter' button;
- Select checkboxes of folders you want to delete;
- Click 'Delete Selected', then click 'Yes' to the prompt message.

Open 'Logs' tab to check logs of the program. They're in-memory logs. Once you close the application they're not persisted (yet).
