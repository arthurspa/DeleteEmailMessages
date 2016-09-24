# DeleteEmailMessages
DeleteEmailMessages - A C# WPF windows application to bulk delete email messages through IMAP protocol using https://github.com/jstedfast/MailKit lib.

## How to Build
Open the project with Visual Studio 2015+ and build the solution.

## Usage
- Open 'Configuration' tab and set Host, Port, Username and Password options;
![Configuration](http://image.prntscr.com/image/bfcdc543e4a845c4a50c989e7bb4af02.png)
- Select a date to filter messages in all email folders that are older than it;
- Click 'Filter' button;
- Select checkboxes of folders you want to delete;
![Folders Selection](http://image.prntscr.com/image/fc8fc930d59045978911884dd471a7ab.png)
- Click 'Delete Selected', then click 'Yes' button into the prompt message.
![Confirm Deletition](http://image.prntscr.com/image/350af54ba77d4e0e8ca4b11e0cd81b65.png)

### Note
Open 'Logs' tab to check logs of the program. They're in-memory logs. Once you close the application they're not persisted (yet).

![Logs](http://image.prntscr.com/image/b4ae538fde6345558fb0593f9cd031b8.png)
