# MMM-MicrosoftToDo
MMM-MicrosoftToDo is a 3rd party module for [MagicMirror](https://github.com/MichMich/MagicMirror) that shows open tasks from task lists. The tasks are retrieved from Microsoft To Do / Outlook.

The idea of this module is to display a shopping list on your mirror that you can control from your Microsoft To Do app and collaborate on with your family. Originally I used [MMM-Wunderlist](https://github.com/paviro/MMM-Wunderlist), though Wunderlist will be shut down and replaced by Microsoft To Do.

## Functionality
* Display open to do list items from a to do list managed by Microsoft To Do
* More than one instance can be created
* Update contents at user configurable intervals
* Display avatars for assigned tasks
* Sort task by various criteria
* Use name for folder, not its ID
* Can display time of last update

Preparation
======

In order to configure this module, you'll need some configuration parameters from Microsoft, which can be obtained as described in this section. All steps are required as the configuration parameters are needed for the installation below.

### Step 1: Web App Registration (Client ID)
Register an app at https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps (sign-in first if needed)
1. Click on "+ New registration"
2. Enter as "Name": "Magic Mirror"
3. Select as "Supported account types": "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
4. Select as "Redirect URI": "Web"
5. Set as "Redirect URI": "https://localhost:1234"
6. Click "Register"
7. Go to "Overview" and copy the "Application (client) ID" (Client ID), e.g. 4ef19f40-4892-4905-b999-76041e991f53`
8. This will be your `oauth2ClientId` in the configuration file


### Step 2: Client Secret Key Generation (Client Secret)
1. Go to "Certificates & secrets"
2. Click "+ New client secret"
3. Set as "Description": "Magic Mirror"
4. Select as "Expires": "Never"
5. Click "Add"
6. Copy the "Value", e.g. `1Q25ls?TKEDf4RWBKVUbKjnaVu=ytP.[`, it will be the client secret you need later
7. This will be your `oauth2ClientSecret` in the configuration file

### Step 3: Refresh Token
#### Authorization Code Generation (Code)
Create authorization code by opening the following URL in your browser (but replace the client_id with your value):
```
https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&client_id=4ef19f40-4892-4905-b999-76041e991f53&scope=offline_access%20user.read%20tasks.read&redirect_uri=https://localhost:1234&response_mode=query&state=12345
```
1. Click on "Yes" and wait to be redirected
2. Copy the `code` parameter from the URL in the browser, e.g. `M30cd0dff-af91-d061-8755-ffb3b328aa03` from
```
https://localhost:1234/?code=M30cd0dff-af91-d061-8755-ffb3b328aa03&state=12345
```
3. This will be used in the next step to generate a Refresh Token

#### Refresh Token Generation (Refresh Token)
Generate refresh token to allow the module to retrieve your task list going forward
1. Open Terminal app
2. Run the following command with your `client_id` (Client ID), `code` (Code) and `client_secret` (Client Secret) parameters
```bash
curl -X POST \
  https://login.microsoftonline.com/common/oauth2/v2.0/token \
  -H 'Content-Type: application/x-www-form-urlencoded' \
  -d 'grant_type=authorization_code&client_id=4ef19f40-4892-4905-b999-76041e991f53&scope=offline_access%20user.read%20tasks.read&code=M30cd0dff-af91-d061-8755-ffb3b328aa03&redirect_uri=https%3A%2F%2Flocalhost%3A1234&client_secret=1Q25ls?TKEDf4RWBKVUbKjnaVu=ytP.['
```
3. Copy the value of the `refresh_token` parameter from the response, e.g. `refresh_token`:
```
MCVDmtyumC8ZdlGxrkMwCdur4FbPh7GHjy1gOaOHEA8*c5tTI4oIcJAHU8AflLZMLmcCRWNEa7nMkag2hbipYbZ3QgL9JB1G6Hc*JZY9oA5j46Cq79b6BDx8mivTVen0r*39h2hw9pngKS7MGXjKc9efwayQ7UhsFoCJ2wdhLJzv5gTR2EozgcsKsxHODj!1fV8a1sqVqJ!sU*CBLTMsm0twI2hZqsCkPoSWaMIS5kKFAq7bCsfNaumS11TZYni3XjVwuyHn4DEqfLGjHuSNo1TYRifcG422MdG6bYt7tBxWiw5dlq2iIEMl6O*N2sLWU7W0Ns2*oQbcXIFxRsZ8cHP*6M9dn*9Axg8WbtERj!*jHTfC!Ax6IILXrHDJkdAXgh01PVH!huztVrgd0UC6E4R4zJSOa4ytOKWfQcL9h0NdA
```
from the full response
```javascript
{"token_type":"Bearer","scope":"User.Read Tasks.Read","expires_in":3600,"ext_expires_in":3600,"access_token":"EwBwA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAAQMC79sEwx4KyeAHQmWi48AYLyFufrOurKrSQRBVaB1CPFwrn/CniDdPAEQYSzb95gEseUsnH3QAjbfi1s5Ya8oW5tLW2atfMG8ttdCnN/Rl8/ZdzfZcFkFQSgJ3kibm33Ov9NGXkKubHBgk0etwhyBsFR25PcDRZgc9hLIgR3dt9OCTF+Ph5lRQqOX8ABQbYy6zm15BHKSOqbHF82QHntrqUZf9u94J7S0ez0reltsOdrFTK0rWcbTMAWjNQfBLoMvdJnqVgTSF++WvkTQCCjTpu+LDEfA7axWBDfAQIXh+KFh2OfLVhST4NKPQNrRgsyI0Mi3kKak/gC7/are8bxQDZgAACHz3HW3uBQHCQAKbkNX8hKy6+AODSWnD/zP0vkRXsoRkiR3JJH7VxrnB31qo9W3Y9PokxT2SBj/97BPmuiJDJ98j190k+MJ/W1kLsL7PbZN9svwLEk8UpPDF2MDiu7GTdhvS7fPDODRBBrK75fU+n1wynvTas+q3ybFM+d0fiqW0p9eUxb7GsK1A0BwXpARW+oFQ2W1g7lzq/w2ss+NyvbjUpizEWw9hUFgJhi32b3ww/r7f0faZRz9vmsLSahbvywo2HG7gzJRdSFbUnoxXRJdFeQpCfxRIhLbE/fx9rB9e0+ENWRuuWqpuJNFjHW9VcV837KAqEwJE9T0F0bacYIq1mcm/n+7g6HSOs27H9LtEHZec6Z+nPPYLxFiN2uRqKaVC7ZBCf6DLfUtQ9sbaDlzI+UUe+GnCTnc2is/hbjYPtK/leiUvOk3jxYBGuU2Jb1kPxH3ahRFlmBYB0mQtbqgOa9DNi0An9eMy+Y/uaLP7EVweQd2bFcBt7tF1nLMQQNUNBfYUva/b2S4SsTVsbHp2tCRmrufFmLrQtFDaG/zJY+tZmnwy8QjbYch9CwEMWRe8+mrtwWy4JKwK00SxoApoDgsaQK/aRYkR5+kn1N/NSwkByyKK6oMbRlFf9SnRsqdFIlgW7gmeizEPhRGqvwTnltghoQdM/aSqY53UQ2jd6H0vYy1PFBtq2s6EGL0Mf9az7JguOgee/WBrI3bN4hQn/xGnrGpkaghsD/2mtRxOAkoEMdxfYQRO60eknbF2QsyYjoOS6V/9ClhzAg==","refresh_token":"MCVDmtyumC8ZdlGxrkMwCdur4FbPh7GHjy1gOaOHEA8*c5tTI4oIcJAHU8AflLZMLmcCRWNEa7nMkag2hbipYbZ3QgL9JB1G6Hc*JZY9oA5j46Cq79b6BDx8mivTVen0r*39h2hw9pngKS7MGXjKc9efwayQ7UhsFoCJ2wdhLJzv5gTR2EozgcsKsxHODj!1fV8a1sqVqJ!sU*CBLTMsm0twI2hZqsCkPoSWaMIS5kKFAq7bCsfNaumS11TZYni3XjVwuyHn4DEqfLGjHuSNo1TYRifcG422MdG6bYt7tBxWiw5dlq2iIEMl6O*N2sLWU7W0Ns2*oQbcXIFxRsZ8cHP*6M9dn*9Axg8WbtERj!*jHTfC!Ax6IILXrHDJkdAXgh01PVH!huztVrgd0UC6E4R4zJSOa4ytOKWfQcL9h0NdA“}
```
4. This will be your `oauth2RefreshToken` in the configuration file
   
Avatars
======
Microsoft ToDo has the capability to assign tasks to somebody else in the app or the web interface.  This effectively allows to manage a list of task for a set of people from a single task folder.
<br>Unfortunately, at this time, ToDo has a bug in its developper API which does not fill out the `assignedTo` property necessary to make that determination.  Until this bug is fixed there is no way to identify the assigned person responsible for a task. See this [issue](https://github.com/thobach/MMM-MicrosoftToDo/issues/12).</br>
<br>
The avatar feature enables this assignement feature to work again at the expense of an extra step being performed when creating a task.  Basically it is necessary to assign a `tag` to every task in the MSToDo app, which can then be parsed back by this module.
</br>
<br>A resonable compromise it to preprend every task with a short `tag`, as illustrated below. Let's assume there are 4 members who will be assigned tasks. The mapping will be as below:
</br>
|Actual Task Name |Member Name | Tag | Final Task Name entered MSToDo App|
|---|---|---|---|
|Something to do |John Glee | JG: | JG: Something to do
|Another thing |Bob Carr | BC: | BC: Another thing
|Remember to do this |Ann Sollis | AS: | AS: Remember to do this
|Must do today |Cindy Dunn | CD: | CD: Must do today

<br>We can now configure this module as follow (the tag doesn't have to be the initials of the name):

```javascript
avatars: [
	{ tag: "JG:", assignedTo: "John Glee", icon: "simple-icons:katana" },
	{ tag: "BC:", assignedTo: "Bob Carr", icon: "twemoji-superhero-light-skin-tone" },
	{ tag: "AS:", assignedTo: "Ann Sollis", icon: "emojione-woman-medium-skin-tone" },
	{ tag: "CD:", assignedTo: "Cindy Dunn", icon: "emojione-girl-medium-skin-tone" }
]
```

The module will remove the tag when displaying the task, and display the corresponding avatar picture next to it.  This can be either an icon from `iconify` or simply a circle with the initials inside.<br>
<br>Once the issue in fixed in MSToDo, the tag can go away and be removed (both in the MSToDo App and the MMM config file, but the icons and initials functionality can remain).
</br>
<br>This feature is enabled by setting the `displayAvatar` property to `true` and providing the `avatars` array property above.

Installation
======

1. `cd ~/MagicMirror/modules` 
2. `git clone https://github.com/etsandig/MMM-MicrosoftToDo.git`
3. `cd MMM-MicrosoftToDo`
4. `npm install`

Configuration
======

## Example
```javascript
  {
    module: 'MMM-MicrosoftToDo',
    position: 'top_left',
    header: 'Family List',
    config: {
      oauth2ClientSecret: '1Q25ls%3FTKEDf4RWBKVUbKjnaVu%3DytP.%5B',
      oauth2RefreshToken: 'MCVDmtyumC8ZdlGxrkMwCdur4FbPh7GHjy1gOaOHEA8*c5tTI4oIcJAHU8AflLZMLmcCRWNEa7nMkag2hbipYbZ3QgL9JB1G6Hc*JZY9oA5j46Cq79b6BDx8mivTVen0r*39h2hw9pngKS7MGXjKc9efwayQ7UhsFoCJ2wdhLJzv5gTR2EozgcsKsxHODj!1fV8a1sqVqJ!sU*CBLTMsm0twI2hZqsCkPoSWaMIS5kKFAq7bCsfNaumS11TZYni3XjVwuyHn4DEqfLGjHuSNo1TYRifcG422MdG6bYt7tBxWiw5dlq2iIEMl6O*N2sLWU7W0Ns2*oQbcXIFxRsZ8cHP*6M9dn*9Axg8WbtERj!*jHTfC!Ax6IILXrHDJkdAXgh01PVH!huztVrgd0UC6E4R4zJSOa4ytOKWfQcL9h0NdA',
      oauth2ClientId: '4ef19f40-4892-4905-b999-76041e991f53',
      folderName: 'Family',
      hideIfEmpty: false,
      itemLimit: 200,
      displayAvatar: true,
      avatars: [
      	{ tag: "JG:", assignedTo: "John Glee", icon: "simple-icons:katana" },
      	{ tag: "BC:", assignedTo: "Bob Carr", icon: "twemoji-superhero-light-skin-tone" },
      	{ tag: "AS:", assignedTo: "Ann Sollis", icon: " emojione-woman-medium-skin-tone" },
      	{ tag: "CD:", assignedTo: "Cindy Dunn", icon: "emojione-girl-medium-skin-tone" }
      ]      
    }
  }
```
## Options
|field |value type |default value |value example |comments |
|---|---|---|---|---|
|oauth2ClientId |String | - |"client string" |**REQUIRED**<br/>The client ID acquired in step 1 above.
|oauth2ClientSecret |String | - | "secret string" |**REQUIRED**<br/>The client secret acquired in step 2 above.
|oauth2RefreshToken|String | - |"token string" | **REQUIRED**<br/>The refresh token acquired in step 3 above. |
|folderName |String |"Tasks" | "Family" | The name of the task list to dispay
|itemLimit |Integer | 30 | 50 | Limit on the number of tasks to pull from the task folder
|refreshInterval | Integer | 15 | 30 | How often to refresh the list in minutes
|sortOrder |String |"none" | "dueDate" | '`none`' : The order tasks were created<br>'`assignedTo`' : The name of the task assignee <br>'`subject`' : Sort alphabetically by subject<br>'`importance`' : Sort by importance<br>'`dueDate`' : Sort by ascending due date<br>'`reminderDate`' : Sort by ascending reminder date
|displayWhenEmpty |Boolean |false | true | Whether to show the module if no tasks are present in the task list
|displayLastUpdate |Boolean |false | true | Show how long ago the task list was last updated
|displayDecorations |Boolean |false | true | Show additional information under the task name (recurrence, reminder date, notes...)
|displayAvatar |String |"none" | "initials" | Display an avatar `icon` or name `initials` for each tasks (see [Avatars](Avatars))
|avatar |Array  | - | See [Avatars](Avatars) | An array describing the avatars to display for each tasks (see [Avatars](Avatars))

## CSS
The module comes with a `.css` file which provide the overall look of the task list.  It is possible to overwrite its value by adding directives in the `custom.css` file (do not directly change the `css` file in this module as future updates may overwrite your changes).

Some values of interest would be the `font-size`, `min-width`.
```javascript
.MSToDo {
  min-width: 400px;  /* how wide the module will be on the screen.  Tasks will be truncated with ... when too long */
  font-size: 0.6em;  /* how big the font should be for this module */
}
```

If needed it is possible to put the module into a shaded box in order to increase visibility, especially when using background images.  This is done by adding the following in the `custom.css` file 
```javascript
.module.MMM-MicrosoftToDo {
	background-color: rgba(0, 0, 0, 0.5);
	border-radius: 8px;
	padding: 10px;
}
```


The order of the the fields in the module is [`avatar`, `task name`, `dueDate`, `importance`], but it possible to change this order by modifying the `order` field from the corresponding selectors from the `css` file.
<br><br>For instance, to change the order to [`dueDate`, `task name`, `importance`, `avatar`], add the following to your `custom.css`.

```javascript
.MSToDo .importance {
	order: 3;
}

.MSToDo .caption {
	margin-left: 15px;
	order: 2;
}

.MSToDo .dueDate {
	order: 1;
}

.MSToDo .avatar {
	order: 4;
}
```
