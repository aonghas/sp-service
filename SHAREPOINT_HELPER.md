# SharePoint Helper Documentation

This file documents the SharePoint REST helper defined in `index.js`.

## Overview

The helper exports a singleton instance of the `SharePoint` class. It wraps Axios and provides methods for common SharePoint REST operations including file management, list/item CRUD, users, groups, comments, and recycle bin actions.

## Import

Install the package:

```bash
npm install sp-service
```

Then import it:

```js
import SP from "sp-service";
```

Or create your own instance:

```js
import { SharePoint } from "sp-service";
const sp = new SharePoint({ baseUrl: "https://your-sharepoint-site" });
```

## Configuration

- `options.baseUrl` — base URL for SharePoint API requests.
- If `options.baseUrl` is omitted, `import.meta.env.VITE_APP_URL` is used.

The helper also manages SharePoint request digests automatically via `generateDigest()`.

## REST API Documentation References

- General SharePoint REST overview: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints
- Folders and files: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
- Lists and list items: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
- Search and query: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-search-and-query-with-rest
- Comments: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-comments-by-using-sharepoint-rest-api
- Users and groups: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-user-profiles-and-site-users-rest

## Common Usage

```js
await SP.createItem("MyList", { Title: "New item" });
const items = await SP.getItems("MyList");
await SP.createFolder("/sites/site/Shared Documents", "NewFolder");
```

---

## Methods

### Authentication / Digest

- `generateDigest()`
  - Requests a new SharePoint form digest and stores it in `this.DIGEST`.
  - Returns a promise resolving to the digest string.

### Folder and File Operations

- `createFolder(folderLocation, folderName)`
  - Creates a folder at the given server-relative path.
  - Example: `SP.createFolder("/sites/site/Shared Documents", "MyFolder")`

- `createFileInFolder(library, folder, fileName, fileContents)`
  - Ensures the folder exists, then uploads a file.
  - Sends file contents as the request body.
  - Example: `SP.createFileInFolder("Documents", "MyFolder", "file.txt", fileData)`
  - See Add file to folder: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest#add-a-file-to-a-folder

- `createFile(folder, fileName, fileContents)`
  - Adds a file directly to a folder.
  - Sends file contents as the request body.
  - Example: `SP.createFile("/sites/site/Shared Documents/MyFolder", "file.txt", fileData)`
  - See Add file to folder: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest#add-a-file-to-a-folder

- `moveFile(payload)`
  - Moves or copies a file using `SP.MoveCopyUtil.MoveFileByPath()`.
  - Payload must match the SharePoint MoveCopyUtil contract.
  - Example:
    ```js
    await SP.moveFile({
      srcPath: "/sites/site/Shared Documents/MyFolder/file.txt",
      destPath: "/sites/site/Shared Documents/Archive/file.txt",
      options: { keepBoth: false }
    });
    ```
  - See Move or copy a file: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest#move-or-copy-a-file

- `deleteFile(folder, fileName)`
  - Deletes a file from a folder.
  - Example: `SP.deleteFile("/sites/site/Shared Documents/MyFolder", "file.txt")`

- `getFileContent(filePath, params)`
  - Downloads raw file content.
  - Example: `SP.getFileContent("/sites/site/Shared Documents/MyFolder/file.txt")`

- `getFileByPath(filePath, params)`
  - Retrieves file metadata by server-relative URL.
  - Example: `SP.getFileByPath("/sites/site/Shared Documents/MyFolder/file.txt")`

- `getFileProperties(filePath, params)`
  - Retrieves file property values.
  - Example: `SP.getFileProperties("/sites/site/Shared Documents/MyFolder/file.txt")`

- `getFileById(id, params)`
  - Retrieves a file by SharePoint file ID.
  - Example: `SP.getFileById("c4f1a05a-...-1234")`

- `getFiles(folder, params)`
  - Gets folder data by server-relative URL.
  - Example: `SP.getFiles("/sites/site/Shared Documents/MyFolder")`

- `getFilesByFolder(folder, params)`
  - Retrieves the `Files` collection for a folder.
  - Example: `SP.getFilesByFolder("/sites/site/Shared Documents/MyFolder")`

### List and Item Operations

- `createItem(list, payload)`
  - Creates a new item in the specified list.
  - Payload should contain the list item fields and may require `__metadata` when using verbose REST requests.
  - Example: `SP.createItem("Tasks", { Title: "Task 1" })`
  - See Add an item to a list: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#add-an-item-to-a-list

- `updateItem(list, id, payload)`
  - Updates an item by ID.
  - Payload should include only the fields to update; the request uses `IF-MATCH: *` and `X-HTTP-Method: MERGE`.
  - Example: `SP.updateItem("Tasks", 42, { Status: "Completed" })`
  - See Update a list item: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#update-a-list-item

- `addItemAttachment(list, id, payload, fileName)`
  - Adds an attachment file to a list item.
  - Payload is the binary file data, sent to `/AttachmentFiles/add(FileName='fileName')`.
  - Example: `SP.addItemAttachment("Tasks", 42, fileBlob, "notes.txt")`
  - See Add an attachment to a list item: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#add-an-attachment-to-a-list-item

- `deleteItemAttachment(list, id, fileName)`
  - Deletes an attachment from a list item.
  - Example: `SP.deleteItemAttachment("Tasks", 42, "notes.txt")`
  - See Delete an attachment from a list item: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#delete-an-attachment-from-a-list-item

- `deleteItem(list, id)`
  - Deletes a list item by ID.
  - Uses `IF-MATCH: *` and `X-HTTP-Method: DELETE`.
  - Example: `SP.deleteItem("Tasks", 42)`
  - See Delete a list item: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#delete-a-list-item

- `getItems(list, params)`
  - Retrieves items from a list.
  - Example: `SP.getItems("Tasks", { $select: "Id,Title" })`

- `getAllItems(list, params)`
  - Retrieves all pages of list items using SharePoint pagination.
  - Example: `SP.getAllItems("Tasks", { $top: 100 })`

- `getItemsByListId(list, params)`
  - Retrieves list items by list GUID.
  - Example: `SP.getItemsByListId("00000000-0000-0000-0000-000000000000", { $select: "Title" })`

- `getAllItemsByListId(list, params)`
  - Retrieves all list items by GUID across pages.
  - Example: `SP.getAllItemsByListId("00000000-0000-0000-0000-000000000000")`

- `getPostItems(list, data, params)`
  - Executes a CAML query against a list using POST.
  - Example: `SP.getPostItems("Tasks", "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='Text'>New</Value></Eq></Where></Query></View>")`
  - See Retrieve list items with a CAML query: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#retrieve-list-items-with-a-caml-query

- `getItem(list, id, params)`
  - Gets a single list item by ID.
  - Example: `SP.getItem("Tasks", 42)`

- `getItemCount(list)`
  - Returns the total item count for the specified list.
  - Example: `SP.getItemCount("Tasks")`

- `getListCountWithFilter(listTitle, params)`
  - Returns the item count for a list using the legacy `/_vti_bin/listdata.svc/{listTitle}/$count` endpoint.
  - Use `params` to apply query filters and paging options supported by the listdata service.
  - Example: `SP.getListCountWithFilter("Tasks", { $filter: "Status eq 'Completed'" })`

- `getFields(list, params)`
  - Retrieves non-hidden, editable fields from a list.
  - Example: `SP.getFields("Tasks", { $select: "Title,InternalName" })`

- `getList(listTitle)`
  - Retrieves list metadata by title.
  - Example: `SP.getList("Tasks")`

- `getLists()`
  - Retrieves all lists on the site.
  - Example: `SP.getLists()`

- `createList(data)`
  - Creates a new custom list.
  - Required data fields: `title`, `description`.
  - Example: `SP.createList({ title: "Projects", description: "Project tracking list" })`

- `updateList(data, listId)`
  - Updates a list by its ID.
  - Example: `SP.updateList({ title: "Projects", description: "Updated" }, "00000000-0000-0000-0000-000000000000")`

- `deleteList(listId)`
  - Deletes a list by its ID.
  - Example: `SP.deleteList("00000000-0000-0000-0000-000000000000")`

- `createColumn(listName, name, type)`
  - Adds a new field to the specified list.
  - `type` is a SharePoint `FieldTypeKind` integer (default `2` = Text).
  - Example: `SP.createColumn("Tasks", "NewField", 2)`
  - See Add a field to a list: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#add-a-field-to-a-list

- `getListViews(listTitle)`
  - Retrieves views for a list.
  - Example: `SP.getListViews("Tasks")`

- `getListView(listTitle, viewId)`
  - Gets a specific view by ID.
  - Example: `SP.getListView("Tasks", "00000000-0000-0000-0000-000000000000")`

- `getListViewFields(listTitle, viewId)`
  - Retrieves the view fields defined for a view.
  - Example: `SP.getListViewFields("Tasks", "00000000-0000-0000-0000-000000000000")`

- `getListFields(listTitle, params)`
  - Retrieves all fields for a list, optionally filtered.
  - Example: `SP.getListFields("Tasks", { $filter: "Hidden eq false" })`

- `getItemVersions(list, id, params)`
  - Retrieves version history metadata for an item.
  - Example: `SP.getItemVersions("Tasks", 42)`

### Comments and Likes

- `addComment(list, id, payload)`
  - Adds a comment to a list item.
  - Payload should follow the SharePoint Comments REST schema.
  - Example: `SP.addComment("Tasks", 42, { text: "Please review." })`
  - See Comments REST API: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-comments-by-using-sharepoint-rest-api

- `deleteComment(list, id, commentID)`
  - Deletes a comment by comment ID.
  - Example: `SP.deleteComment("Tasks", 42, 10)`

- `getComments(list, id, params)`
  - Retrieves comments for an item.
  - Example: `SP.getComments("Tasks", 42)`

- `likeComment(list, id, commentID, params)`
  - Likes a specific comment.
  - Example: `SP.likeComment("Tasks", 42, 10)`

- `unlikeComment(list, id, commentID, params)`
  - Unlikes a comment.
  - Example: `SP.unlikeComment("Tasks", 42, 10)`

- `likeItem(list, id, params)`
  - Likes a list item.
  - Example: `SP.likeItem("Tasks", 42)`

- `unlikeItem(list, id, params)`
  - Unlikes a list item.
  - Example: `SP.unlikeItem("Tasks", 42)`

### Recycle Bin and Restore

- `recycleItem(list, id)`
  - Moves a list item to the recycle bin.
  - Example: `SP.recycleItem("Tasks", 42)`
  - See Recycle a list item: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints#recycle-a-list-item

- `recycleFile(folder, fileName)`
  - Moves a file to the recycle bin.
  - Example: `SP.recycleFile("/sites/site/Shared Documents/MyFolder", "file.txt")`
  - See Recycle a file: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest#recycle-a-file

- `recycleFolder(folder)`
  - Moves a folder to the recycle bin.
  - Example: `SP.recycleFolder("/sites/site/Shared Documents/MyFolder")`
  - See Recycle a folder: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest#recycle-a-folder

- `deleteRecycledItem(itemID)`
  - Permanently deletes a recycle bin item.
  - Example: `SP.deleteRecycledItem("<recycle-bin-item-id>")`
  - See Delete a recycle bin item: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints#delete-a-recycle-bin-item

- `listRecycleBin(params)`
  - Lists recycle bin items.
  - Example: `SP.listRecycleBin({ $filter: "ItemType eq 1" })`
  - See List recycle bin items: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints#list-recycle-bin-items

- `restoreRecycledItem(id)`
  - Restores a recycle bin item.
  - Example: `SP.restoreRecycledItem("<recycle-bin-item-id>")`
  - See Restore a recycle bin item: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints#restore-a-recycle-bin-item

### Users and Groups

- `getUser(account)`
  - Retrieves a user profile by account name.
  - Example: `SP.getUser("i:0#.f|membership|user@contoso.com")`

- `getUserByEmail(email)`
  - Retrieves a user profile by email.
  - Example: `SP.getUserByEmail("user@contoso.com")`

- `getUserById(id)`
  - Retrieves a SharePoint user by ID.
  - Example: `SP.getUserById(15)`

- `getMyProperties()`
  - Retrieves current user properties.
  - Example: `SP.getMyProperties()`

- `getCurrentUser()`
  - Retrieves the current user and expands membership groups.
  - Example: `SP.getCurrentUser()`

- `getUserGroups()`
  - Retrieves groups for the current user.
  - Example: `SP.getUserGroups()`

- `getSiteGroups()`
  - Retrieves all site groups.
  - Example: `SP.getSiteGroups()`

- `getSiteGroup(groupName)`
  - Retrieves a site group by name.
  - Example: `SP.getSiteGroup("Site Members")`

- `createSiteGroup(groupName)`
  - Creates a new site group.
  - Example: `SP.createSiteGroup("Project Contributors")`

- `deleteSiteGroup(groupID)`
  - Removes a site group by ID.
  - Example: `SP.deleteSiteGroup(32)`

- `ensureUser(email)`
  - Ensures a user exists in the site user information list.
  - Example: `SP.ensureUser("user@contoso.com")`

- `getUsersFromGroup(groupName)`
  - Returns users in a group.
  - Example: `SP.getUsersFromGroup("Site Members")`

- `addUserToGroup(groupId, email)`
  - Adds a user to a site group using an email login name.
  - Example: `SP.addUserToGroup(32, "user@contoso.com")`

- `removeUserFromGroup(groupId, email)`
  - Removes a user from a group by email.
  - Example: `SP.removeUserFromGroup(32, "user@contoso.com")`

### Search

- `searchItems(list, params)`
  - Executes a SharePoint search query.
  - Example: `SP.searchItems(null, { querytext: "sharepoint" })`
  - See Search REST API: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-search-and-query-with-rest

- `searchUser(query, params)`
  - Searches people picker users by query.
  - Uses the client people picker service payload format.
  - Example: `SP.searchUser("Jane Doe")`
  - See People Picker REST search: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-search-and-query-with-rest

- `searchGroup(query)`
  - Searches people picker groups.
  - Uses the same people picker search endpoint with `PrincipalType = 8`.
  - Example: `SP.searchGroup("Project")`
  - See People Picker REST search: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-search-and-query-with-rest

### Advanced / Misc

- `getSubSites()`
  - Retrieves the site’s subsites.
  - Example: `SP.getSubSites()`

- `getPage(id)`
  - Retrieves a Site Pages page by ID.
  - Example: `SP.getPage(12)`

- `getChangeHistory(list, id)`
  - Scrapes version history from the SharePoint version page and returns structured change data.
  - Example: `SP.getChangeHistory("Tasks", 42)`

---

## Notes

- Many methods require a valid `X-RequestDigest` header.
- The helper uses `this.DIGEST`, and it attempts to refresh when necessary.
- Some SharePoint operations may require elevated permissions or list/library-specific configuration.

## Example Workflow

```js
import SP from "sp-service";

async function example() {
  await SP.generateDigest();
  const list = await SP.getList("Announcements");
  const item = await SP.createItem("Announcements", { Title: "Hello" });
  console.log(list, item);
}

example();
```
