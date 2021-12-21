/*****
  SP REST services helper
    Author: Aonghas Anderson
    Date: July 2019
*****/
import Axios from "axios";

// hack to get the relative url for production Sharepoint API
const SP = Axios.create({
  baseURL: process.env.VUE_APP_URL
});

let DIGEST = "";

let cancelToken = null;

// const expandPerson = function(key) {
//   return key + "/Title," + key + "/EMail," + key + "/ID,";
// };

// Get digest value for Sharepoint
SP.post("/_api/contextinfo", null, {
  headers: {
    Accept: "application/json;odata=verbose"
  }
})
  .then((result) => {
    console.log("Initialised DIGEST");
    DIGEST = result.data.d.GetContextWebInformation.FormDigestValue;
  })
  .catch((error) => {
    console.error("Could not generate digest value: ", error);
  });

export default {
  createFolder(folderLocation, folderName) {
    return SP.post(
      `/_api/web/folders`,
      {
        ServerRelativeUrl: folderLocation + "/" + folderName
      },
      {
        params: {},
        headers: {
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  async createFileInFolder(library, folder, fileName, fileContents) {
    const exists = await SP.get(
      `/_api/web/GetFolderByServerRelativeUrl('${library}/${folder}')/Exists`
    ).then((resp) => resp.data.value);

    if (!exists) {
      await this.createFolder(library, folder);
    }
    return SP.post(
      `/_api/web/lists/GetByTitle('${library}')/RootFolder/${folder
        .toString()
        .split("/")
        .map((i) => "folders('" + i + "')")
        .join("/")}/files/add(url='${fileName}',overwrite=true)`,
      fileContents,
      {
        headers: {
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  createFile(folder, fileName, fileContents) {
    return SP.post(
      `/_api/web/GetFolderByServerRelativeUrl('${folder}')/Files/add(url='${fileName}',overwrite=true)`,
      fileContents,
      {
        params: {
          $expand: "ListItemAllFields"
        },
        headers: {
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  deleteFile(folder, fileName) {
    return SP.post(
      `/_api/web/GetFolderByServerRelativeUrl('${folder}/${fileName}')`,
      {},
      {
        params: {},
        headers: {
          "X-HTTP-Method": "DELETE",
          "If-Match": "*",
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  createItem(list, payload) {
    return SP.post(`/_api/web/lists/GetByTitle('${list}')/items`, payload, {
      headers: {
        "X-RequestDigest": DIGEST
      }
    }).then((response) => {
      return response.data;
    });
  },
  updateItem(list, id, payload) {
    return SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})`,
      payload,
      {
        headers: {
          "X-RequestDigest": DIGEST,
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE"
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  getSubSites() {
    return SP.get(
      `/_api/web/webinfos`,
      {
        $select: "ServerRelativeUrl,Title"
      },
      {}
    ).then((response) => {
      return response.data;
    });
  },
  addItemAttachment(list, id, payload, fileName) {
    return SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/AttachmentFiles/add(FileName='${fileName}')`,
      payload,
      {
        responseType: "arraybuffer",
        headers: {
          "X-RequestDigest": DIGEST,
          "Content-Type": undefined,
          "X-Requested-With": "XMLHttpRequest"
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  deleteItemAttachment(list, id, fileName) {
    return SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/AttachmentFiles/getByFileName('${fileName}')`,
      {},
      {
        responseType: "arraybuffer",
        headers: {
          "X-RequestDigest": DIGEST,
          "Content-Type": undefined,
          "X-HTTP-Method": "DELETE",
          "X-Requested-With": "XMLHttpRequest"
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  deleteItem(list, id) {
    return SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})`,
      {},
      {
        headers: {
          "X-RequestDigest": DIGEST,
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*"
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  createColumn(listName, name, type) {
    return SP.post(
      `/_api/web/lists/getByTitle('${listName}')/fields`,
      {
        FieldTypeKind: type || 2,
        Title: name
      },
      {
        headers: {
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  getFileContent(filePath, params) {
    return SP.get(
      `/_api/web/GetFileByServerRelativeUrl('${filePath}')/$value`,
      {
        params: params || {},
        headers: {
          // Accept: "application/octet-stream",
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  getFileByPath(filePath, params) {
    return SP.get(`/_api/web/GetFileByServerRelativeUrl('${filePath}')`, {
      params: params || {},
      headers: {
        // Accept: "application/octet-stream",
        "X-RequestDigest": DIGEST
      }
    }).then((response) => {
      return response.data;
    });
  },
  getFileProperties(filePath, params) {
    return SP.get(
      `/_api/web/GetFileByServerRelativeUrl('${filePath}')/Properties`,
      {
        params: params || {},
        headers: {
          // Accept: "application/octet-stream",
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  getFileById(id, params) {
    return SP.get(`/_api/web/GetFileById('${id}')`, {
      params: params || {},
      headers: {
        // Accept: "application/octet-stream",
        "X-RequestDigest": DIGEST
      }
    }).then((response) => {
      return response.data;
    });
  },
  getFiles(folder, params) {
    return SP.get(`/_api/web/GetFolderByServerRelativeUrl('${folder}')`, {
      params: params || {}
    }).then((response) => {
      return response.data;
    });
  },
  getFilesByFolder(folder, params) {
    return SP.get(`/_api/web/GetFolderByServerRelativeUrl('${folder}')/Files`, {
      params: params || {}
    }).then((response) => {
      return response.data;
    });
  },
  getItemCount(list) {
    return SP.get(`/_api/web/lists/GetByTitle('${list}')/itemcount`).then(
      (response) => {
        return response.data;
      }
    );
  },
  getFields(list, params) {
    return SP.get(`/_api/web/lists/getbytitle('${list}')/fields`, {
      params: Object.assign({}, params, {
        $filter: "Hidden eq false and ReadOnlyField eq false"
      })
    }).then((response) => {
      return response.data;
    });
  },
  getItems(list, params) {
    return SP.get(`/_api/web/lists/GetByTitle('${list}')/items`, {
      params: params || {},
      headers: {
        Accept: "application/json; odata=nometadata"
      }
    }).then((response) => {
      return response.data;
    });
  },
  searchItems(list, params) {
    return SP.get(`/_api/search/query`, {
      params: params || {},
      headers: {
        Accept: "application/json; odata=verbose"
      }
    }).then((response) => {
      return response;
    });
  },
  getPostItems(list, data, params) {
    return SP.post(
      `/_api/web/lists/GetByTitle('${list}')/GetItems`,
      (data && {
        query: {
          __metadata: { type: "SP.CamlQuery" },
          ViewXml: data
        }
      }) ||
        {},
      {
        params: params || {},
        headers: {
          "content-type": "application/json;odata=verbose",
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data.value;
    });
  },
  getItemVersions(list, id, params) {
    return SP.get(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/versions`,
      {
        params: params || {}
      }
    ).then((response) => {
      return response.data;
    });
  },
  getItem(list, id, params) {
    return SP.get(`/_api/web/lists/GetByTitle('${list}')/items(${id})`, {
      params: params || {}
    }).then((response) => {
      return response.data;
    });
  },
  getPage(id) {
    return SP.get(`/_api/sitepages/pages(${id})`).then((response) => {
      return response.data;
    });
  },
  getLists() {
    return SP.get("/_api/web/lists").then((response) => {
      return response.data;
    });
  },
  getList(listTitle) {
    return SP.get(`/_api/web/lists/GetByTitle('${listTitle}')`).then(
      (response) => {
        return response.data;
      }
    );
  },
  getListFields(listTitle, params) {
    return SP.get(`/_api/web/lists/GetByTitle('${listTitle}')/fields`, {
      params: params || {}
    }).then((response) => {
      return response.data;
    });
  },
  createList(data) {
    return SP.post(
      "/_api/web/lists",
      {
        Title: data.title,
        Description: data.description,
        ContentTypesEnabled: true,
        AllowContentTypes: true,
        BaseTemplate: 100
      },
      {
        headers: {
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  updateList(data, listId) {
    return SP.post(
      "/_api/web/lists('" + listId + "')",
      {
        Title: data.title,
        Description: data.description,
        ContentTypesEnabled: true,
        AllowContentTypes: true,
        BaseTemplate: 100
      },
      {
        headers: {
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  deleteList(listId) {
    return SP.post(
      `/_api/web/lists('${listId}')`,
      {},
      {
        headers: {
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  getUser(account) {
    return SP.get("/_api/sp.userprofiles.peoplemanager/getpropertiesfor(@v)", {
      params: {
        "@v": `'${account}'`
      }
    }).then((response) => {
      return response.data;
    });
  },
  getUserByEmail(email) {
    return SP.get("/_api/sp.userprofiles.peoplemanager/getpropertiesfor(@v)", {
      params: {
        "@v": `%27i:0%23.f%7Cmembership%7C${email}%27`
      }
    }).then((response) => {
      return response.data;
    });
  },
  getUserById(id) {
    return SP.get(`/_api/web/getuserbyid(${id})`).then((response) => {
      return response.data;
    });
  },
  getMyProperties() {
    return SP.get(
      "/_api/SP.UserProfiles.PeopleManager/GetMyProperties/UserProfileProperties"
    ).then((response) => {
      return response.data;
    });
  },
  getCurrentUser() {
    return SP.get("_api/web/currentuser", {
      params: {
        $expand: "groups"
      }
    }).then((response) => {
      return response.data;
    });
  },
  getUserGroups() {
    return SP.get("_api/web/currentuser", {
      params: {
        $expand: "groups"
      }
    }).then((response) => {
      return response.data.Groups;
    });
  },
  ensureUser(email) {
    return SP.post(
      `/_api/web/ensureuser('${email}')`,
      {},
      {
        headers: {
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  addComment(list, id, payload) {
    return SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments()`,
      payload,
      {
        headers: {
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  deleteComment(list, id, commentID) {
    return SP.delete(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments(${commentID})`,
      {
        headers: {
          Accept: "application/json; odata=nometadata",
          "X-RequestDigest": DIGEST,
          "If-Match": "*"
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  getComments(list, id, params) {
    return SP.get(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments()`,
      {
        params: params || {}
      }
    ).then((response) => {
      return response.data;
    });
  },
  likeComment(list, id, commentID, params) {
    return SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments(${commentID})/like`,
      {
        params: params || {}
      }
    ).then((response) => {
      return response.data;
    });
  },
  unlikeComment(list, id, commentID, params) {
    return SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments(${commentID})/unlike`,
      {
        params: params || {}
      }
    ).then((response) => {
      return response.data;
    });
  },
  searchUser(query, params) {
    if (cancelToken) {
      cancelToken.cancel();
    }
    cancelToken = Axios.CancelToken.source();
    return SP.post(
      "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser",
      {
        queryParams: {
          // __metadata: {
          //   type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters"
          // },
          AllowEmailAddresses: true,
          AllowMultipleEntities: false,
          AllUrlZones: false,
          MaximumEntitySuggestions: 50,
          PrincipalSource: 15,
          PrincipalType: 1,
          QueryString: query
          //'Required':false,
          //'SharePointGroupID':null,
          //'UrlZone':null,
          //'UrlZoneSpecified':false,
          //'Web':null,
          //'WebApplicationID':null
        }
      },
      {
        cancelToken: cancelToken.token,
        params: params || {},
        headers: {
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  getUsersFromGroup(groupName) {
    return SP.get(`/_api/Web/SiteGroups/GetByName('${groupName}')/users`).then(
      (response) => {
        return response.data;
      }
    );
  },
  addUserToGroup(groupId, email) {
    return SP.post(
      `/_api/Web/SiteGroups(${groupId})/users`,
      {
        LoginName: `i:0#.f|membership|${email}`
      },
      {
        headers: {
          "X-RequestDigest": DIGEST
        }
      }
    ).then((response) => {
      return response.data;
    });
  },
  removeUserFromGroup(groupId, email) {
    return SP.post(
      `/_api/Web/SiteGroups(${groupId})/users/getByEmail('${email}')`,
      {},
      {
        headers: {
          "X-RequestDigest": DIGEST,
          "X-HTTP-Method": "DELETE"
        }
      }
    ).then((response) => {
      return response.data;
    });
  }
};
