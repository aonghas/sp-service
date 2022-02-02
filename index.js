/*****
  SP REST services helper
    Author: Aonghas Anderson
    Original Date: July 2019
    Updated: Class version Dec 2021
*****/

import Axios from "axios";

export class SharePoint {
  constructor(options) {
    this.baseUrl = (options && options.baseUrl) || process.env.VUE_APP_URL;
    this.DIGEST = "";
    this.cancelToken = null;

    this.SP = Axios.create({
      baseURL: this.baseUrl,
    });

    console.log("connected to: " + this.baseUrl);

    this.SP.post("/_api/contextinfo", null, {
      headers: {
        Accept: "application/json;odata=verbose",
      },
    })
      .then((result) => {
        console.log("Initialised DIGEST");
        this.DIGEST = result.data.d.GetContextWebInformation.FormDigestValue;
      })
      .catch((error) => {
        console.error("Could not generate digest value: ", error);
      });
  }
  createFolder(folderLocation, folderName) {
    return this.SP.post(
      `/_api/web/folders`,
      {
        ServerRelativeUrl: folderLocation + "/" + folderName,
      },
      {
        params: {},
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  async createFileInFolder(library, folder, fileName, fileContents) {
    const exists = await this.SP.get(
      `/_api/web/GetFolderByServerRelativeUrl('${library}/${folder}')/Exists`
    ).then((resp) => resp.data.value);

    if (!exists) {
      await this.createFolder(library, folder);
    }
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${library}')/RootFolder/${folder
        .toString()
        .split("/")
        .map((i) => "folders('" + i + "')")
        .join("/")}/files/add(url='${fileName}',overwrite=true)`,
      fileContents,
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  createFile(folder, fileName, fileContents) {
    return this.SP.post(
      `/_api/web/GetFolderByServerRelativeUrl('${folder}')/Files/add(url='${fileName}',overwrite=true)`,
      fileContents,
      {
        params: {
          $expand: "ListItemAllFields",
        },
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  deleteFile(folder, fileName) {
    return this.SP.post(
      `/_api/web/GetFolderByServerRelativeUrl('${folder}/${fileName}')`,
      {},
      {
        params: {},
        headers: {
          "X-HTTP-Method": "DELETE",
          "If-Match": "*",
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  createItem(list, payload) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items`,
      payload,
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  updateItem(list, id, payload) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})`,
      payload,
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  getSubSites() {
    return this.SP.get(
      `/_api/web/webinfos`,
      {
        $select: "ServerRelativeUrl,Title",
      },
      {}
    ).then((response) => {
      return response.data;
    });
  }
  addItemAttachment(list, id, payload, fileName) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/AttachmentFiles/add(FileName='${fileName}')`,
      payload,
      {
        responseType: "arraybuffer",
        headers: {
          "X-RequestDigest": this.DIGEST,
          "Content-Type": undefined,
          "X-Requested-With": "XMLHttpRequest",
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  deleteItemAttachment(list, id, fileName) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/AttachmentFiles/getByFileName('${fileName}')`,
      {},
      {
        responseType: "arraybuffer",
        headers: {
          "X-RequestDigest": this.DIGEST,
          "Content-Type": undefined,
          "X-HTTP-Method": "DELETE",
          "X-Requested-With": "XMLHttpRequest",
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  deleteItem(list, id) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})`,
      {},
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*",
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  recycleItem(list, id) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/recycle()`,
      {},
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*",
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  listRecycleBin(params) {
    return SP.get(`/_api/web/recyclebin()`, {
      params: params || {},
      headers: {
        Accept: "application/json; odata=nometadata",
      },
    }).then((response) => {
      return response.data;
    });
  }
  restoreRecycledItem(id) {
    return SP.post(
      `/_api/web/recyclebin('${id}')/restore()`,
      {},
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  createColumn(listName, name, type) {
    return this.SP.post(
      `/_api/web/lists/getByTitle('${listName}')/fields`,
      {
        FieldTypeKind: type || 2,
        Title: name,
      },
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  getFileContent(filePath, params) {
    return this.SP.get(
      `/_api/web/GetFileByServerRelativeUrl('${filePath}')/$value`,
      {
        params: params || {},
        headers: {
          // Accept: "application/octet-stream",
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  getFileByPath(filePath, params) {
    return this.SP.get(`/_api/web/GetFileByServerRelativeUrl('${filePath}')`, {
      params: params || {},
      headers: {
        // Accept: "application/octet-stream",
        "X-RequestDigest": this.DIGEST,
      },
    }).then((response) => {
      return response.data;
    });
  }
  getFileProperties(filePath, params) {
    return this.SP.get(
      `/_api/web/GetFileByServerRelativeUrl('${filePath}')/Properties`,
      {
        params: params || {},
        headers: {
          // Accept: "application/octet-stream",
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  getFileById(id, params) {
    return this.SP.get(`/_api/web/GetFileById('${id}')`, {
      params: params || {},
      headers: {
        // Accept: "application/octet-stream",
        "X-RequestDigest": this.DIGEST,
      },
    }).then((response) => {
      return response.data;
    });
  }
  getFiles(folder, params) {
    return this.SP.get(`/_api/web/GetFolderByServerRelativeUrl('${folder}')`, {
      params: params || {},
    }).then((response) => {
      return response.data;
    });
  }
  getFilesByFolder(folder, params) {
    return this.SP.get(
      `/_api/web/GetFolderByServerRelativeUrl('${folder}')/Files`,
      {
        params: params || {},
      }
    ).then((response) => {
      return response.data;
    });
  }
  getItemCount(list) {
    return this.SP.get(`/_api/web/lists/GetByTitle('${list}')/itemcount`).then(
      (response) => {
        return response.data;
      }
    );
  }
  getFields(list, params) {
    return this.SP.get(`/_api/web/lists/getbytitle('${list}')/fields`, {
      params: Object.assign({}, params, {
        $filter: "Hidden eq false and ReadOnlyField eq false",
      }),
    }).then((response) => {
      return response.data;
    });
  }
  getItems(list, params) {
    return this.SP.get(`/_api/web/lists/GetByTitle('${list}')/items`, {
      params: params || {},
      headers: {
        Accept: "application/json; odata=nometadata",
      },
    }).then((response) => {
      return response.data;
    });
  }
  getItemsByListId(list, params) {
    return this.SP.get(`/_api/web/lists(guid'${list}')/items`, {
      params: params || {},
      headers: {
        Accept: "application/json; odata=nometadata",
      },
    }).then((response) => {
      return response.data;
    });
  }
  searchItems(list, params) {
    return this.SP.get(`/_api/search/query`, {
      params: params || {},
      headers: {
        Accept: "application/json; odata=verbose",
      },
    }).then((response) => {
      return response;
    });
  }
  getPostItems(list, data, params) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/GetItems`,
      (data && {
        query: {
          __metadata: { type: "SP.CamlQuery" },
          ViewXml: data,
        },
      }) ||
        {},
      {
        params: params || {},
        headers: {
          "content-type": "application/json;odata=verbose",
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data.value;
    });
  }
  getChangeHistory(list, id) {
    return new Promise((resolve) => {
      this.getList(list).then((data) => {
        this.SP.get(`/_layouts/15/Versions.aspx?list=${data.Id}&ID=${id}`).then(
          (resp) => {
            const parser = new DOMParser();
            const s = new XMLSerializer();
            const xmlString = resp.data;
            const doc = parser.parseFromString(xmlString, "text/html");
            const table = doc.querySelector(".ms-settingsframe");

            const items = table.querySelector("tbody").children;

            const versionArray = [];

            items.forEach((item) => {
              const version = {};
              if (item.children.length == 3 && item.querySelector("td")) {
                const author = item.children[2];

                versionArray.push({
                  changes: [],
                  author: {
                    id: parseInt(
                      author
                        .querySelector(".ms-subtleLink")
                        .getAttribute("href")
                        .split("ID=")[1]
                    ),
                    name: author.textContent.replace(/[\n\t]+/g, ""),
                    email: author
                      .querySelector(".ms-imnSpan > a > span > img")
                      .getAttribute("sip"),
                  },
                  date: new Date(
                    item
                      .querySelector(".ms-vb-title")
                      .textContent.replace(/[\n\t]+/g, "")
                  ).toISOString(),
                  versionId: parseInt(
                    item
                      .querySelector(".ms-vb-title > table")
                      .getAttribute("verid")
                  ),
                  version: parseFloat(item.querySelector(".ms-vb2").innerText),
                });
              } else if (
                item.children.length == 2 &&
                item.querySelector("td") &&
                item.querySelector("tbody")
              ) {
                const rows = item.querySelector("tbody").children;
                const versionId = rows[0].id.match(/(\d+)/g)[0];
                const indexToAdd = versionArray.findIndex(
                  (v) => v.versionId == versionId
                );

                const changes = [];

                rows.forEach((change) => {
                  const previousValue =
                    (change.getAttribute("title") &&
                      change.getAttribute("title").split("Previous Value: ")) ||
                    [];
                  changes.push({
                    id: change.id,
                    field:
                      change.querySelector(".ms-propertysheet") &&
                      change
                        .querySelector(".ms-propertysheet")
                        .innerText.replace(/[\n\t]+/g, "")
                        .trim(),
                    previousValue: previousValue[previousValue.length - 1],
                    value:
                      change.querySelector(".ms-vb") &&
                      change
                        .querySelector(".ms-vb")
                        .innerText.replace(/[\n\t]+/g, ""),
                  });
                });

                versionArray[indexToAdd].changes = changes;
              }
            });

            resolve(versionArray);
          }
        );
      });
    });
  }
  getItemVersions(list, id, params) {
    return this.SP.get(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/versions`,
      {
        params: params || {},
      }
    ).then((response) => {
      return response.data;
    });
  }
  getItem(list, id, params) {
    return this.SP.get(`/_api/web/lists/GetByTitle('${list}')/items(${id})`, {
      params: params || {},
    }).then((response) => {
      return response.data;
    });
  }
  getPage(id) {
    return this.SP.get(`/_api/sitepages/pages(${id})`).then((response) => {
      return response.data;
    });
  }
  getLists() {
    return this.SP.get("/_api/web/lists").then((response) => {
      return response.data;
    });
  }
  getList(listTitle) {
    return this.SP.get(`/_api/web/lists/GetByTitle('${listTitle}')`).then(
      (response) => {
        return response.data;
      }
    );
  }
  getListFields(listTitle, params) {
    return this.SP.get(`/_api/web/lists/GetByTitle('${listTitle}')/fields`, {
      params: params || {},
    }).then((response) => {
      return response.data;
    });
  }
  createList(data) {
    return this.SP.post(
      "/_api/web/lists",
      {
        Title: data.title,
        Description: data.description,
        ContentTypesEnabled: true,
        AllowContentTypes: true,
        BaseTemplate: 100,
      },
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  updateList(data, listId) {
    return this.SP.post(
      "/_api/web/lists('" + listId + "')",
      {
        Title: data.title,
        Description: data.description,
        ContentTypesEnabled: true,
        AllowContentTypes: true,
        BaseTemplate: 100,
      },
      {
        headers: {
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  deleteList(listId) {
    return this.SP.post(
      `/_api/web/lists('${listId}')`,
      {},
      {
        headers: {
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  getUser(account) {
    return this.SP.get(
      "/_api/sp.userprofiles.peoplemanager/getpropertiesfor(@v)",
      {
        params: {
          "@v": `'${account}'`,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  getUserByEmail(email) {
    return this.SP.get(
      "/_api/sp.userprofiles.peoplemanager/getpropertiesfor(@v)",
      {
        params: {
          "@v": `%27i:0%23.f%7Cmembership%7C${email}%27`,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  getUserById(id) {
    return this.SP.get(`/_api/web/getuserbyid(${id})`).then((response) => {
      return response.data;
    });
  }
  getMyProperties() {
    return this.SP.get(
      "/_api/SP.UserProfiles.PeopleManager/GetMyProperties/UserProfileProperties"
    ).then((response) => {
      return response.data;
    });
  }
  getCurrentUser() {
    return this.SP.get("_api/web/currentuser", {
      params: {
        $expand: "groups",
      },
    }).then((response) => {
      return response.data;
    });
  }
  getUserGroups() {
    return this.SP.get("_api/web/currentuser", {
      params: {
        $expand: "groups",
      },
    }).then((response) => {
      return response.data.Groups;
    });
  }
  ensureUser(email) {
    return this.SP.post(
      `/_api/web/ensureuser('${email}')`,
      {},
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  addComment(list, id, payload) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments()`,
      payload,
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  deleteComment(list, id, commentID) {
    return this.SP.delete(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments(${commentID})`,
      {
        headers: {
          Accept: "application/json; odata=nometadata",
          "X-RequestDigest": this.DIGEST,
          "If-Match": "*",
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  getComments(list, id, params) {
    return this.SP.get(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments()`,
      {
        params: params || {},
      }
    ).then((response) => {
      return response.data;
    });
  }
  likeComment(list, id, commentID, params) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments(${commentID})/like`,
      {
        params: params || {},
      }
    ).then((response) => {
      return response.data;
    });
  }
  unlikeComment(list, id, commentID, params) {
    return this.SP.post(
      `/_api/web/lists/GetByTitle('${list}')/items(${id})/Comments(${commentID})/unlike`,
      {
        params: params || {},
      }
    ).then((response) => {
      return response.data;
    });
  }
  searchUser(query, params) {
    if (this.cancelToken) {
      this.cancelToken.cancel();
    }
    this.cancelToken = Axios.CancelToken.source();
    return this.SP.post(
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
          QueryString: query,
          //'Required':false,
          //'SharePointGroupID':null,
          //'UrlZone':null,
          //'UrlZoneSpecified':false,
          //'Web':null,
          //'WebApplicationID':null
        },
      },
      {
        cancelToken: this.cancelToken.token,
        params: params || {},
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  searchGroup(query) {
    return this.SP.post(
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
          PrincipalType: 8,
          QueryString: query,
          //'Required':false,
          //'SharePointGroupID':null,
          //'UrlZone':null,
          //'UrlZoneSpecified':false,
          //'Web':null,
          //'WebApplicationID':null
        },
      },
      {
        headers: {
          "X-RequestDigest": DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  getUsersFromGroup(groupName) {
    return this.SP.get(
      `/_api/Web/SiteGroups/GetByName('${groupName}')/users`
    ).then((response) => {
      return response.data;
    });
  }
  addUserToGroup(groupId, email) {
    return this.SP.post(
      `/_api/Web/SiteGroups(${groupId})/users`,
      {
        LoginName: `i:0#.f|membership|${email}`,
      },
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
  removeUserFromGroup(groupId, email) {
    return this.SP.post(
      `/_api/Web/SiteGroups(${groupId})/users/getByEmail('${email}')`,
      {},
      {
        headers: {
          "X-RequestDigest": this.DIGEST,
          "X-HTTP-Method": "DELETE",
        },
      }
    ).then((response) => {
      return response.data;
    });
  }
}

export default new SharePoint();
