/* eslint-disable indent */
/*
  Node Helper module for MMM-MicrosoftToDo

  Purpose: Microsoft's OAutht 2.0 Token API endpoint does not support CORS,
  therefore we cannot make AJAX calls from the browser without disabling
  webSecurity in Electron.
*/
var NodeHelper = require("node_helper");
const moment = require("moment");
const axios = require("axios");
const qs = require('qs');

module.exports = NodeHelper.create({

    start: function () {
        console.log(this.name + " helper started ...");
    },

    socketNotificationReceived: function (notification, payload) {
        if (notification === "FETCH_DATA") {
            this.fetchData(payload);
        } else {
            console.log(this.name + " - Did not process event: " + notification);
        }
    },

    fetchData: function (config) {
        console.log("fetchData() Called");

        var self = this;

        //
        // Utility function to retrieve all the tasks in a given folder
        //
        var _getTodos = function (config) {
            console.log("_getTodos() Called");
            // based on new configuration data (listId), get tasks
            var listUrl = "https://graph.microsoft.com/beta/me/outlook/taskFolders/"
                + config.listId
                + "/tasks?$top="
                + config.itemLimit
                + "&$select=assignedTo,importance,body,isReminderOn,sensitivity,subject,status,dueDateTime,recurrence,reminderDateTime"
                + "&$filter=status%20ne%20%27completed%27";
            axios.get(listUrl, {
                headers: {
                    "Authorization": "Bearer " + config.accessToken,
                    "Prefer": 'outlook.timezone="' + Intl.DateTimeFormat().resolvedOptions().timeZone + '"'
                },
            })
                .then(response => {
                    // Populate assignedTo field
                    response.data.value.forEach(item => {
                        config.avatars.forEach(avatar => {
                            let regex = new RegExp(`^${avatar.tag}`, "i");
                            if (item.subject.search(regex) != -1) {
                                item.subject = item.subject.replace(regex, "");
                                item.assignedTo = avatar.assignedTo;
                            }
                        });
                    });

                    // Sort tasks
                    switch (config.sortOrder) {
                        case "subject":
                            response.data.value.sort((a, b) => a.subject.localeCompare(b.subject));
                            break;
                        case "importance":
                            response.data.value.sort((a, b) => a.importance.localeCompare(b.importance));
                            break;
                        case "dueDate":
                            response.data.value.sort((a, b) => {
                                if (!a.dueDateTime) return 1;
                                else if (!b.dueDateTime) return -1;
                                else return a.dueDateTime.dateTime.localeCompare(b.dueDateTime.dateTime);
                            });
                            break;
                        case "reminderDate":
                            response.data.value.sort((a, b) => {
                                if (!a.reminderDateTime) return 1;
                                else if (!b.reminderDateTime) return -1;
                                else return a.reminderDateTime.dateTime.localeCompare(b.reminderDateTime.dateTime);
                            });
                            break;
                        default:
                            // no sort
                            break;
                    }
                    console.log(JSON.stringify(response.data.value,null,2))
                    self.sendSocketNotification("DATA_FETCHED_" + config.id, {
                        value: response.data.value,
                        listId: config.listId,
                        accessToken: config.accessToken,
                        accessTokenExpiry: config.accessTokenExpiry
                    });

                })
                .catch(error => {
                    console.log(this.name + " - Error while requesting tasks:");
                    console.log(error);
                    self.sendSocketNotification("FETCH_INFO_ERROR" + config.id, error);
                });
        };

        //
        // Get a new token if it has expired before we get the list of tasks
        //
        if (moment().isAfter(config.accessTokenExpiry)) {
            console.log("token has EXPIRED!!!");

            var data = {
                client_id: config.oauth2ClientId,
                scope: "offline_access user.read tasks.read",
                refresh_token: config.oauth2RefreshToken,
                grant_type: "refresh_token",
                client_secret: config.oauth2ClientSecret
            };

            // get access token and folder ID if necessary
            axios({
                method: 'POST',
                headers: { 'content-type': 'application/x-www-form-urlencoded' },
                data: qs.stringify(data),
                url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            })
                .then(response => {
                    console.log("Got Token Response");
                    config.accessToken = response.data.access_token;
                    config.accessTokenExpiry = moment().add(response.data.expires_in, 'seconds');

                    // set config to the default list if no list ID was provided
                    if (!config.listId) {
                        console.log("Retrieving folder ID");
                        axios.get("https://graph.microsoft.com/beta/me/outlook/taskFolders/?$top=200", {
                            headers: { 'Authorization': 'Bearer ' + config.accessToken },
                        })
                            .then(response => {
                                console.log("Got Folder List");

                                const found = response.data.value.find(element => element.name === config.folderName);
                                if (found) {
                                    config.listId = found.id;
                                    _getTodos(config);
                                } else {
                                    console.log("Couldn't find ", config.folderName);
                                    self.sendSocketNotification("FETCH_INFO_ERROR" + config.id, error);
                                }
                            })
                            .catch(error => {
                                console.log(this.name + " - Error while requesting task list");
                                console.log(error);
                                self.sendSocketNotification("FETCH_INFO_ERROR" + config.id, error);
                            });
                    } else {
                        _getTodos(config);
                    }

                })
                .catch(error => {
                    console.log(this.name + " - Error while requesting access token:");
                    console.log(error);
                    self.sendSocketNotification("FETCH_INFO_ERROR" + config.id, error);
                });
        } else {
            console.log("token still valid!!!");
            _getTodos(config);
        }
    }
});
