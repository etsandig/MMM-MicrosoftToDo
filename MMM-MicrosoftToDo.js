/* eslint-disable indent */
Module.register("MMM-MicrosoftToDo", {

  defaults: {
    oauth2ClientId: "",
    oauth2ClientSecret: "",
    oauth2RefreshToken: "",
    folderName: "Tasks",
    itemLimit: 20, // optional: limit on the number of items to show from the list
    refreshInterval: 15, // Refresh every 15 minutes
    sortOrder: "none", // 'none', 'assignedTo', 'subject', 'importance', 'dueDate', 'reminderDate';
    displayWhenEmpty: false, // optional: default will show the module even when the todo list is empty
    displayLastUpdate: false,
    displayDecorations: false, //display information under the task name
    displayAvatar: "none", //could be:  'none', 'initials', 'icon'

    avatars: [],
    list: [],
  },

  // Override dom generator.
  getDom: function () {
    const today = moment();
    const yesterday = moment(today).subtract(1, "days");
    const tomorrow = moment(today).add(1, "days");

    var wrapper = document.createElement("div");
    wrapper.className = "MSToDo";

    if (this.list != null && this.list.length != 0) {
      /* iterate through items, i.e. todo's */
      this.list.forEach((item) => {
        var row = document.createElement("div");
        row.className = "task";

        /* cell for importance indicator */
        var importanceCell = document.createElement("div");
        if (item.importance === "high") {
          importanceCell.className = "importance high iconify icons";
          importanceCell.dataset.icon = "noto-star";
        } else if (item.importance === "normal") {
          importanceCell.className = "importance normal iconify icons";
        } else {
          importanceCell.className = "importance iconify icons";
        }
        row.appendChild(importanceCell);

        var captionWrapper = document.createElement("div");
        captionWrapper.className = "caption";
        {
          /* cell for todo content */
          var todoCell = document.createElement("div");
          todoCell.className = "title";
          todoCell.innerHTML = item.subject;
          captionWrapper.appendChild(todoCell);

          if (this.config.displayDecorations) {
            var decorationsWrapper = document.createElement("div");
            decorationsWrapper.className = "decorations";
            if (item.isReminderOn) {
              var remindIconCell = document.createElement("div");
              remindIconCell.className = "iconify icons";
              remindIconCell.dataset.icon = "bx-bx-bell";
              decorationsWrapper.appendChild(remindIconCell);

              var remindTextCell = document.createElement("div");
              remindTextCell.className = "text";
              var rDateTime = moment(item.reminderDateTime.dateTime);
              if (rDateTime.isSame(today, "day")) {
                remindTextCell.innerHTML = "Today";
              } else if (rDateTime.isSame(tomorrow, "day")) {
                remindTextCell.innerHTML = "Today";
              } else {
                remindTextCell.innerHTML = rDateTime.format("ddd, MMM D");
              }
              decorationsWrapper.appendChild(remindTextCell);
            }

            if (item.recurrence) {
              var recurIconCell = document.createElement("div");
              recurIconCell.className = "iconify icons";
              recurIconCell.dataset.icon = "entypo-cycle";
              decorationsWrapper.appendChild(recurIconCell);

              var recurTextCell = document.createElement("div");
              recurTextCell.className = "text";
              if (item.recurrence.pattern.interval != 1) {
                //Go simple for now and avoid complex cases
                const array1 = ["daily", "weekly", "monthly", "yearly"];
                const array2 = ["days", "weeks", "months", "years"];
                const found = array1.indexOf(item.recurrence.pattern.type);
                recurTextCell.innerHTML = "Every "
                  + item.recurrence.pattern.interval
                  + array2[found];
              } else {
                recurTextCell.innerHTML = item.recurrence.pattern.type;
              }
              decorationsWrapper.appendChild(recurTextCell);
            }

            if (item.body && item.body.content != "") {
              var bodyIconCell = document.createElement("div");
              bodyIconCell.className = "iconify icons";
              bodyIconCell.dataset.icon = "bx-bx-note";
              decorationsWrapper.appendChild(bodyIconCell);
            }
            captionWrapper.appendChild(decorationsWrapper);
          }
        }
        row.appendChild(captionWrapper);

        /* cell for due date */
        var dueDateCell = document.createElement("div");
        dueDateCell.className = "dueDate";

        //var today = moment("2020-04-04");
        if (item.dueDateTime != null) {
          var date = moment(item.dueDateTime.dateTime);

          if (date.isSame(today, "day")) {
            dueDateCell.innerHTML = this.translate("TODAY");
            dueDateCell.classList.add("today");
          } else if (date.isSame(yesterday, "day")) {
            dueDateCell.innerHTML = this.translate("YESTERDAY");
            dueDateCell.classList.add("overdue");
          } else if (date.isSame(tomorrow, "day")) {
            dueDateCell.innerHTML = this.translate("TOMORROW");
            dueDateCell.classList.add("tomorrow");
          } else {
            if (date.isSame(today, "year")) {
              dueDateCell.innerHTML = date.format("MMM Do");
            } else {
              dueDateCell.innerHTML = date.format("MMM YYYY");
            }
            if (date.isBefore(yesterday)) {
              dueDateCell.classList.add("overdue");
            }
          }
        } else {
          dueDateCell.innerHTML = "";
        }
        row.appendChild(dueDateCell);

        /* cell for assignee avatar */
        var avatarCell = document.createElement("div");
        if (this.config.displayAvatar === "icons") {
          var avatar = this.config.avatars
            .filter(obj => obj.assignedTo === item.assignedTo)
            .map(obj => obj.icon);
          avatarCell.className = "avatar iconify icons";
          avatarCell.dataset.icon = avatar[0];
        } else if (this.config.displayAvatar === "initials") {
          var fillerCell = document.createElement("div");
          fillerCell.className = "initials";
          fillerCell.innerHTML = "HT";
          avatarCell.appendChild(fillerCell);
          avatarCell.className = "avatar";
        } else {
          avatarCell.className = "avatar none";
        }
        row.appendChild(avatarCell);

        // Create fade effect by MichMich (MIT)
        if (this.config.fade && this.config.fadePoint < 1) {
          if (this.config.fadePoint < 0) {
            this.config.fadePoint = 0;
          }
          var startingPoint = this.tasks.items.length * this.config.fadePoint;
          var steps = this.tasks.items.length - startingPoint;
          if (i >= startingPoint) {
            var currentStep = i - startingPoint;
            row.style.opacity = 1 - (1 / steps * currentStep);
          }
        }
        // End Create fade effect by MichMich (MIT)

        wrapper.appendChild(row);
      });
    }
    else {
      // otherwise indicate that there are no list entries
      var row = document.createElement("div");
      row.className = "task";
      row.className = "taskCount0";
      row.innerHTML = this.translate("NO_ENTRIES");
      wrapper.appendChild(row);
    }

    // display the update time at the end, if defined so by the user config
    if (this.config.displayLastUpdate && this.lastUpdate != null) {
      var updateinfo = document.createElement("div");
      updateinfo.className = "lastUpdated";
      var now = moment();
      var value = moment.duration(this.lastUpdate.diff(now));
//      console.log(this.identifier + ": "
//        + this.lastUpdate.format("HH:mm:ss")
//        + " -- " + now.format("HH:mm:ss")
//        + " -- " + value.hours() + ":"
//        + value.minutes() + ":"
//        + value.seconds()
//      );
      updateinfo.innerHTML = "Last updated " + moment.duration(this.lastUpdate.diff(moment())).humanize() + " ago";
      wrapper.appendChild(updateinfo);
    }

    return wrapper;
  },

  getStyles: function () {
    return ["MMM-MicrosoftToDo.css"];
  },

  getTranslations: function () {
    return {
      en: "translations/en.js",
      de: "translations/de.js",
      fr: "translations/fr.js",
    };
  },

  getScripts: function () {
    return ["moment.js", "//code.iconify.design/1/1.0.5/iconify.min.js"];
  },

  socketNotificationReceived: function (notification, payload) {
    if (notification === ("DATA_FETCHED_" + this.config.id)) {
      //Save connection data in config
      this.config.accessToken = payload.accessToken;
      this.config.accessTokenExpiry = payload.accessTokenExpiry;
      this.config.listId = payload.listId;

      this.list = payload.value;

      if (this.config.displayLastUpdate) {

        this.lastUpdate = moment();
        //console.log(this.identifier + ": socketNotificationReceived() last update: " + this.lastUpdate.format("HH:mm:ss"));
      }

      // check if module should be hidden according to list size and the module's configuration
      if (this.config.displayWhenEmpty) {
        if (this.hidden) this.show();
      } else {
        if (this.list.length == 0) {
          if (!this.hidden) this.hide();
        } else {
          if (this.hidden) this.show();
        }
      }

    } else if (notification === ("FETCH_INFO_ERROR" + this.config.id)) {
      console.log(payload);
    } else {
    }

    this.updateDom();
  },

  notificationReceived: function (notification, payload, sender) {
    var self = this;

    if (notification === "DOM_OBJECTS_CREATED") {
      console.log("DOM_OBJECTS_CREATED");
    }

    if (notification === "CLOCK_MINUTE") {
      //console.log(this.identifier + ": Minute Last Updated");
      if (this.config.displayLastUpdate && this.lastUpdate != null) {
        var updateinfo = document.getElementsByClassName("lastUpdated");
        updateinfo[0].innerHTML = "Last updated " + moment.duration(this.lastUpdate.diff(moment())).humanize() + " ago";
        this.updateDom();
      }
    }
  },

  start: function () {
    // copy module object to be accessible in callbacks
    var self = this;

    // start with empty list that shows loading indicator
    self.list = [{ subject: this.translate("LOADING_ENTRIES") }];

    // in case there are multiple instances of this module, ensure the responses from node_helper are mapped to the correct module
    self.config.id = this.identifier;

    // make the token expired
    self.config.accessTokenExpiry = moment();

    var refreshFunction = function () {
      //console.log(self.identifier + ": refreshFunction() Called.");
      self.sendSocketNotification("FETCH_DATA", self.config);
    };

    refreshFunction();
    setInterval(refreshFunction, self.config.refreshInterval * 1000 * 60);
  },

});
