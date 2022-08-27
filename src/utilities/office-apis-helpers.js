// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const SetRuntimeVisibleHelper = (visible) => {
  let p;
  if (visible) {
    p = Office.addin.showAsTaskpane();
  } else {
    p = Office.addin.hide();
  }

  return p
    .then(() => {
      return visible;
    })
    .catch((error) => {
      return error.code;
    });
};

const SetStartupBehaviorHelper = (isStarting) => {
  if (isStarting) {
    Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  } else {
    Office.addin.setStartupBehavior(Office.StartupBehavior.none);
  }
  let g = getGlobal();
  g.isStartOnDocOpen = isStarting;
};

function updateRibbon() {
  // Update ribbon based on state tracking
  const g = getGlobal();

  Office.ribbon.requestUpdate({
    tabs: [
      {
        id: "ShareTime",
        controls: [
          {
            id: "BtnEnableAddinStart",
            enabled: !g.state.isStartOnDocOpen,
          },
          {
            id: "BtnDisableAddinStart",
            enabled: g.state.isStartOnDocOpen,
          },
          {
            id: "BtnOpenTaskpane",
            enabled: !g.state.isTaskpaneOpen,
          },
          {
            id: "BtnCloseTaskpane",
            enabled: g.state.isTaskpaneOpen,
          },
        ],
      },
    ],
  });
}

/*
    Managing the dialogs.
*/

//This will check if state is initialized, and if not, initialize it.
//Useful as there are multiple entry points that need the state and it is not clear which one will get called first.
async function ensureStateInitialized(isOfficeInitializing) {
  let g = getGlobal();
  let initValue = false;
  if (isOfficeInitializing) {
    //we are being called in response to Office Initialize
    if (g.state !== undefined) {
      if (g.state.isInitialized === false) {
        g.state.isInitialized = true;
      }
    }
    if (g.state === undefined) {
      initValue = true;
    }
  }

  if (g.state === undefined) {
    g.state = {
      isStartOnDocOpen: false,
      isSignedIn: false,
      isTaskpaneOpen: false,
      isSyncEnabled: false,
      isFirstSyncCall: true,
      isSumEnabled: false,
      isInitialized: initValue,
      updateRct: () => {},
      setTaskpaneStatus: (opened) => {
        g.state.isTaskpaneOpen = opened;
        updateRibbon();
      },
    };

    //track startup behavior
    if (g.state.isInitialized) {
      let addinState = await Office.addin.getStartupBehavior();
      if (addinState === Office.StartupBehavior.load) {
        g.state.isStartOnDocOpen = true;
      }
    }

    //track sign in status
    if (localStorage.getItem("loggedIn") === "yes") {
      g.state.isSignedIn = true;
    }
  }
  if (g.state.isInitialized) {
    updateRibbon();
  }
}
