// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

 function btnOpenTaskpane(event) {
  SetRuntimeVisibleHelper(true);
  g.state.isTaskpaneOpen = true;
  updateRibbon();
  event.completed();
}

 function btnCloseTaskpane(event) {
  SetRuntimeVisibleHelper(false);
  g.state.isTaskpaneOpen = false;
  updateRibbon();
  event.completed();
}

function btnEnableAddinStart(event) {
  SetStartupBehaviorHelper(true);
  g.state.isStartOnDocOpen = true;
  updateRibbon();
  event.completed();
}


function btnDisableAddinStart(event) {
  SetStartupBehaviorHelper(false);
  g.state.isStartOnDocOpen = false;
  updateRibbon();

  event.completed();
}

const g = getGlobal();
  
Office.actions.associate("btnOpenTaskpane", btnOpenTaskpane);
Office.actions.associate("btnCloseTaskpane", btnCloseTaskpane);
Office.actions.associate("btnEnableAddinStart", btnEnableAddinStart);
Office.actions.associate("btnDisableAddinStart", btnDisableAddinStart);
