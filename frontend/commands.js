/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // Office.js 准备好了
});

/**
 * 显示任务窗格
 * @param event {Office.AddinCommands.Event}
 */
function showTaskpane(event) {
  Office.addin.showAsTaskpane();
  event.completed();
}

// 注册函数
Office.actions.associate("showTaskpane", showTaskpane);