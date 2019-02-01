/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {
    
  };
})();

(window as any).refreshValues = (event) => {
  Office.select("binding#1", function (asyncResult) {
    if (asyncResult.status.toString() == "failed") {
        console.log(asyncResult.error);
    }
  }).setDataAsync("hello from Function file");
  event.completed();
}