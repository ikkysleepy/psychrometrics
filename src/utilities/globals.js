// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function getGlobal() {
    return typeof self !== "undefined"
      ? self
      : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
      ? global
      : undefined;
  }
  
  