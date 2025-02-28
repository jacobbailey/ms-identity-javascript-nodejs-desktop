// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { protocol } = require("electron");

/**
 * CustomProtocolListener can be instantiated in order
 * to register and unregister a custom typed protocol on which
 * MSAL can listen for Auth Code reponses.
 *
 * For information on available protocol types, check the Electron
 * protcol docs: https://www.electronjs.org/docs/latest/api/protocol/
 */
class CustomProtocolListener {
  hostName;
  /**
   * Constructor
   * @param hostName - A string that represents the host name that should be listened on (i.e. 'msal' or '127.0.0.1')
   */
  constructor(hostName) {
    this.hostName = hostName; //A string that represents the host name that should be listened on (i.e. 'msal' or '127.0.0.1')
  }

  get host() {
    return this.hostName;
  }

  /**
   * Registers a custom string protocol on which the library will
   * listen for Auth Code response.
   */
  start() {
    const codePromise = new Promise((resolve, reject) => {
      protocol.registerStringProtocol(this.host, (req, callback) => {
        const requestUrl = new URL(req.url);
        const authCode = requestUrl.searchParams.get("code");
        if (authCode) {
          resolve(authCode);
        } else {
          protocol.unregisterProtocol(this.host);
          reject(new Error("No code found in URL"));
        }
      });
    });

    return codePromise;
  }

  /**
   * Unregisters a custom string protocol to stop listening for Auth Code response.
   */
  close() {
    protocol.unregisterProtocol(this.host);
  }
}

module.exports = CustomProtocolListener;
