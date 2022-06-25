/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

require('dotenv').config();

const path = require('path');
const { app, ipcMain, BrowserWindow, ipcRenderer } = require('electron');
const { IPC_MESSAGES } = require('./constants');

const { callEndpointWithToken } = require('./fetch');
const AuthProvider = require('./AuthProvider');
const { protectedResources } = require('./authConfig');

const authProvider = new AuthProvider();
let mainWindow;

if (process.defaultApp) {
    if (process.argv.length >= 2) {
        console.log('set path');
        app.setAsDefaultProtocolClient('electron-fiddle', process.execPath, [path.resolve(process.argv[1])]);
    } else {
        app.setAsDefaultProtocolClient('electron-fiddle');
    }
}

const gotTheLock = app.requestSingleInstanceLock();

if (!gotTheLock) {
    console.log('in app');
    app.quit();
} else {
    app.on('second-instance', (event, commandLine, workingDirectory) => {
        // Someone tried to run a second instance, we should focus our window.
        if (mainWindow) {
            console.log('should be here');
            if (mainWindow.isMinimized()) mainWindow.restore();
            mainWindow.focus();
        }
        const url = getDeepLinkUrl(commandLine);

        if (url) {
            fetchedCode(url);
        }
    });

    app.whenReady().then(() => {
        createWindow();
        mainWindow.loadFile(path.join(__dirname, './index.html'));
    });
}

const getDeepLinkUrl = (argv) => {
    for (const arg of argv) {
        const value = arg;
        if (value.indexOf('electron-fiddle') !== -1) {
            return value;
        }
    }
    return null;
};

const fetchedCode = async (url) => {
    const account = await authProvider.getCodeFromUrl(url);
    await mainWindow.loadFile(path.join(__dirname, './index.html'));
    mainWindow.webContents.send(IPC_MESSAGES.SHOW_WELCOME_MESSAGE, account);
};

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: { preload: path.join(__dirname, 'preload.js') },
    });
}

app.on('window-all-closed', () => {
    app.quit();
});

// Event handlers

ipcMain.on(IPC_MESSAGES.LOGIN, async () => {
    const account = await authProvider.login();

    // await mainWindow.loadFile(path.join(__dirname, "./index.html"));
    // mainWindow.webContents.send(IPC_MESSAGES.SHOW_WELCOME_MESSAGE, account);
});

ipcMain.on(IPC_MESSAGES.LOGOUT, async () => {
    await authProvider.logout();
    await mainWindow.loadFile(path.join(__dirname, './index.html'));
});

ipcMain.on(IPC_MESSAGES.GET_PROFILE, async () => {
    const tokenRequest = {
        scopes: protectedResources.graphMe.scopes,
    };

    const token = await authProvider.getToken(tokenRequest);
    const account = authProvider.account;

    await mainWindow.loadFile(path.join(__dirname, './index.html'));

    const graphResponse = await callEndpointWithToken(protectedResources.graphMe.endpoint, token);

    mainWindow.webContents.send(IPC_MESSAGES.SHOW_WELCOME_MESSAGE, account);
    mainWindow.webContents.send(IPC_MESSAGES.SET_PROFILE, graphResponse);
});

ipcMain.on(IPC_MESSAGES.GET_MAIL, async () => {
    const tokenRequest = {
        scopes: protectedResources.graphMessages.scopes,
    };

    const token = await authProvider.getToken(tokenRequest);
    const account = authProvider.account;

    await mainWindow.loadFile(path.join(__dirname, './index.html'));

    const graphResponse = await callEndpointWithToken(protectedResources.graphMessages.endpoint, token);

    mainWindow.webContents.send(IPC_MESSAGES.SHOW_WELCOME_MESSAGE, account);
    mainWindow.webContents.send(IPC_MESSAGES.SET_MAIL, graphResponse);
});
