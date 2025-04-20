const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const fs = require('fs');

// Keep a global reference of the window object
let mainWindow;

// Start the Express server
const server = require('./server.js');

function createWindow() {
    // Create the browser window
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });

    // Load the index.html file
    mainWindow.loadFile('GVM.html');

    // Open DevTools in development
    // mainWindow.webContents.openDevTools();

    // Emitted when the window is closed
    mainWindow.on('closed', function () {
        mainWindow = null;
    });
}

// This method will be called when Electron has finished initialization
app.whenReady().then(createWindow);

// Quit when all windows are closed
app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', function () {
    if (mainWindow === null) {
        createWindow();
    }
});

// Handle Excel folder opening
ipcMain.handle('open-excel-folder', () => {
    const excelPath = path.join(__dirname, 'excel_data');
    require('electron').shell.openPath(excelPath);
}); 