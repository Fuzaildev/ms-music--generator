/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, Excel, PowerPoint */
import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './taskpane.css';

// Initialize auth manager
let authManager = null;

Office.onReady((info) => {
  // Support for Word, Excel, and PowerPoint
  if (info.host === Office.HostType.Word || 
      info.host === Office.HostType.Excel || 
      info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Initialize auth manager
    authManager = new OfficeAuthManager();
    
    // Render the React app
    const container = document.getElementById('app-body');
    const root = createRoot(container);
    root.render(<App authManager={authManager} />);
  }
}); 