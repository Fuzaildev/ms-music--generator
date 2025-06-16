/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, Excel, PowerPoint */

// Initialize auth manager
let authManager = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word || 
      info.host === Office.HostType.Excel || 
      info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Initialize auth manager
    authManager = new OfficeAuthManager();
    
    // Ensure generate button is visible
    const runButton = document.getElementById("run");
    if (runButton) {
      runButton.style.display = "flex";
      runButton.style.visibility = "visible";
      runButton.style.opacity = "1";
    }
    
    // Set up event listeners
    document.getElementById("run").onclick = generateImage;
    document.querySelector(".enhance-button").onclick = enhancePrompt;
    document.getElementById("get-more-credits").onclick = getMoreCredits;
    document.getElementById("cancel-generation").onclick = cancelGeneration;
    document.getElementById("logoutButton").onclick = handleLogout;
    document.getElementById("insert-music").onclick = insertMusicToDocument;

    // Set up duration slider
    const durationSlider = document.getElementById("duration-slider");
    const durationInput = document.getElementById("duration-input");
    
    if (durationSlider && durationInput) {
      // Update input when slider changes
      durationSlider.addEventListener('input', (e) => {
        const value = parseInt(e.target.value);
        durationInput.value = value;
      });

      // Update slider when input changes
      durationInput.addEventListener('input', (e) => {
        let value = parseInt(e.target.value);
        
        // Enforce min/max constraints
        if (value < 1) value = 1;
        if (value > 20) value = 20;
        
        durationSlider.value = value;
        durationInput.value = value;
      });

      // Validate input on blur
      durationInput.addEventListener('blur', (e) => {
        let value = parseInt(e.target.value);
        
        // If empty or invalid, reset to default
        if (isNaN(value) || value === '') {
          value = 15;
        }
        
        // Enforce min/max constraints
        if (value < 1) value = 1;
        if (value > 20) value = 20;
        
        durationSlider.value = value;
        durationInput.value = value;
      });

      // Initialize with default value
      durationInput.value = durationSlider.value;
    }

    // Initial token check
    checkTokens();
    
    // Check authentication status
    checkAuthStatus();

    // Set up interval for token checking every 3 seconds
    const tokenCheckInterval = setInterval(checkTokens, 3000);

    // Clean up interval when the page is unloaded
    window.addEventListener('unload', () => {
      if (tokenCheckInterval) {
        clearInterval(tokenCheckInterval);
      }
    });
  }
});

// Check authentication status
function checkAuthStatus() {
  console.log("Checking authentication status");
  if (!authManager) {
    console.error("Auth manager not initialized");
    return;
  }
  
  const isAuthenticated = !authManager.isTokenExpired();
  console.log("Authentication status:", isAuthenticated);
  updateAuthUI(isAuthenticated);
}

// Update authentication UI
function updateAuthUI(isAuthenticated) {
  console.log("Updating UI for authentication status:", isAuthenticated);
  const logoutButton = document.getElementById('logoutButton');
  const runButton = document.getElementById('run');
  const enhanceButton = document.querySelector('.enhance-button');
  const getMoreCreditsButton = document.getElementById('get-more-credits');
  const generateButtonLabel = runButton.querySelector('.ms-Button-label');
  
  if (isAuthenticated) {
    logoutButton.style.display = 'inline-block';
    runButton.style.display = 'flex';
    runButton.style.visibility = 'visible';
    runButton.disabled = false;
    generateButtonLabel.textContent = 'Generate Image';
    enhanceButton.disabled = false;
    getMoreCreditsButton.disabled = false;
  } else {
    logoutButton.style.display = 'none';
    runButton.style.display = 'flex';
    runButton.style.visibility = 'visible';
    runButton.disabled = false;
    generateButtonLabel.textContent = 'Sign in to Generate Image';
    enhanceButton.disabled = false;
    getMoreCreditsButton.disabled = true;
  }
}

// Add variable to track current authentication process
let currentAuthProcess = null;

// Handle login
async function handleLogin() {
  console.log("Login button clicked");
  try {
    showLoader("Authenticating...", false);
    console.log("Starting authentication process");
    
    // Store the authentication promise
    currentAuthProcess = authManager.authenticate();
    const result = await currentAuthProcess;
    currentAuthProcess = null;
    
    console.log("Authentication result:", result);
    
    if (result.success) {
      console.log("Authentication successful");
      showSuccess("Authentication successful!");
      updateAuthUI(true);
    } else {
      // Provide more specific error messages based on the error type
      let errorMessage = result.error;
      console.error("Authentication failed:", errorMessage);
      
      if (errorMessage.includes("Popup was blocked")) {
        errorMessage = "Your browser blocked the authentication popup. Please follow the instructions in the modal.";
      } else if (errorMessage.includes("Authentication cancelled by user")) {
        errorMessage = "Authentication was cancelled.";
      } else if (errorMessage.includes("Authentication timed out")) {
        errorMessage = "Authentication timed out. Please try again.";
      } else if (errorMessage.includes("State mismatch")) {
        errorMessage = "Security verification failed. Please try again.";
      }
      
      showError(`Authentication failed: ${errorMessage}`);
      updateAuthUI(false);
    }
  } catch (error) {
    console.error("Login error:", error);
    showError(`Error: ${error.message}`);
    updateAuthUI(false);
  } finally {
    // Always hide the loader and reset the current process
    hideLoader();
    currentAuthProcess = null;
  }
}

// Add cancel authentication function
async function cancelAuthentication() {
  console.log("Cancelling authentication process");
  if (authManager) {
    authManager.cancelAuth();
  }
  if (currentAuthProcess) {
    currentAuthProcess = null;
  }
  hideLoader();
  showNotification("Authentication cancelled.", "info");
  updateAuthUI(false);
}

// Add cancel credits dialog function
async function cancelCreditsDialog() {
  console.log("Cancelling credits dialog");
  hideLoader();
  showNotification("Purchase cancelled.", "info");
}

// Update showLoader function to handle authentication cancellation
function showLoader(message = "Generating your image...", isGenerating = true) {
  const loader = document.getElementById("loader");
  const loaderText = loader.querySelector(".loader-text");
  const cancelButton = document.getElementById("cancel-generation");
  const authLoadingContent = document.getElementById("auth-loading-content");
  
  loaderText.textContent = message;
  loader.classList.add("active");
  
  // Show appropriate cancel button text based on the operation
  if (cancelButton) {
    if (isGenerating) {
      cancelButton.querySelector(".ms-Button-label").textContent = "Cancel Generation";
      cancelButton.onclick = cancelGeneration;
    } else if (message.includes("Authenticating")) {
      cancelButton.querySelector(".ms-Button-label").textContent = "Cancel";
      cancelButton.onclick = cancelAuthentication;
    } else if (message.includes("purchase")) {
      cancelButton.querySelector(".ms-Button-label").textContent = "Cancel";
      cancelButton.onclick = cancelCreditsDialog;
    } else {
      cancelButton.querySelector(".ms-Button-label").textContent = "Cancel";
      cancelButton.onclick = cancelCreditsDialog;
    }
  }
  
  // Handle authentication loading experience
  if (!isGenerating && message.includes("Authenticating")) {
    // Show the auth loading content
    authLoadingContent.style.display = "block";
    loaderText.style.display = "none";
    
    // Start the animated loading steps
    startAuthLoadingAnimation();
  } else {
    // Hide the auth loading content for other operations
    authLoadingContent.style.display = "none";
    loaderText.style.display = "block";
  }
}

// Function to animate the authentication loading steps
function startAuthLoadingAnimation() {
  const steps = document.querySelectorAll('.auth-loading-step');
  const dots = document.querySelectorAll('.auth-progress-dot');
  let currentStep = 0;
  
  // Clear any existing interval
  if (window.authLoadingInterval) {
    clearInterval(window.authLoadingInterval);
    window.authLoadingInterval = null; // Ensure it's reset
  }
  
  // Function to update the active step
  function updateActiveStep() {
    // Remove active class from all steps and dots
    steps.forEach(step => step.classList.remove('active'));
    dots.forEach(dot => dot.classList.remove('active'));
    
    // Add active class to current step and dot
    if (currentStep < steps.length) {
      steps[currentStep].classList.add('active');
      dots[currentStep].classList.add('active');
      currentStep++;
      
      // If this was the last step, clear the interval
      if (currentStep >= steps.length) {
        if (window.authLoadingInterval) {
          clearInterval(window.authLoadingInterval);
          window.authLoadingInterval = null;
          console.log("Auth animation finished and stopped.");
        }
      }
    } else {
      // This part should ideally not be reached if interval is cleared correctly
      // but as a fallback, clear interval here too
      if (window.authLoadingInterval) {
        clearInterval(window.authLoadingInterval);
        window.authLoadingInterval = null;
      }
    }
  }
  
  // Set initial active step
  updateActiveStep();
  
  // Set interval to change steps every 10 seconds (for 50 second total auth time)
  window.authLoadingInterval = setInterval(updateActiveStep, 10000);
}

// Update hideLoader function to clean up auth loading animation
function hideLoader() {
  const loader = document.getElementById("loader");
  const authLoadingContent = document.getElementById("auth-loading-content");
  
  loader.classList.remove("active");
  
  // Reset auth loading content
  if (authLoadingContent) {
    authLoadingContent.style.display = "none";
    
    // Reset all steps
    const steps = document.querySelectorAll('.auth-loading-step');
    steps.forEach(step => step.classList.remove('active'));
    steps[0].classList.add('active');
  }
  
  // Clear any existing interval
  if (window.authLoadingInterval) {
    clearInterval(window.authLoadingInterval);
    window.authLoadingInterval = null;
  }
  
  // Reset the cancel button to default state
  const cancelButton = document.getElementById("cancel-generation");
  if (cancelButton) {
    cancelButton.querySelector(".ms-Button-label").textContent = "Cancel Generation";
    cancelButton.onclick = cancelGeneration;
  }
  
  // Reset the current controller when hiding loader
  currentGenerationController = null;
}

// Add variable to track the current generation request
let currentGenerationController = null;

// Add variable to track purchase polling interval
let purchaseCheckInterval = null;

/**
 * Check if a user is premium
 * @param {number} userId - The user ID to check premium status for
 * @returns {Promise<boolean>} Whether the user is premium
 */
async function checkPremiumStatus(userId) {
    try {
        if (!userId) {
            throw new Error('User ID is required');
        }
        
        console.log('🔍 Checking premium status for user:', userId);
        const response = await fetch(`https://multiplewords.com/api/account/user_settings/${userId}`);

        if (!response.ok) {
            console.error('❌ Failed to get premium status:', response.status, response.statusText);
            throw new Error(`Failed to get premium status: ${response.status}`);
        }

        const data = await response.json();
        const userRecord = data?.user_info?.find(user => user.user_id === parseInt(userId));
        const isPremium = userRecord?.is_user_paid || false;

        console.log('✅ Premium status retrieved successfully:', isPremium);
        return isPremium;
    } catch (error) {
        console.error('❌ Error checking premium status:', error);
        throw error;
    }
}

/**
 * Get the premium purchase URL
 * @param {number} userId - The user ID to use in the URL
 * @returns {string} The premium purchase URL
 */
function getPremiumPurchaseUrl(userId) {
    return `https://saifs.ai/canva_pricing/${userId}/16`;
}

// Check tokens function for music generation
async function checkTokens() {
  const tokenDisplay = document.getElementById("token-display");
  const generateButton = document.getElementById("run");
  const creditsFooter = document.getElementById("credits-footer");

  console.log("🔍 checkTokens called - DOM elements check:", {
    tokenDisplay: !!tokenDisplay,
    generateButton: !!generateButton,
    creditsFooter: !!creditsFooter,
    tokenDisplayElement: tokenDisplay,
    currentTokenText: tokenDisplay?.textContent
  });

  try {
    // Check authentication status first
    if (!authManager || authManager.isTokenExpired()) {
      console.log("User is not authenticated. Hiding credits footer.");
      
      // Hide the entire credits footer when not authenticated
      if (creditsFooter) {
        creditsFooter.style.display = "none";
        console.log("✅ Credits footer hidden for non-authenticated user");
      }
      
      // Clear token display when not authenticated
      if (tokenDisplay) {
        tokenDisplay.textContent = '';
        tokenDisplay.classList.remove("premium");
      }
      
      // Set appropriate button text for non-authenticated users
      if (generateButton) {
        const generateButtonLabel = generateButton.querySelector('.ms-Button-label');
        generateButton.disabled = false;
        generateButton.classList.remove('disabled');
        generateButtonLabel.textContent = 'Sign in to Generate Music';
      }
      return; // Exit the function early
    }

    // Show the credits footer when authenticated
    if (creditsFooter) {
      creditsFooter.style.display = "flex";
      console.log("✅ Credits footer shown for authenticated user");
    }

    // Get user ID from auth manager only if authenticated
    const userId = authManager.getUserId();
    if (!userId) {
      console.warn("Authenticated user has no ID available. Check auth flow.");
      if (tokenDisplay) {
        tokenDisplay.textContent = "Error";
        tokenDisplay.classList.remove("premium");
      }
      if (generateButton) {
        generateButton.disabled = true;
        generateButton.classList.add('disabled');
        generateButton.querySelector(".ms-Button-label").textContent = "Authentication Error";
      }
      return;
    }

    console.log("Checking tokens for authenticated user:", userId);

    // First check premium status
    const isPremium = await checkPremiumStatus(userId);

    const response = await fetch(`https://shorts.multiplewords.com/api/tokens_left/get/${userId}`, {
      method: "GET"
    });

    console.log("Token check API response status:", response.status);

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    console.log("Token check full response:", data);
    console.log("Credits object:", data.credits);

    console.log("Is premium user?", isPremium);

    if (isPremium) {
      // Handle premium user
      console.log("Handling premium user display");
      if (tokenDisplay) {
        console.log("🔄 Setting token display to ∞ for premium user");
        tokenDisplay.textContent = "∞";
        tokenDisplay.classList.add("premium");
        console.log("✅ Token display updated:", {
          textContent: tokenDisplay.textContent,
          classList: tokenDisplay.classList.toString(),
          style: tokenDisplay.style.cssText
        });
      }
      
      // Ensure button is enabled for premium users, regardless of token count
      if (generateButton) {
        generateButton.disabled = false;
        generateButton.classList.remove('disabled');
        generateButton.querySelector(".ms-Button-label").textContent = "Generate Music";
        console.log("✅ Generate button updated for premium user");
      }
    } else {
      // Handle regular user
      const tokenCount = data.credits && typeof data.credits.videos !== 'undefined' ? data.credits.videos : 0;
      console.log("Regular user token count:", tokenCount);
      if (tokenDisplay) {
        tokenDisplay.textContent = tokenCount;
        tokenDisplay.classList.remove("premium");
        console.log("✅ Token display updated for regular user:", tokenCount);
      }

      // Disable/enable generate button based on token availability
      if (tokenCount <= 0) {
        console.log("No tokens available, disabling generate button");
        if (generateButton) {
          generateButton.disabled = true;
          generateButton.classList.add('disabled');
          generateButton.querySelector(".ms-Button-label").textContent = "No Credits Available";
        }
      } else {
        console.log("Tokens available, enabling generate button");
        if (generateButton) {
          generateButton.disabled = false;
          generateButton.classList.remove('disabled');
          generateButton.querySelector(".ms-Button-label").textContent = "Generate Music";
        }
      }
    }
  } catch (error) {
    console.error("Detailed error in checkTokens:", {
      error: error,
      message: error.message,
      stack: error.stack
    });
    if (tokenDisplay) {
      tokenDisplay.textContent = "Error";
      tokenDisplay.classList.remove("premium");
    }
    // Disable button on error and show appropriate message
    if (generateButton) {
      generateButton.disabled = true;
      generateButton.classList.add('disabled');
      generateButton.querySelector(".ms-Button-label").textContent = "Error Checking Credits";
    }
  }
}

// Add cancel generation function
async function cancelGeneration() {
  if (currentGenerationController) {
    // Abort the current request
    currentGenerationController.abort();
    currentGenerationController = null;
    
    // Hide loader and show notification
    hideLoader();
    showNotification("Image generation cancelled.", "success");
    
    // Re-enable generate button
    const generateButton = document.getElementById("run");
    generateButton.disabled = false;
  }
}

// Update generateImage function to handle music generation
async function generateImage() {
  try {
    // Check if user is authenticated
    if (authManager.isTokenExpired()) {
      // If not authenticated, handle login first
      showLoader("Authenticating...", false);
      const authResult = await authManager.authenticate();
      
      if (!authResult.success) {
        hideLoader();
        showError(`Authentication failed: ${authResult.error}`);
        return;
      }
      
      // Authentication successful, continue with music generation
      hideLoader();
      showSuccess("Authentication successful!");
      updateAuthUI(true);
    }
    
    // Now proceed with music generation
    const userId = authManager.getUserId();
    
    if (!userId) {
      showError("User ID not available. Please try logging in again.");
      return;
    }
    
    const isPremium = await checkPremiumStatus(userId);
    
    const promptText = document.querySelector(".input-field").value.trim();
    const categorySelect = document.getElementById("music-category-select");
    const duration = getDuration();
    
    if (!promptText) {
      showError("Please enter a prompt for the music generation.");
      return;
    }

    if (!categorySelect.value) {
      showError("Please select a music category.");
      return;
    }

    // Check tokens before generating
    console.log("Checking tokens before generation for user:", userId);
    
    const tokenResponse = await fetch(`https://shorts.multiplewords.com/api/tokens_left/get/${userId}`, {
      method: "GET"
    });

    console.log("Pre-generation token check status:", tokenResponse.status);

    if (!tokenResponse.ok) {
      throw new Error(`HTTP error! status: ${tokenResponse.status}`);
    }

    const tokenData = await tokenResponse.json();
    console.log("Pre-generation token check response:", tokenData);
    
    console.log("User premium status:", isPremium);
    console.log("Available video tokens:", tokenData.credits?.videos);
    
    if (!isPremium && (!tokenData.credits || !tokenData.credits.videos || tokenData.credits.videos < 1)) {
      console.log("Token check failed:", {
        hasCredits: !!tokenData.credits,
        videoTokens: tokenData.credits?.videos
      });
      showError("Insufficient tokens. Please get more credits to continue.");
      return;
    }

    // Show loader and disable generate button
    showLoader("Generating your music...");
    const generateButton = document.getElementById("run");
    generateButton.disabled = true;

    // Create AbortController for the request
    currentGenerationController = new AbortController();

    // Create FormData for music generation
    const musicFormData = new FormData();
    musicFormData.append('user_id', userId);
    musicFormData.append('music_category_id', categorySelect.value);
    musicFormData.append('music_description', promptText);
    musicFormData.append('music_name', promptText.substring(0, 50)); // Use first 50 chars of prompt as name
    musicFormData.append('reference_music_id', '1'); // Default reference music ID

    // Log the request
    console.log("Music generation request:", {
        user_id: userId,
        music_category_id: categorySelect.value,
        music_description: promptText,
        music_name: promptText.substring(0, 50),
        reference_music_id: '1'
    });

    try {
        const response = await fetch("https://shorts.multiplewords.com/mwvideos/api/music_prompt", {
            method: "POST",
            body: musicFormData,
            signal: currentGenerationController.signal
        });

        console.log("Music generation response status:", response.status);
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        console.log("Music generation response:", data);

        if (data.status === 1 && data.music_id) {
            console.log("Successfully received music ID:", data.music_id);
            
            // Reduce token count immediately after successful API call for non-premium users
            if (!isPremium) {
                console.log("Reducing token count for non-premium user after successful API call");
                try {
                    // Get current token count
                    const tokenResponse = await fetch(`https://shorts.multiplewords.com/api/tokens_left/get/${userId}`, {
                        method: "GET"
                    });

                    if (!tokenResponse.ok) {
                        throw new Error(`Failed to get token count: ${tokenResponse.status}`);
                    }

                    const tokenData = await tokenResponse.json();
                    console.log("Current token data:", tokenData);

                    // Update token display
                    const tokenDisplay = document.getElementById("token-display");
                    const currentTokens = tokenData.credits?.videos || 0;
                    tokenDisplay.textContent = currentTokens;
                    console.log("Updated token count:", currentTokens);

                    // Update generate button state if no tokens left
                    const generateButton = document.getElementById("run");
                    if (currentTokens <= 0) {
                        generateButton.disabled = true;
                        generateButton.classList.add('disabled');
                        generateButton.querySelector(".ms-Button-label").textContent = "No Credits Available";
                    }
                } catch (error) {
                    console.error("Error updating tokens:", error);
                    // Don't block the music generation process if token update fails
                    // Just log the error and continue
                }
            }
            
            // Update loader message
            showLoader("Checking music generation status...");
            
            // Add retry logic for checking music status
            let retryCount = 0;
            const maxRetries = 15; // Increased max retries
            const retryDelay = 5000; // Increased to 5 seconds
            const maxRetryDelay = 15000; // Maximum delay of 15 seconds

            const checkMusicStatus = async () => {
                try {
                    const response = await fetch(`https://multiplewords.com/api/check_queue_music/${data.music_id}`);
                    const data = await response.json();
                    
                    if (data.status === "completed" && data.music_url) {
                        console.log("Music generation completed successfully");
                        
                        // Insert the music into the document
                        await insertMusicToDocument(data.music_url);
                        hideLoader();
                        showSuccess("Music generated and inserted successfully!");
                        clearInterval(checkInterval);
                        return;
                    } else if (data.status === "failed") {
                        hideLoader();
                        showError("Music generation failed. Please try again.");
                        clearInterval(checkInterval);
                        return;
                    }
                    // Continue checking if still processing
                    console.log("Music still processing...");
                } catch (error) {
                    console.error("Error checking music status:", error);
                    if (retryCount >= maxRetries) {
                        hideLoader();
                        showError("Failed to check music status. Please try again.");
                        clearInterval(checkInterval);
                        return;
                    }
                    retryCount++;
                }
            };

            // Start checking music status
            await checkMusicStatus();
        } else {
            hideLoader();
            console.error("Invalid API response:", data);
            const errorMsg = data.msg || data.message || "Failed to generate music. Please try again.";
            showError(errorMsg);
        }
    } catch (error) {
        console.error("Error in music generation:", error);
        hideLoader();
        showError(error.message || "Failed to generate music. Please try again.");
    }

    // Reset button state
    generateButton.disabled = false;

  } catch (error) {
    console.error("Detailed error in generateImage:", {
      error: error,
      message: error.message,
      stack: error.stack
    });
    hideLoader();
    
    let errorMessage = "An error occurred while generating the music.";
    if (error.name === 'AbortError') {
      errorMessage = "Music generation was cancelled.";
    } else if (error.message.includes("HTTP error")) {
      errorMessage = "Failed to connect to the music generation service. Please try again later.";
    } else if (error.name === "TypeError" && error.message.includes("fetch")) {
      errorMessage = "Network error. Please check your internet connection.";
    }
    
    showError(errorMessage);
    
    // Reset button state
    const generateButton = document.getElementById("run");
    generateButton.disabled = false;
  }
}

async function insertImageToDocument(imageUrl) {
  try {
    const base64Image = await fetchImageAsBase64(imageUrl);
    
    return new Promise((resolve, reject) => {
      Office.onReady((info) => {
        try {
          switch (info.host) {
            case Office.HostType.Word:
              Word.run(async (context) => {
                const range = context.document.getSelection();
                range.insertInlinePictureFromBase64(base64Image, "Replace");
                await context.sync();
                resolve();
              });
              break;
              
            case Office.HostType.Excel:
              Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                const shape = range.worksheet.shapes.addImage(base64Image);
                shape.width = 300; // Set default width
                shape.height = 300; // Set default height
                await context.sync();
                resolve();
              });
              break;
              
            case Office.HostType.PowerPoint:
              // Use setSelectedDataAsync for PowerPoint
              Office.context.document.setSelectedDataAsync(base64Image, {
                coercionType: Office.CoercionType.Image,
                imageLeft: 100,    // Position from the left in points
                imageTop: 100,     // Position from the top in points
                imageWidth: 300,   // Width in points
                imageHeight: 300   // Height in points
              }, function(asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log('Image inserted successfully in PowerPoint');
                  resolve();
                } else {
                  console.error('Failed to insert image in PowerPoint:', asyncResult.error.message);
                  reject(new Error(asyncResult.error.message));
                }
              });
              break;
              
            default:
              reject(new Error("Unsupported Office application"));
          }
        } catch (error) {
          reject(error);
        }
      });
    });
  } catch (error) {
    console.error("Error inserting image:", error);
    throw error;
  }
}

// Helper function to create a temporary file from base64
async function createTempFileFromBase64(base64Data) {
  try {
    // Convert base64 to blob
    const byteCharacters = atob(base64Data);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: 'image/png' });

    // Create a temporary file
    const tempFile = new File([blob], 'temp_image.png', { type: 'image/png' });
    return tempFile;
  } catch (error) {
    console.error("Error creating temporary file:", error);
    throw error;
  }
}

async function fetchImageAsBase64(imageUrl) {
  try {
    console.log("Fetching image from URL:", imageUrl);
    const response = await fetch(imageUrl);
    
    if (!response.ok) {
      throw new Error(`Failed to fetch image: ${response.status} ${response.statusText}`);
    }
    
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        try {
          // Get only the base64 data part
          const base64String = reader.result.split(',')[1];
          // Verify we have valid base64 data
          if (!base64String) {
            reject(new Error("Failed to get valid base64 data from image"));
            return;
          }
          console.log("Image successfully converted to base64");
          resolve(base64String);
        } catch (error) {
          console.error("Error processing base64 data:", error);
          reject(error);
        }
      };
      reader.onerror = (error) => {
        console.error("Error converting image to base64:", error);
        reject(error);
      };
      reader.readAsDataURL(blob);
    });
  } catch (error) {
    console.error("Error in fetchImageAsBase64:", error);
    throw new Error("Failed to process the generated image: " + error.message);
  }
}

// Update enhancePrompt function to use the new API endpoint and parameters
async function enhancePrompt() {
  try {
    const textarea = document.querySelector(".input-field");
    const currentPrompt = textarea.value.trim();
    
    if (!currentPrompt) {
      showError("Please enter a basic prompt first.");
      return;
    }

    // Disable the enhance button while processing
    const enhanceButton = document.querySelector(".enhance-button");
    const originalButtonText = enhanceButton.textContent;
    enhanceButton.textContent = "Enhancing...";
    enhanceButton.disabled = true;

    // Create FormData for the API request
    const formData = new FormData();
    formData.append('music_description', currentPrompt);

    console.log('Sending request to enhance prompt:', {
      url: "https://shorts.multiplewords.com/mwvideos/api/enhance_prompt",
      prompt: currentPrompt
    });

    // Make the API call
    const response = await fetch("https://shorts.multiplewords.com/mwvideos/api/enhance_prompt", {
      method: "POST",
      body: formData,
      headers: {
        'Accept': 'application/json'
      }
    });

    // Log the raw response
    console.log('Raw response:', response);

    if (!response.ok) {
      const errorText = await response.text();
      console.error('API Error Response:', {
        status: response.status,
        statusText: response.statusText,
        body: errorText
      });
      throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
    }

    const data = await response.json();
    console.log('API Response data:', data);
    
    // Check for enhanced_prompt in different possible locations in the response
    const enhancedPrompt = data.enhanced_prompt || data.data?.enhanced_prompt || data.result?.enhanced_prompt;
    
    if (enhancedPrompt) {
      textarea.value = enhancedPrompt;
      showSuccess("Prompt enhanced successfully!");
    } else {
      console.error('Unexpected API response format:', data);
      throw new Error("Failed to enhance prompt: No enhanced prompt in response");
    }

  } catch (error) {
    console.error("Error in enhancePrompt:", error);
    showError(`Error enhancing prompt: ${error.message}`);
  } finally {
    // Reset the enhance button state
    const enhanceButton = document.querySelector(".enhance-button");
    enhanceButton.textContent = "Enhance";
    enhanceButton.disabled = false;
  }
}

// Add function to cancel purchase check
function cancelPurchaseCheck() {
  if (purchaseCheckInterval) {
    clearInterval(purchaseCheckInterval);
    purchaseCheckInterval = null;
    hideLoader();
    showNotification("Purchase check cancelled. Please refresh if you completed the purchase.", "info");
  }
}

// Update getMoreCredits function to use dialog instead of new tab
async function getMoreCredits() {
  try {
    // Check if user is authenticated
    if (authManager.isTokenExpired()) {
      showError("Please login to purchase credits");
      return;
    }
    
    // Get user ID from auth manager
    const userId = authManager.getUserId();
    if (!userId) {
      showError("User ID not available. Please try logging in again.");
      return;
    }

    // Get current token count and premium status before opening purchase page
    const [initialTokenResponse, initialPremiumStatus] = await Promise.all([
      fetch(`https://shorts.multiplewords.com/api/tokens_left/get/${userId}`),
      checkPremiumStatus(userId)
    ]);

    if (!initialTokenResponse.ok) {
      throw new Error('Failed to get initial token count');
    }

    const initialData = await initialTokenResponse.json();
    const initialTokens = initialData.credits?.videos || 0;
    
    console.log('Initial state:', {
      tokens: initialTokens,
      isPremium: initialPremiumStatus
    });
    
    // Show loader with purchase message and set it as not a generation operation
    showLoader("Processing your purchase...", false);
    
    // Open the pricing page in a dialog instead of a new window
    try {
      const result = await authManager.openCreditsDialog(userId);
      console.log("Credits dialog result:", result);
      
      // Hide the loader after the dialog is closed
      hideLoader();
      
      // Check if the purchase was successful
      if (result.success) {
        // Start polling for changes to token count
        startPurchaseCheck(userId, initialTokens, initialPremiumStatus);
      } else {
        // Dialog was cancelled or closed without a purchase
        showNotification("Purchase cancelled or completed.", "info");
      }
    } catch (error) {
      console.error("Error opening credits dialog:", error);
      showError("Failed to open credits purchase page. Please try again.");
      hideLoader();
      return;
    }
  } catch (error) {
    console.error("Error in getMoreCredits:", error);
    showError("An error occurred while processing your request. Please try again.");
    hideLoader();
  }
}

// Helper function to start checking for purchase completion
function startPurchaseCheck(userId, initialTokens, initialPremiumStatus) {
  // Clear any existing interval
  if (purchaseCheckInterval) {
    clearInterval(purchaseCheckInterval);
  }
  
  let hasUpdated = false; // Flag to track if we've already handled the update

  // Start polling for changes
  purchaseCheckInterval = setInterval(async () => {
    if (hasUpdated) return; // Skip if we've already handled the update

    try {
      // Check both token count and premium status
      const [tokenResponse, newPremiumStatus] = await Promise.all([
        fetch(`https://shorts.multiplewords.com/api/tokens_left/get/${userId}`),
        checkPremiumStatus(userId)
      ]);
      
      if (!tokenResponse.ok) {
        throw new Error('Failed to get current token count');
      }

      const tokenData = await tokenResponse.json();
      const currentTokens = tokenData.credits?.videos || 0;
      
      console.log('Checking purchase status:', {
        initialTokens,
        currentTokens,
        initialPremiumStatus,
        newPremiumStatus,
        tokenDiff: currentTokens - initialTokens
      });
      
      // Stop checking if either tokens increased or user became premium
      if (currentTokens > initialTokens || (!initialPremiumStatus && newPremiumStatus)) {
        hasUpdated = true; // Set flag to prevent multiple updates
        clearInterval(purchaseCheckInterval);
        purchaseCheckInterval = null;
        hideLoader();
        
        if (newPremiumStatus && !initialPremiumStatus) {
          showSuccess("Premium status activated successfully!");
        } else if (currentTokens > initialTokens) {
          showSuccess(`Credits added successfully! (${currentTokens - initialTokens} tokens added)`);
        }
        
        // Update token display
        await checkTokens();
      }
    } catch (error) {
      console.error("Error checking purchase status:", error);
      // Don't stop polling on error, just log it
    }
  }, 3000); // Check every 3 seconds
  
  // Set a timeout to stop checking after 5 minutes
  setTimeout(() => {
    if (purchaseCheckInterval && !hasUpdated) {
      clearInterval(purchaseCheckInterval);
      purchaseCheckInterval = null;
      hideLoader();
      showNotification("Purchase status check timed out. Please refresh if you completed the purchase.", "error");
    }
  }, 5 * 60 * 1000); // 5 minutes timeout
}

// Clean up interval when the page is unloaded
window.addEventListener('unload', () => {
  if (purchaseCheckInterval) {
    clearInterval(purchaseCheckInterval);
    purchaseCheckInterval = null;
  }
});

function showNotification(message, type) {
  const notification = document.getElementById("notification");
  notification.textContent = message;
  notification.className = `notification ${type}`;
  notification.style.display = "block";

  // Hide the notification after 3 seconds
  setTimeout(() => {
    notification.style.display = "none";
  }, 3000);
}

function showError(message) {
  showNotification(message, "error");
}

function showSuccess(message) {
  showNotification(message, "success");
}

// Handle logout
function handleLogout() {
  console.log("Logout button clicked");
  authManager.logout();
  updateAuthUI(false);
  showSuccess("Logged out successfully");
}

// Get duration value function
function getDuration() {
  const durationSlider = document.getElementById("duration-slider");
  return parseInt(durationSlider.value);
}

// Function to insert music into the document
async function insertMusicToDocument() {
    try {
        const player = document.getElementById('music-player');
        const musicUrl = player.src;
        
        if (!musicUrl) {
            showError("No music available to insert");
            return;
        }

        showLoader("Inserting music into document...");

        await Office.onReady();
        
        // Handle different Office applications
        switch (Office.context.host) {
            case Office.HostType.Word:
                await Word.run(async (context) => {
                    const range = context.document.getSelection();
                    const paragraph = range.insertParagraph("", "After");
                    
                    // Insert an audio icon or placeholder
                    paragraph.insertHtml(
                        `<p>🎵 Audio: <a href="${musicUrl}" target="_blank">Click to play</a></p>`,
                        "Replace"
                    );
                    
                    await context.sync();
                });
                break;

            case Office.HostType.PowerPoint:
                // For PowerPoint, we'll add a shape with a link
                const audioShape = {
                    type: "Text",
                    text: "🎵 Click to play audio",
                    hyperlink: musicUrl,
                    width: 200,
                    height: 50
                };
                
                Office.context.document.setSelectedDataAsync(
                    audioShape,
                    { coercionType: Office.CoercionType.Text },
                    (result) => {
                        if (result.status === Office.AsyncResultStatus.Failed) {
                            throw new Error(result.error.message);
                        }
                    }
                );
                break;

            default:
                throw new Error("This Office application is not supported for music insertion");
        }

        hideLoader();
        showSuccess("Music inserted successfully!");
    } catch (error) {
        console.error("Error inserting music:", error);
        hideLoader();
        showError("Failed to insert music: " + error.message);
    }
}
