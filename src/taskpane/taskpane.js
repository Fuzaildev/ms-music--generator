/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// Initialize auth manager
let authManager = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Initialize auth manager
    authManager = new OfficeAuthManager();
    
    // Add credits footer to the DOM
    const footer = document.createElement('div');
    footer.className = 'credits-footer';
    footer.innerHTML = `
      <div class="credits-display">
        <span>Available Credits: <span id="token-display">0</span></span>
        <button id="get-more-credits" class="ms-Button ms-Button--small">
          <span class="ms-Button-label">Get More Credits</span>
        </button>
      </div>
    `;
    document.body.appendChild(footer);
    
    // Set up event listeners
    document.getElementById("run").onclick = generateImage;
    document.querySelector(".enhance-button").onclick = enhancePrompt;
    document.getElementById("get-more-credits").onclick = getMoreCredits;
    document.getElementById("cancel-generation").onclick = cancelGeneration;
    
    // Set up auth event listeners
    document.getElementById("loginButton").onclick = handleLogin;
    document.getElementById("logoutButton").onclick = handleLogout;

    // Add styles for the credits footer
    const style = document.createElement('style');
    style.textContent = `
      .credits-footer {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        background: #f3f3f3;
        border-top: 1px solid #ddd;
        padding: 10px;
        z-index: 1000;
      }
      .credits-display {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 14px;
      }
      #token-display {
        font-weight: bold;
        color: #107C10;
      }
      #token-display.premium {
        color: #B4009E;
      }
      #get-more-credits {
        background: #0078D4;
        color: white;
        border: none;
        padding: 4px 8px;
        cursor: pointer;
        font-size: 12px;
      }
      #get-more-credits:hover {
        background: #106EBE;
      }
    `;
    document.head.appendChild(style);

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
  const authStatus = document.getElementById('auth-status');
  const loginButton = document.getElementById('loginButton');
  const logoutButton = document.getElementById('logoutButton');
  const runButton = document.getElementById('run');
  const enhanceButton = document.querySelector('.enhance-button');
  const getMoreCreditsButton = document.getElementById('get-more-credits');
  
  if (isAuthenticated) {
    authStatus.textContent = 'Authenticated';
    authStatus.className = 'ms-fontSize-m ms-fontWeight-semibold authenticated';
    loginButton.style.display = 'none';
    logoutButton.style.display = 'inline-block';
    runButton.disabled = false;
    enhanceButton.disabled = false;
    getMoreCreditsButton.disabled = false;
  } else {
    authStatus.textContent = 'Not authenticated';
    authStatus.className = 'ms-fontSize-m ms-fontWeight-semibold not-authenticated';
    loginButton.style.display = 'inline-block';
    logoutButton.style.display = 'none';
    runButton.disabled = true;
    enhanceButton.disabled = true;
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
    cancelButton.querySelector(".ms-Button-label").textContent = isGenerating ? "Cancel Generation" : "Cancel";
    cancelButton.onclick = isGenerating ? cancelGeneration : cancelAuthentication;
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
    } else {
      // If we've gone through all steps, start over
      currentStep = 0;
      steps[currentStep].classList.add('active');
      dots[currentStep].classList.add('active');
    }
  }
  
  // Set initial active step
  updateActiveStep();
  
  // Set interval to change steps every 6-7 seconds (for 30-35 second total auth time)
  window.authLoadingInterval = setInterval(updateActiveStep, 6500);
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
    return `https://saifs.ai/canva_pricing/${userId}/15`;
}

// Update checkTokens function
async function checkTokens() {
  const tokenDisplay = document.getElementById("token-display");
  const generateButton = document.getElementById("run");

  try {
    // Check authentication status first
    if (!authManager || authManager.isTokenExpired()) {
      console.log("User is not authenticated. Skipping token check.");
      tokenDisplay.textContent = "0";
      tokenDisplay.classList.remove("premium");
      generateButton.disabled = true;
      generateButton.classList.add('disabled');
      generateButton.querySelector(".ms-Button-label").textContent = "Login Required";
      return; // Exit the function early
    }

    // Get user ID from auth manager only if authenticated
    const userId = authManager.getUserId();
    if (!userId) {
      console.warn("Authenticated user has no ID available. Check auth flow.");
      tokenDisplay.textContent = "0";
      tokenDisplay.classList.remove("premium");
      generateButton.disabled = true;
      generateButton.classList.add('disabled');
      generateButton.querySelector(".ms-Button-label").textContent = "Error"; // Indicate an issue
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
      tokenDisplay.textContent = "∞";
      tokenDisplay.classList.add("premium");
      // Ensure button is enabled for premium users, regardless of token count
      generateButton.disabled = false;
      generateButton.classList.remove('disabled');
      generateButton.querySelector(".ms-Button-label").textContent = "Generate Image";
    } else {
      // Handle regular user
      const tokenCount = data.credits && typeof data.credits.videos !== 'undefined' ? data.credits.videos : 0;
      console.log("Regular user token count:", tokenCount);
      tokenDisplay.textContent = tokenCount;
      tokenDisplay.classList.remove("premium");

      // Disable/enable generate button based on token availability
      if (tokenCount <= 0) {
        console.log("No tokens available, disabling generate button");
        generateButton.disabled = true;
        generateButton.classList.add('disabled');
        generateButton.querySelector(".ms-Button-label").textContent = "No Tokens Available";
        // Don't show error here, just disable the button
        // showError("You have no tokens left. Please get more credits to continue.");
      } else {
        console.log("Tokens available, enabling generate button");
        generateButton.disabled = false;
        generateButton.classList.remove('disabled');
        generateButton.querySelector(".ms-Button-label").textContent = "Generate Image";
      }
    }
  } catch (error) {
    console.error("Detailed error in checkTokens:", {
      error: error,
      message: error.message,
      stack: error.stack
    });
    tokenDisplay.textContent = "0";
    tokenDisplay.classList.remove("premium");
    // Disable button on error and show appropriate message
    generateButton.disabled = true;
    generateButton.classList.add('disabled');
    generateButton.querySelector(".ms-Button-label").textContent = "Error Checking Tokens";
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

// Update generateImage function
async function generateImage() {
  try {
    // Check if user is authenticated
    if (authManager.isTokenExpired()) {
      showError("Please login to generate images");
      return;
    }
    
  
   
    const userId = authManager.getUserId();
    
    if (!userId) {
      showError("User ID not available. Please try logging in again.");
      return;
    }
    
    const isPremium = await checkPremiumStatus(userId);
    
    const promptText = document.querySelector(".input-field").value.trim();
    const purposeSelect = document.getElementById("image-purpose-select");
    
    if (!promptText) {
      showError("Please enter a prompt for the image generation.");
      return;
    }

    if (!purposeSelect.value) {
      showError("Please select an image purpose.");
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
    showLoader();
    const generateButton = document.getElementById("run");
    generateButton.disabled = true;

    // Create AbortController for the request
    currentGenerationController = new AbortController();

    // Create FormData for image generation
    const imageFormData = new FormData();
    imageFormData.append('isPro', isPremium ? '1' : '0');
    imageFormData.append('user_id', userId);
    imageFormData.append('prompt', promptText);

    console.log("Sending image generation request:", {
      isPro: isPremium ? '1' : '0',
      user_id: userId,
      prompt: promptText,
      endpoint: "https://shorts.multiplewords.com/mwvideos/api/generate_image"
    });

    const response = await fetch("https://shorts.multiplewords.com/mwvideos/api/generate_image", {
      method: "POST",
      body: imageFormData,
      signal: currentGenerationController.signal
    });

    console.log("Image generation response status:", response.status);
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    console.log("Image generation full response:", data);

    if (data.status === 1 && data.generated_image && data.generated_image.image_url) {
      console.log("Successfully received image URL:", data.generated_image.image_url);
      // Update loader message
      showLoader("Inserting image into document...");
      // Insert the generated image into the document
      await insertImageToDocument(data.generated_image.image_url);
      hideLoader();
      showSuccess("Image generated and inserted successfully!");
      // Update token count after successful generation
      if (!isPremium) {
        console.log("Updating tokens after generation for non-premium user");
        checkTokens();
      }
    } else {
      hideLoader();
      console.error("Invalid API response:", {
        status: data.status,
        hasGeneratedImage: !!data.generated_image,
        hasImageUrl: data.generated_image?.image_url
      });
      const errorMsg = data.msg || data.message || "Failed to generate image. Please try again.";
      showError(errorMsg);
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
    
    let errorMessage = "An error occurred while generating the image.";
    if (error.name === 'AbortError') {
      errorMessage = "Image generation was cancelled.";
    } else if (error.message.includes("HTTP error")) {
      errorMessage = "Failed to connect to the image generation service. Please try again later.";
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
    console.log("Starting image insertion for URL:", imageUrl);
    await Word.run(async (context) => {
      // Insert a paragraph at the end of the document
      const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
      
      try {
        // Get the image as base64
        const imageContentBytes = await fetchImageAsBase64(imageUrl);
        console.log("Image converted to base64 successfully");
        
        // Ensure we're at the end of the document
        const range = context.document.body.getRange(Word.RangeLocation.end);
        
        // Insert the image at the range
        range.insertInlinePictureFromBase64(imageContentBytes, Word.InsertLocation.after);
        
        await context.sync();
        console.log("Image inserted successfully");
      } catch (error) {
        console.error("Error inserting image:", error);
        throw error;
      }
    });
  } catch (error) {
    console.error("Error in insertImageToDocument:", error);
    throw new Error("Failed to insert the image into the document: " + error.message);
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

// Update enhancePrompt function to use authentication
async function enhancePrompt() {
  try {

    
    const textarea = document.querySelector(".input-field");
    const currentPrompt = textarea.value.trim();
    const purposeSelect = document.getElementById("image-purpose-select");
    
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
    formData.append('prompt', currentPrompt);

    // Make the API call
    const response = await fetch("https://shorts.multiplewords.com/mwvideos/api/enhance_ai_image_prompt", {
      method: "POST",
      body: formData
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    
    if (data.status === "success" && data.enhanced_prompt) {
      textarea.value = data.enhanced_prompt;
      showSuccess("Prompt enhanced successfully!");
    } else {
      throw new Error("Failed to enhance prompt");
    }

  } catch (error) {
    console.error("Error in enhancePrompt:", error);
    showError(`Error: ${error.message}`);
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

// Update getMoreCredits function to use new URL generator
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
    
    // Open the pricing page in a new window
    const pricingUrl = getPremiumPurchaseUrl(userId);
    window.open(pricingUrl, "_blank");
    
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

  } catch (error) {
    console.error("Error in getMoreCredits:", error);
    hideLoader();
    showError("Failed to process purchase. Please try again later.");
    
    // Clean up interval on error
    if (purchaseCheckInterval) {
      clearInterval(purchaseCheckInterval);
      purchaseCheckInterval = null;
    }
  }
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
