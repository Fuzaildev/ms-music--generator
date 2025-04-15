// auth-manager.js
class OfficeAuthManager {
    constructor() {
        this.clientId = 'HGn3aX2z6aOFhikeyc2MXLcrEdxw6apkZo2W0MiW';
        this.redirectUri = 'https://multiplewords.com/oauth/office/callback/';
        this.authEndpoint = 'https://multiplewords.com/oauth/office/authorize/';
        this.tokenEndpoint = 'https://multiplewords.com/oauth/office/token/';
        this.dialog = null;
        this.state = null;
        this.platform = null;
        this.pollInterval = null;
        this.authWindow = null;
        this.pollDelay = 0; // Changed to 0 for immediate polling
        this.maxAboutBlankAttempts = 20; // Maximum number of attempts to wait for about:blank to change
        this.authCompleted = false; // Flag to track if authentication has completed
        this.isCancelled = false; // Flag to track if authentication was cancelled
        this.codeUsed = false; // Flag to track if the authorization code has been used
        
        console.log("🔐 OfficeAuthManager initialized with:", {
            clientId: this.clientId,
            redirectUri: this.redirectUri,
            authEndpoint: this.authEndpoint,
            tokenEndpoint: this.tokenEndpoint
        });
    }

    // Generate random state for security
    generateState() {
        const array = new Uint32Array(8);
        window.crypto.getRandomValues(array);
        const state = Array.from(array, dec => ('0' + dec.toString(16)).substr(-2)).join('');
        console.log("🔑 Generated state parameter:", state);
        return state;
    }

    // Detect platform (web or desktop)
    detectPlatform() {
        const platform = Office.context.platform === 'OfficeOnline' ? 'web' : 'desktop';
        console.log("🌐 Detected platform:", platform, "Office.context.platform:", Office.context.platform);
        return platform;
    }

    // Initialize authentication flow
    async startAuthFlow() {
        console.log("🚀 Starting authentication flow");
        this.platform = this.detectPlatform();
        
        // Generate and store state
        this.state = this.generateState();
        sessionStorage.setItem('oauth_state', this.state);
        console.log("💾 Stored state in sessionStorage:", this.state);
        
        // Construct authorization URL with all parameters
        const authUrl = new URL(this.authEndpoint);
        authUrl.searchParams.append('client_id', this.clientId);
        authUrl.searchParams.append('redirect_uri', this.redirectUri);
        authUrl.searchParams.append('response_type', 'code');
        authUrl.searchParams.append('state', this.state);
        authUrl.searchParams.append('platform', this.platform);
        authUrl.searchParams.append('scope', 'read write');
        
        const authUrlString = authUrl.toString();
        console.log("🔗 Constructed auth URL:", authUrlString);
        
        // Use the same popup window approach for both web and desktop
        console.log(`Using popup window authentication flow for ${this.platform} platform`);
        return this.handlePopupAuth(authUrlString);
    }

    // Handle authentication using popup window for both web and desktop platforms
    async handlePopupAuth(authUrl) {
        console.log(`Starting popup authentication with URL for ${this.platform} platform:`, authUrl);
        return new Promise((resolve, reject) => {
            if (this.isCancelled) {
                console.log("❌ Authentication cancelled by user");
                reject(new Error("Authentication cancelled by user"));
                return;
            }
            
            // Reset auth completed flag
            this.authCompleted = false;
            
            // Try to open popup window
            console.log("🪟 Attempting to open popup window");
            this.authWindow = window.open(authUrl, 'oauth', 'width=600,height=600');
            
            if (!this.authWindow) {
                console.log("⚠️ Popup window was blocked by browser");
                // Show fallback modal if popup is blocked
                this.showFallbackModal(authUrl, resolve, reject);
                return;
            }
            
            console.log("✅ Popup window opened successfully");
            
            // Start checking for auth code
            this.checkAuthCode(resolve, reject);
        });
    }

    // Check for auth code by polling the API endpoint
    checkAuthCode(resolve, reject) {
        console.log("🔄 Starting to check for auth code via API");
        
        let isResolved = false;
        
        const cleanup = () => {
            if (this.pollInterval) {
                clearTimeout(this.pollInterval);
                this.pollInterval = null;
            }
            if (this.authWindow && !this.authWindow.closed) {
                this.authWindow.close();
            }
            this.authWindow = null;
        };

        const checkForCode = async () => {
            if (isResolved || this.authCompleted) {
                cleanup();
                return;
            }
            
            if (this.isCancelled) {
                cleanup();
                if (typeof window.hideLoader === 'function') {
                    window.hideLoader();
                }
                reject(new Error("Authentication cancelled by user"));
                return;
            }

            try {
                const response = await fetch(`https://multiplewords.com/oauth/office/code-by-state/?state=${this.state}`, {
                    cache: 'no-store'  // Prevent caching
                });
                
                if (response.ok) {
                    const data = await response.json();
                    
                    if (data.code && data.state && !isResolved) {
                        if (data.state !== this.state) {
                            cleanup();
                            reject(new Error("State mismatch - possible CSRF attack"));
                            return;
                        }
                        
                        // Hide loader immediately
                        if (typeof window.hideLoader === 'function') {
                            window.hideLoader();
                        }
                        
                        this.authCompleted = true;
                        isResolved = true;
                        
                        if (typeof window.showSuccess === 'function') {
                            window.showSuccess("Authentication successful!");
                        }
                        
                        cleanup();
                        
                        // Directly exchange code for token without checking token validity
                        try {
                            const tokenData = await this.exchangeCodeForToken(data.code);
                            resolve({ code: data.code, state: data.state, token: tokenData });
                        } catch (error) {
                            reject(error);
                        }
                        return;
                    }
                }
            } catch (error) {
                console.error("Error checking auth code:", error);
            }

            // Check if window was closed
            if (this.authWindow && this.authWindow.closed && !isResolved) {
                cleanup();
                if (!this.authCompleted) {
                    if (typeof window.hideLoader === 'function') {
                        window.hideLoader();
                    }
                    reject(new Error("Authentication window was closed"));
                }
                return;
            }

            // Continue checking immediately
            if (!isResolved && !this.isCancelled) {
                this.pollInterval = setTimeout(checkForCode, 0);
            }
        };

        // Start checking immediately
        checkForCode();
    }

    // Show fallback modal when popup is blocked
    showFallbackModal(authUrl, resolve, reject) {
        console.log("🔄 Showing fallback modal for blocked popup");
        
        const modal = document.createElement('div');
        modal.className = 'auth-modal';
        modal.innerHTML = `
            <div class="auth-modal-content">
                <h3>Authentication Required</h3>
                <p>Your browser blocked the authentication popup. Please follow these steps:</p>
                <ol>
                    <li>Click the link below to open the authentication page in a new tab</li>
                    <li>Complete the authentication process</li>
                    <li>Return to this window and click "I've Completed Authentication"</li>
                </ol>
                <div class="auth-modal-actions">
                    <a href="${authUrl}" target="_blank" class="ms-Button ms-Button--primary">
                        <span class="ms-Button-label">Open Authentication Page</span>
                    </a>
                    <button id="auth-complete-btn" class="ms-Button ms-Button--primary">
                        <span class="ms-Button-label">I've Completed Authentication</span>
                    </button>
                    <button id="auth-cancel-btn" class="ms-Button ms-Button--danger">
                        <span class="ms-Button-label">Cancel</span>
                    </button>
                </div>
            </div>
        `;
        
        // Add styles for the modal
        if (!document.querySelector('style.auth-modal-style')) {
            const style = document.createElement('style');
            style.className = 'auth-modal-style';
            style.textContent = `
                .auth-modal {
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: rgba(0, 0, 0, 0.5);
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    z-index: 9999;
                }
                .auth-modal-content {
                    background: white;
                    padding: 20px;
                    border-radius: 4px;
                    max-width: 500px;
                    width: 90%;
                    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
                }
                .auth-modal h3 {
                    margin-top: 0;
                    color: #0078D4;
                }
                .auth-modal ol {
                    margin-bottom: 20px;
                }
                .auth-modal-actions {
                    display: flex;
                    flex-direction: column;
                    gap: 10px;
                }
                .auth-modal-actions a, .auth-modal-actions button {
                    width: 100%;
                    text-align: center;
                    text-decoration: none;
                }
            `;
            document.head.appendChild(style);
        }
        
        document.body.appendChild(modal);
        
        // Set up event listeners for the modal buttons
        document.getElementById('auth-complete-btn').onclick = () => {
            console.log("👆 User clicked 'I've Completed Authentication' button");
            document.body.removeChild(modal);
            this.checkAuthCode(resolve, reject);
        };
        
        document.getElementById('auth-cancel-btn').onclick = () => {
            console.log("❌ User cancelled authentication via modal");
            document.body.removeChild(modal);
            reject(new Error("Authentication cancelled by user"));
        };
    }

    // Exchange code for token
    async exchangeCodeForToken(code) {
        console.log("🔄 Starting token exchange process");
        
        // Check if code has already been used
        if (this.codeUsed) {
            console.error("❌ Authorization code has already been used");
            throw new Error("Authorization code has already been used - this is expected if you're already authenticated");
        }
        
        const data = new URLSearchParams();
        data.append('grant_type', 'authorization_code');
        data.append('code', code);
        data.append('client_id', this.clientId);
        data.append('redirect_uri', this.redirectUri);
        data.append('include_user_info', 'true'); // Request user info in response
        
        try {
            console.log("📡 Sending token request to:", this.tokenEndpoint);
            console.log("📤 Request data:", {
                grant_type: 'authorization_code',
                code: code.substring(0, 5) + '...', // Log only part of the code for security
                client_id: this.clientId,
                redirect_uri: this.redirectUri
            });
            
            const response = await fetch(this.tokenEndpoint, {
                method: 'POST',
                body: data,
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Accept': 'application/json'
                }
            });
            
            console.log("📥 Token response status:", response.status);
            console.log("📥 Token response headers:", Object.fromEntries(response.headers.entries()));
            
            if (!response.ok) {
                const errorText = await response.text();
                console.error("❌ Token exchange failed with status:", response.status);
                console.error("❌ Error response:", errorText);
                console.error("❌ Request URL:", this.tokenEndpoint);
                console.error("❌ Request headers:", {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Accept': 'application/json'
                });
                console.error("❌ Request body (excluding code):", {
                    grant_type: 'authorization_code',
                    client_id: this.clientId,
                    redirect_uri: this.redirectUri,
                    include_user_info: 'true'
                });
                
                let errorMessage = `Token exchange failed: ${response.status} ${response.statusText}`;
                try {
                    const errorJson = JSON.parse(errorText);
                    if (errorJson.error) {
                        errorMessage += ` - ${errorJson.error}`;
                        if (errorJson.error_description) {
                            errorMessage += `: ${errorJson.error_description}`;
                        }
                    }
                } catch (e) {
                    errorMessage += ` - ${errorText}`;
                }
                
                throw new Error(errorMessage);
            }
            
            const tokenData = await response.json();
            console.log("✅ Token exchange successful");
            
            // Mark the code as used
            this.codeUsed = true;
            
            // Get user ID from check-canva-token endpoint
            try {
                console.log("🔍 Getting user ID from check-canva-token endpoint");
                const userResponse = await fetch('https://multiplewords.com/oauth/check-canva-token', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                        'Accept': 'application/json'
                    },
                    body: new URLSearchParams({
                        'token': tokenData.access_token
                    })
                });
                
                console.log("📥 Check-canva-token response status:", userResponse.status);
                
                if (userResponse.ok) {
                    const userData = await userResponse.json();
                    if (userData.user_id) {
                        tokenData.user_id = userData.user_id;
                        console.log("✅ User ID found from check-canva-token:", userData.user_id);
                    } else {
                        console.warn("⚠️ No user ID found in check-canva-token response");
                    }
                } else {
                    const errorText = await userResponse.text();
                    console.warn("⚠️ Failed to get user ID from check-canva-token:", userResponse.status);
                    console.warn("⚠️ Error response:", errorText);
                }
            } catch (e) {
                console.warn("⚠️ Error fetching from check-canva-token:", e);
            }
            
            // Store tokens and user info
            await this.storeTokens(tokenData);
            
            // Show success message
            if (typeof window.showSuccess === 'function') {
                window.showSuccess("Authentication successful!");
            }
            
            return tokenData;
        } catch (error) {
            console.error("❌ Token exchange failed:", error);
            throw error;
        }
    }

    // Store tokens securely with immediate effect
    async storeTokens(tokens) {
        try {
            sessionStorage.setItem('access_token', tokens.access_token);
            sessionStorage.setItem('refresh_token', tokens.refresh_token);
            sessionStorage.setItem('token_expiry', Date.now() + (tokens.expires_in * 1000));
            
            if (tokens.user_id) {
                sessionStorage.setItem('user_id', tokens.user_id);
                console.log("✅ User ID stored:", tokens.user_id);
            } else {
                console.warn("⚠️ No user ID available in token data");
            }
            
            console.log("✅ Tokens stored successfully");
            
            // Update UI immediately after storing tokens
            if (typeof window.checkTokens === 'function') {
                window.checkTokens();
            }
        } catch (error) {
            console.error("❌ Error storing tokens:", error);
            throw error;
        }
    }

    // Get user ID
    getUserId() {
        const userId = sessionStorage.getItem('user_id');
        if (!userId) {
            console.warn("⚠️ No user ID found in session storage");
        }
        return userId;
    }

    // Check if user is authenticated (based on user ID)
    isTokenExpired() {
        const userId = this.getUserId();
        const isExpired = !userId;
        console.log(`🔍 Authentication check: ${isExpired ? 'Not authenticated (no user ID)' : 'Authenticated'}`);
        return isExpired;
    }
    
    // // Get access token (refreshing if needed)
    // async getAccessToken() {
    //     console.log("🔑 Getting access token");
    //     if (this.isTokenExpired()) {
    //         console.log("🔄 Token expired, refreshing...");
    //         await this.refreshToken();
    //     } else {
    //         console.log("✅ Using existing valid token");
    //     }
    //     const token = sessionStorage.getItem('access_token');
    //     console.log(`🔑 Access token: ${token ? `${token.substring(0, 10)}...` : 'undefined'}`);
    //     return token;
    // }
    
    // Refresh token
    async refreshToken() {
        console.log("🔄 Refreshing token");
        const refreshToken = sessionStorage.getItem('refresh_token');
        if (!refreshToken) {
            console.error("❌ No refresh token available");
            throw new Error('No refresh token available');
        }
        
        console.log(`🔑 Using refresh token: ${refreshToken.substring(0, 10)}...`);
        
        const data = new URLSearchParams();
        data.append('grant_type', 'refresh_token');
        data.append('refresh_token', refreshToken);
        data.append('client_id', this.clientId);
        
        console.log("📤 Refresh token request data:", {
            grant_type: 'refresh_token',
            refresh_token: `${refreshToken.substring(0, 10)}...`,
            client_id: this.clientId
        });
        
        try {
            console.log(`📡 Sending refresh token request to: ${this.tokenEndpoint}`);
            const response = await fetch(this.tokenEndpoint, {
                method: 'POST',
                body: data,
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            
            console.log(`📥 Refresh token response status: ${response.status}`);
            
            if (!response.ok) {
                console.error("❌ Token refresh failed");
                throw new Error('Token refresh failed');
            }
            
            const tokens = await response.json();
            console.log("✅ Token refresh successful");
            console.log("🔑 New token data:", {
                access_token: tokens.access_token ? `${tokens.access_token.substring(0, 10)}...` : 'undefined',
                refresh_token: tokens.refresh_token ? `${tokens.refresh_token.substring(0, 10)}...` : 'undefined',
                expires_in: tokens.expires_in,
                token_type: tokens.token_type
            });
            
            this.storeTokens(tokens);
            return tokens;
        } catch (error) {
            console.error("❌ Error refreshing token:", error);
            throw error;
        }
    }
    
    // Start the complete authentication process
    async authenticate() {
        try {
            this.isCancelled = false;
            this.codeUsed = false; // Reset the code used flag
            
            const authResult = await this.startAuthFlow();
            
            if (this.isCancelled) {
                return {
                    success: false,
                    error: "Authentication cancelled by user"
                };
            }
            
            try {
                const tokens = await this.exchangeCodeForToken(authResult.code);
                
                if (typeof window.hideLoader === 'function') {
                    window.hideLoader();
                }
                
                return {
                    success: true,
                    token: tokens.access_token,
                    userId: this.getUserId()
                };
            } catch (error) {
                // Check if the error is because the code was already used
                if (error.message.includes("Authorization code has already been used")) {
                    // Check if the user is already authenticated
                    if (!this.isTokenExpired()) {
                        // User is already authenticated, return success
                        if (typeof window.hideLoader === 'function') {
                            window.hideLoader();
                        }
                        
                        return {
                            success: true,
                            token: sessionStorage.getItem('access_token'),
                            userId: this.getUserId()
                        };
                    }
                }
                
                // If we get here, it's a real error
                throw error;
            }
        } catch (error) {
            console.error("❌ Authentication process failed:", error);
            if (typeof window.hideLoader === 'function') {
                window.hideLoader();
            }
            return {
                success: false,
                error: error.message
            };
        }
    }
    
    // Logout function
    logout() {
        console.log("🚪 Logging out user");
        sessionStorage.removeItem('access_token');
        sessionStorage.removeItem('refresh_token');
        sessionStorage.removeItem('token_expiry');
        sessionStorage.removeItem('oauth_state');
        sessionStorage.removeItem('user_id');
        console.log("✅ Logout completed, all tokens and user ID removed");
    }

    // Cancel authentication process
    cancelAuth() {
        console.log("❌ Cancelling authentication process");
        this.isCancelled = true;
        
        if (this.pollInterval) {
            console.log("🛑 Clearing polling interval");
            clearInterval(this.pollInterval);
            this.pollInterval = null;
        }
        
        if (this.authWindow && !this.authWindow.closed) {
            console.log("🪟 Closing auth window");
            this.authWindow.close();
            this.authWindow = null;
        }
        
        this.authCompleted = false;
        this.state = null;
        console.log("✅ Authentication cancellation completed");
    }

    // Open dialog for purchasing credits
    async openCreditsDialog(userId) {
        console.log("🛒 Opening credits purchase dialog for user:", userId);
        
        // Construct the URL for the credits purchase page
        const creditsUrl = `https://saifs.ai/canva_pricing/${userId}/16`;
        console.log("🔗 Constructed credits URL:", creditsUrl);
        
        return new Promise((resolve, reject) => {
            // Create a separate dialog instance for credits to avoid conflicts with auth dialog
            Office.context.ui.displayDialogAsync(creditsUrl, 
                { height: 80, width: 80 }, 
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.error("❌ Failed to open credits dialog:", result.error);
                        reject(new Error(`Failed to open credits dialog: ${result.error.message}`));
                        return;
                    }
                    
                    // Store the credits dialog separately from the auth dialog
                    const creditsDialog = result.value;
                    console.log("✅ Credits dialog opened successfully");
                    
                    // Set up event handler for dialog events
                    creditsDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                        console.log("📬 Credits dialog event received:", arg);
                        // Handle any events from the dialog if needed
                    });
                    
                    // Set up event handler for dialog messages
                    creditsDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                        console.log("📬 Credits dialog message received:", arg);
                        // Handle any messages from the dialog if needed
                    });
                    
                    // Set up event handler for dialog close
                    creditsDialog.addEventHandler(Office.EventType.DialogClosed, (arg) => {
                        console.log("🚪 Credits dialog closed:", arg);
                        resolve({ success: true });
                    });
                    
                    // Display the dialog
                    creditsDialog.displayAsync();
                }
            );
        });
    }
}

// Export the class
window.OfficeAuthManager = OfficeAuthManager; 