import React, { useState, useEffect } from 'react';
import Header from './components/Header';
import ImageGenerator from './components/ImageGenerator';
import CreditsFooter from './components/CreditsFooter';
import Loader from './components/Loader';
import Notification from './components/Notification';
import { AuthProvider } from './contexts/AuthContext';
import { useAuth } from './hooks/useAuth';

const App = ({ authManager }) => {
  const [notification, setNotification] = useState({ message: '', type: '' });
  const [showLoader, setShowLoader] = useState(false);
  const [loaderMessage, setLoaderMessage] = useState('Generating your image...');
  const [isGenerating, setIsGenerating] = useState(true);

  // Show notification function
  const showNotification = (message, type) => {
    setNotification({ message, type });
    setTimeout(() => {
      setNotification({ message: '', type: '' });
    }, 3000);
  };

  // Show error function
  const showError = (message) => {
    showNotification(message, 'error');
  };

  // Show success function
  const showSuccess = (message) => {
    showNotification(message, 'success');
  };

  // Show loader function
  const showLoaderWithMessage = (message = "Generating your image...", isGenerating = true) => {
    setLoaderMessage(message);
    setIsGenerating(isGenerating);
    setShowLoader(true);
  };

  // Hide loader function
  const hideLoader = () => {
    setShowLoader(false);
  };

  return (
    <AuthProvider authManager={authManager}>
      <div className="app-container">
        <Notification 
          message={notification.message} 
          type={notification.type} 
        />
        
        <Header />
        
        <ImageGenerator 
          showLoader={showLoaderWithMessage}
          hideLoader={hideLoader}
          showError={showError}
          showSuccess={showSuccess}
        />
        
        <CreditsFooter />
        
        {showLoader && (
          <Loader 
            message={loaderMessage} 
            isGenerating={isGenerating}
            onCancel={isGenerating ? hideLoader : null}
          />
        )}
      </div>
    </AuthProvider>
  );
};

export default App; 