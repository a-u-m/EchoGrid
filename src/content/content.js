const createSheetExtensionIcon = () => {
  const iconContainer = document.createElement('div');
  iconContainer.id = 'sheet-extension-icon';
  iconContainer.style.width = '24px';
  iconContainer.style.height = '24px';
  iconContainer.style.backgroundColor = 'red';
  iconContainer.style.cursor = 'pointer';

  return iconContainer;
};

const initializeVoiceRecognition = () => {
  let finalTranscript = '';
  const recognition = new webkitSpeechRecognition();
  recognition.continuous = true;
  recognition.interimResults = false;

  recognition.onstart = () => {
    console.log('Recognition started');
  };

  recognition.onerror = (event) => {
    console.error('Recognition error:', event.error);
    if (event.error === 'not-allowed') {
      alert('Microphone access denied. Please allow microphone access in your browser settings.');
    }
  };

  recognition.onresult = (event) => {
    finalTranscript = event.results[0][0].transcript;
  };

  recognition.onend = () => {
    console.log(finalTranscript);
    chrome.runtime.sendMessage({ action: 'recognizedText', text: finalTranscript });
    finalTranscript = '';
  };

  return recognition;
};

window.onload = () => {
  const sheetsToolbar = document.querySelector('.docs-titlebar-buttons');
  const iconContainer = createSheetExtensionIcon();
  // const backgroundPortConnect = chrome.runtime.connect({ name: 'bg-port' });
  const recognition = initializeVoiceRecognition();
  let isRecording = false;

  if (sheetsToolbar) {
    sheetsToolbar.insertBefore(iconContainer, sheetsToolbar.firstChild);
    iconContainer.addEventListener('click', () => {
      if (isRecording) {
        recognition.stop();
        iconContainer.style.backgroundColor = 'red';
      } else {
        recognition.start();
        iconContainer.style.backgroundColor = 'green';
      }
      isRecording = !isRecording;
    });
  } else {
    console.error('Toolbar not found. Icon not injected.');
  }
};
