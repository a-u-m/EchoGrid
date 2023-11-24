import './content.css';
const createSheetExtensionIcon = () => {
  const iconContainer = document.createElement('div');
  iconContainer.id = 'sheet-extension-icon';
  iconContainer.classList.add('icon-container'); // Added a CSS class for styling
  iconContainer.innerHTML = '<i class="fas fa-microphone"></i>'; // Using Font Awesome for the microphone icon
  return iconContainer;
};
const showRecognition = () => {
  const textContainer = document.createElement('div');
  textContainer.id = 'NLP';
  textContainer.style.display = 'flex';
  textContainer.style.alignItems = 'center';
  textContainer.style.justifyContent = 'center';
  textContainer.style.position = 'absolute';
  textContainer.style.maxWidth = '50vw';
  textContainer.style.zIndex = '1000';
  textContainer.style.left = '20%';
  textContainer.style.top = '5%';
  textContainer.style.fontSize = '3rem';
  textContainer.style.color = 'white';
  textContainer.style.background='#848884'
  return textContainer;
};

const initializeVoiceRecognition = () => {
  let finalTranscript = '';
  const recognition = new webkitSpeechRecognition();
  recognition.continuous = true;
  recognition.interimResults = true;

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
    const interimTranscript = event.results[0][0].transcript;
    finalTranscript = interimTranscript;
    document.getElementById('NLP').innerHTML = finalTranscript;
  };

  recognition.onend = () => {
    console.log('Recognition ended');
    chrome.runtime.sendMessage({ action: 'recognizedText', text: finalTranscript });
    finalTranscript = '';
  };

  return recognition;
};

window.onload = () => {
  const sheetsToolbar = document.querySelector('.docs-titlebar-buttons');
  const iconContainer = createSheetExtensionIcon();
  const sheet = document.querySelector('.docs-gm');
  const textContainer = showRecognition();
  const recognition = initializeVoiceRecognition();
  let isRecording = false;
//  const fontAwesomeLink = document.createElement('link');
// fontAwesomeLink.rel = 'stylesheet';
// fontAwesomeLink.href = 'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css';
// fontAwesomeLink.integrity = 'sha384-XI8J7luUdDl5LxxHqDEPvLYBqYZfMy8v1CswPn8CgpR9giRHzp5Gn14Gr1rAx7Bd';
// fontAwesomeLink.crossOrigin = 'anonymous';

// document.head.appendChild(fontAwesomeLink);

  if (sheetsToolbar) {
    sheetsToolbar.insertBefore(iconContainer, sheetsToolbar.firstChild);
    iconContainer.addEventListener('click', () => {
      if (isRecording) {
        recognition.stop();
        iconContainer.style.backgroundColor = '#3498db';
        document.getElementById('NLP').remove();
      } else {
        sheet.insertBefore(textContainer, sheet.children[0]);
        document.getElementById('NLP').innerHTML = 'Listening...';
        recognition.start();
        iconContainer.style.backgroundColor = 'green';
      }
      isRecording = !isRecording;
    });
  } else {
    console.error('Toolbar not found. Icon not injected.');
  }
};
