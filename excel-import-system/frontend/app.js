const API_URL = window.location.origin;

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');
const filePreview = document.getElementById('filePreview');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const fileType = document.getElementById('fileType');
const removeBtn = document.getElementById('removeBtn');
const uploadBtn = document.getElementById('uploadBtn');
const progressSection = document.getElementById('progressSection');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const alert = document.getElementById('alert');
const alertIcon = document.getElementById('alertIcon');
const alertTitle = document.getElementById('alertTitle');
const alertMessage = document.getElementById('alertMessage');
const toast = document.getElementById('toast');
const toastIcon = document.getElementById('toastIcon');
const toastTitle = document.getElementById('toastTitle');
const toastMessage = document.getElementById('toastMessage');
const toastClose = document.getElementById('toastClose');

let selectedFile = null;
let toastTimeout = null;

document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  checkServerHealth();
  requestNotificationPermission();
});

function setupEventListeners() {

  browseBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });
  

  uploadArea.addEventListener('click', (e) => {

    if (e.target !== browseBtn && !browseBtn.contains(e.target)) {
      fileInput.click();
    }
  });
  
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFile(e.target.files[0]);
    }
  });
  
  uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
  });
  
  uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
  });
  
  uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    
    if (e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0]);
    }
  });
  
  removeBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    resetForm();
  });
  
  uploadBtn.addEventListener('click', uploadFile);
  
  toastClose.addEventListener('click', hideToast);
}

function handleFile(file) {

  const allowedExts = ['xlsx', 'xls', 'csv'];
  const ext = file.name.split('.').pop().toLowerCase();
  
  if (!allowedExts.includes(ext)) {
    showAlert('error', 'Format Tidak Didukung', 'Gunakan file Excel (.xlsx, .xls) atau CSV (.csv)');
    return;
  }
  
  if (file.size > 10 * 1024 * 1024) {
    showAlert('error', 'File Terlalu Besar', 'Ukuran maksimal 10MB');
    return;
  }
  
  selectedFile = file;

  fileName.textContent = file.name;
  fileSize.textContent = formatFileSize(file.size);
  fileType.textContent = ext.toUpperCase();
  
  uploadArea.style.display = 'none';
  filePreview.classList.add('show');
  uploadBtn.disabled = false;
  
  hideAlert();
}

function resetForm() {
  selectedFile = null;
  fileInput.value = '';
  uploadArea.style.display = 'block';
  filePreview.classList.remove('show');
  uploadBtn.disabled = true;
  hideAlert();
}

async function uploadFile() {
  if (!selectedFile) return;
  
  uploadBtn.disabled = true;
  showProgress('Mengunggah file...', 10);
  
  try {
    const formData = new FormData();
    formData.append('file', selectedFile);
    
    updateProgress('Memproses data...', 30);
    
    const response = await fetch(`${API_URL}/api/upload`, {
      method: 'POST',
      body: formData
    });
    
    const result = await response.json();
    
    if (result.success && result.data.rowsImported > 1000) {
      updateProgress('Menulis ke Google Sheets...', 50);
      await new Promise(r => setTimeout(r, 500));
      
      updateProgress('Menerapkan formula...', 70);
      await new Promise(r => setTimeout(r, 500));
      
      updateProgress('Memformat data...', 90);
      await new Promise(r => setTimeout(r, 500));
    }
    
    hideProgress();
    
    if (result.success) {

      let detailedMessage = `Berhasil import ${result.data.rowsImported.toLocaleString()} baris data`;
      
      if (result.data.duplicatesSkipped > 0) {
        detailedMessage += `\n${result.data.duplicatesSkipped.toLocaleString()} data duplikat di-skip`;
      }
      
      if (result.data.totalTime) {
        detailedMessage += `\n\nWaktu proses: ${result.data.totalTime} detik`;
      }
      
      showAlert('success', 'Import Berhasil!', detailedMessage);
      showToast('success', 'Import Berhasil!', 
        `${result.data.rowsImported.toLocaleString()} rows imported`);
      showNotification('Import Berhasil!', detailedMessage, 'success');
      playSuccessSound();
      resetForm();
    } else {
      showAlert('error', 'Import Gagal', result.message);
      uploadBtn.disabled = false;
    }
    
  } catch (error) {
    hideProgress();
    const errorMsg = error.message || 'Gagal menghubungi server';
    showAlert('error', 'Terjadi Kesalahan', errorMsg);
    uploadBtn.disabled = false;
  }
}

function showToast(type, title, message) {

  if (toastTimeout) {
    clearTimeout(toastTimeout);
  }

  toastIcon.textContent = type === 'success' ? '✓' : '✗';
  toastTitle.textContent = title;
  toastMessage.textContent = message;

  toast.classList.remove('toast-success', 'toast-error', 'hide');
  toast.classList.add('show', 'toast-' + type);

  toastTimeout = setTimeout(() => {
    hideToast();
  }, 5000);
}

function hideToast() {
  toast.classList.remove('show');
  toast.classList.add('hide');
  
  setTimeout(() => {
    toast.classList.remove('hide');
  }, 300);
}

function showProgress(text, percentage) {
  progressSection.classList.add('show');
  progressText.textContent = text;
  progressFill.style.width = percentage + '%';
}

function updateProgress(text, percentage) {
  progressText.textContent = text;
  progressFill.style.width = percentage + '%';
}

function hideProgress() {
  progressSection.classList.remove('show');
  progressFill.style.width = '0%';
}

function showAlert(type, title, message) {
  alert.classList.remove('alert-success', 'alert-error');
  alert.classList.add('show', 'alert-' + type);
  
  alertIcon.textContent = type === 'success' ? '✓' : '✗';
  alertTitle.textContent = title;
  alertMessage.textContent = message;
}

function hideAlert() {
  alert.classList.remove('show');
}

function requestNotificationPermission() {
  if ('Notification' in window && Notification.permission === 'default') {
    Notification.requestPermission().then(permission => {
      console.log('Notification permission:', permission);
    });
  }
}

function showNotification(title, message, type = 'success') {

  if (!('Notification' in window)) {
    console.log('Browser tidak support notifikasi');
    return;
  }
  
  if (Notification.permission !== 'granted') {
    console.log('Notifikasi tidak diizinkan');
    return;
  }
  
  const options = {
    body: message,
    icon: type === 'success' 
      ? 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="green"><circle cx="12" cy="12" r="10"/><path fill="white" d="M9 12l2 2 4-4"/></svg>'
      : 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="red"><circle cx="12" cy="12" r="10"/><path fill="white" d="M15 9l-6 6m0-6l6 6"/></svg>',
    badge: '/favicon.ico',
    tag: 'excel-import',
    requireInteraction: false,
    silent: false
  };
  
  try {
    const notification = new Notification(title, options);

    setTimeout(() => {
      notification.close();
    }, 5000);
 
    notification.onclick = () => {
      window.focus();
      notification.close();
    };
    
    console.log('✓ Notifikasi ditampilkan');
  } catch (error) {
    console.error('Error showing notification:', error);
  }
}

function playSuccessSound() {
  try {

    const audioContext = new (window.AudioContext || window.webkitAudioContext)();

    const playTone = (frequency, startTime, duration) => {
      const oscillator = audioContext.createOscillator();
      const gainNode = audioContext.createGain();
      
      oscillator.connect(gainNode);
      gainNode.connect(audioContext.destination);
      
      oscillator.frequency.value = frequency;
      oscillator.type = 'sine';
      
      gainNode.gain.setValueAtTime(0.3, startTime);
      gainNode.gain.exponentialRampToValueAtTime(0.01, startTime + duration);
      
      oscillator.start(startTime);
      oscillator.stop(startTime + duration);
    };
    
    const now = audioContext.currentTime;
    playTone(523.25, now, 0.1);       
    playTone(659.25, now + 0.1, 0.2); 
    
    console.log('✓ Success sound played');
  } catch (error) {
    console.log('Audio not available:', error);
  }
}

function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

async function checkServerHealth() {
  try {
    const response = await fetch(`${API_URL}/health`);
    const data = await response.json();
    console.log('✓ Server status:', data.status);
  } catch (error) {
    console.error('✗ Server tidak dapat dijangkau');
    showAlert('error', 'Server Offline', 'Backend server tidak dapat dijangkau. Pastikan server sudah running di port 3000.');
  }
}