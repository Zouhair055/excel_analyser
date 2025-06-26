// Configuration de l'API
const API_BASE_URL = 'http://localhost:5001/api/excel';

// Éléments DOM
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');
const uploadSection = document.getElementById('uploadSection');
const progressSection = document.getElementById('progressSection');
const resultsSection = document.getElementById('resultsSection');
const errorSection = document.getElementById('errorSection');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const downloadBtn = document.getElementById('downloadBtn');
const newFileBtn = document.getElementById('newFileBtn');
const retryBtn = document.getElementById('retryBtn');
const errorMessage = document.getElementById('errorMessage');
const topLogo = document.getElementById('topLogo'); // Nouveau élément

// Variables globales
let currentProcessedFile = null;

// Initialisation
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
});

function initializeEventListeners() {
    // Upload area events
    uploadArea.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);
    
    // File input
    fileInput.addEventListener('change', handleFileSelect);
    
    // Buttons
    browseBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        fileInput.click();
    });
    
    downloadBtn.addEventListener('click', downloadFile);
    newFileBtn.addEventListener('click', resetToUpload);
    retryBtn.addEventListener('click', resetToUpload);
}

// Gestion du drag & drop
function handleDragOver(e) {
    e.preventDefault();
    uploadArea.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

// Gestion de la sélection de fichier
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

// Validation et traitement du fichier
function handleFile(file) {
    // Validation du type de fichier
    const allowedTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    
    if (!allowedTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Type de fichier non supporté. Veuillez utiliser un fichier .xlsx ou .xls');
        return;
    }
    
    // Validation de la taille (max 10MB)
    if (file.size > 10 * 1024 * 1024) {
        showError('Le fichier est trop volumineux. Taille maximale: 10MB');
        return;
    }
    
    uploadFile(file);
}

// Upload et traitement du fichier
async function uploadFile(file) {
    showProgress();
    
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        // Simulation de progression
        updateProgress(20, 'Upload du fichier...');
        
        const response = await fetch(`${API_BASE_URL}/upload`, {
            method: 'POST',
            body: formData
        });
        
        updateProgress(60, 'Analyse des colonnes...');
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Erreur lors de l\'upload');
        }
        
        updateProgress(80, 'Application des règles...');
        
        const result = await response.json();
        
        updateProgress(100, 'Traitement terminé !');
        
        setTimeout(() => {
            showResults(result);
        }, 500);
        
    } catch (error) {
        console.error('Erreur:', error);
        showError(error.message || 'Une erreur s\'est produite lors du traitement du fichier');
    }
}

// Affichage des sections
function showProgress() {
    hideAllSections();
    progressSection.style.display = 'block';
    progressSection.classList.add('fade-in');
    updateProgress(0, 'Préparation...');
    // Afficher le logo petit
    showTopLogo();
}

function showResults(data) {
    hideAllSections();
    resultsSection.style.display = 'block';
    resultsSection.classList.add('fade-in');
    
    // Stocker le fichier traité
    currentProcessedFile = data.processed_file;
    
    // Remplir les informations
    document.getElementById('fileName').textContent = data.original_file;
    document.getElementById('fileRows').textContent = data.columns_info.shape[0];
    document.getElementById('fileColumns').textContent = data.columns_info.shape[1];
    
    // Afficher les colonnes
    displayColumns(data.columns_info.columns, data.columns_info.empty_columns);
    
    // Afficher les règles appliquées
    displayRules(data.changes_applied.rules_applied);
    
    // Garder le logo petit visible
    showTopLogo();
}

function showError(message) {
    hideAllSections();
    errorSection.style.display = 'block';
    errorSection.classList.add('fade-in');
    errorMessage.textContent = message;
    // Afficher le logo petit
    showTopLogo();
}

function hideAllSections() {
    [uploadSection, progressSection, resultsSection, errorSection].forEach(section => {
        section.style.display = 'none';
        section.classList.remove('fade-in');
    });
}

function resetToUpload() {
    hideAllSections();
    uploadSection.style.display = 'block';
    uploadSection.classList.add('fade-in');
    
    // Reset du formulaire
    fileInput.value = '';
    currentProcessedFile = null;
    
    // Cacher le logo petit
    hideTopLogo();
}

// Gestion du logo petit
function showTopLogo() {
    topLogo.style.display = 'block';
    topLogo.classList.add('show');
}

function hideTopLogo() {
    topLogo.style.display = 'none';
    topLogo.classList.remove('show');
}

// Mise à jour de la progression
function updateProgress(percentage, text) {
    progressFill.style.width = `${percentage}%`;
    progressText.textContent = text;
}

// Affichage des colonnes
function displayColumns(columns, emptyColumns) {
    const columnsList = document.getElementById('columnsList');
    columnsList.innerHTML = '';
    
    columns.forEach(column => {
        const tag = document.createElement('div');
        tag.className = 'column-tag';
        tag.textContent = column;
        
        if (emptyColumns.includes(column)) {
            tag.classList.add('empty');
            tag.title = 'Colonne vide - sera remplie automatiquement';
        }
        
        columnsList.appendChild(tag);
    });
}

// Affichage des règles
function displayRules(rules) {
    const rulesList = document.getElementById('rulesList');
    rulesList.innerHTML = '';
    
    rules.forEach(rule => {
        const item = document.createElement('div');
        item.className = 'rule-item';
        item.innerHTML = `<i class="fas fa-check"></i> ${rule}`;
        rulesList.appendChild(item);
    });
}

// Téléchargement du fichier
async function downloadFile() {
    if (!currentProcessedFile) {
        showError('Aucun fichier traité disponible');
        return;
    }
    
    try {
        const response = await fetch(`${API_BASE_URL}/download/${currentProcessedFile}`);
        
        if (!response.ok) {
            throw new Error('Erreur lors du téléchargement');
        }
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = currentProcessedFile;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
    } catch (error) {
        console.error('Erreur de téléchargement:', error);
        showError('Erreur lors du téléchargement du fichier');
    }
}

// Gestion des erreurs réseau
window.addEventListener('online', () => {
    console.log('Connexion rétablie');
});

window.addEventListener('offline', () => {
    showError('Connexion internet perdue. Veuillez vérifier votre connexion.');
});