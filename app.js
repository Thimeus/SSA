// JavaScript (app.js)

/**
 * Globale Variablen
 */
let excelData = null;
let participants = [];
let currentParticipant = null;
let charts = {};

/**
 * DOM-Elemente (Caching für bessere Performance)
 */
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const selectFileBtn = document.getElementById('selectFileBtn');
const participantList = document.getElementById('participantList');
const participantItems = document.getElementById('participantItems');
const dropArea = document.getElementById('dropArea');
const dropText = document.getElementById('dropText');
const analysisSection = document.getElementById('analysisSection');
const errorDisplay = document.getElementById('errorDisplay');
const loadingOverlay = document.getElementById('loadingOverlay');
const mainNavigation = document.getElementById('mainNavigation');

/**
 * Helferfunktionen
 */

/**
 * Zeigt eine Fehlermeldung an.
 */
function showError(message) {
    errorDisplay.textContent = message;
    errorDisplay.style.display = 'block';
    setTimeout(() => {
        errorDisplay.style.display = 'none';
    }, 5000); // Ausblenden nach 5 Sekunden
}

/**
 * Zeigt das Ladeoverlay an.
 */
function showLoading() {
    loadingOverlay.style.display = 'flex';
}

/**
 * Blendet das Ladeoverlay aus.
 */
function hideLoading() {
    loadingOverlay.style.display = 'none';
}

/**
 * Zerstört ein Chart.js Diagramm, falls es existiert.
 */
function destroyChart(chartId) {
    if (charts[chartId]) {
        charts[chartId].destroy();
        delete charts[chartId];
    }
}

/**
 * Hilfsfunktion zum Formatieren von Datumsangaben
 */
function formatDate(date) {
    if (!date) return '';

    // Prüfen, ob es sich um ein Datum-Objekt handelt
    if (date instanceof Date) {
        return date.toLocaleDateString('de-DE');
    }

    // Falls es ein String ist, unverändert zurückgeben
    if (typeof date === 'string') {
        return date;
    }

    // Falls es ein Excel-Serialdatum ist, umwandeln
    if (typeof date === 'number') {
        // Excel-Datum in JS-Datum umwandeln (Excel-Epoche beginnt am 1.1.1900)
        const excelEpoch = new Date(1899, 11, 30);
        const millisecondsPerDay = 24 * 60 * 60 * 1000;
        const jsDate = new Date(excelEpoch.getTime() + date * millisecondsPerDay);
        return jsDate.toLocaleDateString('de-DE');
    }

    return '';
}

/**
 * Excel-Datenverarbeitung
 */

/**
 * Verarbeitet den Excel-Datenimport.
 */
async function handleFileUpload(file) {
    showLoading();
    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, {
            type: 'array',
            cellDates: true,
            cellStyles: true
        });

        excelData = processExcelData(workbook);
        participants = excelData.teilnehmer;

        if (participants.length === 0) {
            showError("Keine Teilnehmerdaten gefunden.");
            return;
        }

        uploadArea.classList.add('hidden');
        participantList.classList.remove('hidden');
        dropArea.classList.remove('hidden');
        mainNavigation.classList.remove('hidden');
        renderParticipantList(); // Render die Liste
        setupParticipantDragDrop(); //Drag&Drop Funktion
    } catch (error) {
        console.error("Fehler beim Verarbeiten der Datei:", error);
        showError(`Fehler beim Verarbeiten der Datei: ${error.message}`);
    } finally {
        hideLoading();
    }
}

/**
 * Verarbeitet die Daten aus dem Excel-Workbook.
 */
function processExcelData(workbook) {
    const result = {
        teilnehmer: [],
        schulisch: {},
        selbstFremd: {},
        berufswahl: {},
        erprobung: {},
        selbstBerufe: {},
        ressourcen: {}
    };

    const teilnehmerSheet = findSheet(workbook.SheetNames, ['Teilnehmerliste', 'Teilnehmer']);
    const schulischSheet = findSheet(workbook.SheetNames, ['Schulische_Basiskompetenzen', 'Schulisch']);
    const selbstFremdSheet = findSheet(workbook.SheetNames, ['Selbst_Fremdeinschaetzung', 'Selbst']);
    const berufsorientierungSheet = findSheet(workbook.SheetNames, ['Berufsorientierung', 'Berufe']);
    const anmerkungenSheet = findSheet(workbook.SheetNames, ['Anmerkungen_Ressourcen', 'Anmerkungen']);

    if (!teilnehmerSheet) {
        throw new Error("Keine Teilnehmerliste gefunden.");
    }

    result.teilnehmer = XLSX.utils.sheet_to_json(workbook.Sheets[teilnehmerSheet]);
    if (schulischSheet) result.schulisch = processSheet(workbook.Sheets[schulischSheet]);
    if (selbstFremdSheet) result.selbstFremd = processSheet(workbook.Sheets[selbstFremdSheet]);
    if (berufsorientierungSheet) result.berufswahl = processSheet(workbook.Sheets[berufsorientierungSheet]);
    if (anmerkungenSheet) result.ressourcen = processSheet(workbook.Sheets[anmerkungenSheet]);

    return result;
}

/**
 * Hilfsfunktion zum Finden eines Tabellenblatts
 */
function findSheet(sheetNames, possibleNames) {
    for (const name of possibleNames) {
        const foundSheet = sheetNames.find(sheet => sheet.toLowerCase().includes(name.toLowerCase()));
        if (foundSheet) return foundSheet;
    }
    return undefined;
}

/**
 * Verarbeitet die einzelnen Blätter
 */
function processSheet(sheet) {
    return XLSX.utils.sheet_to_json(sheet);
}

/**
 * UI-Funktionen
 */

/**
 * Rendert die Teilnehmerliste.
 */
function renderParticipantList() {
    participantItems.innerHTML = '';
    participants.forEach(participant => {
        const item = document.createElement('div');
        item.className = 'participant-item';
        item.draggable = true;
        item.dataset.id = participant.id;
        item.textContent = participant.name || 'Unbenannt';
        participantItems.appendChild(item);
    });
}

/**
 * Richtet Drag & Drop für Teilnehmer ein.
 */
function setupParticipantDragDrop() {
    const items = document.querySelectorAll('.participant-item');

    items.forEach(item => {
        item.addEventListener('dragstart', (e) => {
            e.dataTransfer.setData('text/plain', item.dataset.id);
        });
    });

    dropArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropArea.classList.add('drag-over');
    });

    dropArea.addEventListener('dragleave', () => {
        dropArea.classList.remove('drag-over');
    });

    dropArea.addEventListener('drop', (e) => {
        e.preventDefault();
        dropArea.classList.remove('drag-over');

        const tnId = e.dataTransfer.getData('text/plain');
        const participant = participants.find(p => p.id == tnId);

        if (participant) {
            currentParticipant = participant;
            dropText.textContent = `Teilnehmer: ${participant.name || 'Unbenannt'}`;
            analysisSection.classList.remove('hidden');
            loadParticipantData(participant);
        }
    });
}

/**
 * Lädt die Daten für den ausgewählten Teilnehmer.
 */
function loadParticipantData(participant) {
    // Hier kannst du die Daten des Teilnehmers laden und in die entsprechenden Bereiche der Seite einfügen
    console.log("Daten geladen für Teilnehmer:", participant);
}

/**
 * Definiert Event Listeners
 */

/**
 * Fügt Event Listener für Drag & Drop hinzu.
 */
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('drag-over');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');

    if (e.dataTransfer.files.length) {
        handleFileUpload(e.dataTransfer.files[0]);
    }
});

/**
 * Fügt Event Listener für Dateiauswahl hinzu.
 */
selectFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
        handleFileUpload(e.target.files[0]);
    }
});

/**
 * Event-Listener hinzufügen, wenn das DOM vollständig geladen ist
 */
document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM ist geladen');
});
