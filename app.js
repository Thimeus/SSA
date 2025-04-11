// app.js

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
 * Helferfunktionen (Auslagerung für Übersichtlichkeit)
 */

/**
 * Zeigt eine Fehlermeldung an.
 * @param {string} message - Die Fehlermeldung, die angezeigt werden soll.
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
 * @param {string} chartId - Die ID des Canvas-Elements, das das Diagramm enthält.
 */
function destroyChart(chartId) {
    if (charts[chartId]) {
        charts[chartId].destroy();
        delete charts[chartId];
    }
}

/**
 * Erstellt ein responsives Liniendiagramm mit Chart.js.
 * @param {string} chartId - Die ID des Canvas-Elements.
 * @param {object} data - Die Daten für das Diagramm.
 * @param {string} title - Der Titel des Diagramms.
 */
function createResponsiveLineChart(chartId, data, title) {
    const ctx = document.getElementById(chartId).getContext('2d');
    charts[chartId] = new Chart(ctx, {
        type: 'line',
        data: data,
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    min: 1,
                    max: 6,
                    reverse: true,
                    title: {
                        display: true,
                        text: 'Bewertung',
                        font: {
                            size: 14,
                            weight: 'bold'
                        }
                    },
                    ticks: {
                        stepSize: 1
                    }
                },
                x: {
                    ticks: {
                        autoSkip: false,
                        maxRotation: 50,
                        minRotation: 50
                    }
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: title,
                    font: {
                        size: 16,
                        weight: 'bold'
                    }
                },
                legend: {
                    display: true,
                    position: 'bottom'
                },
                tooltip: {
                    callbacks: {
                        label: function (context) {
                            return `Bewertung: ${context.raw}`;
                        }
                    }
                }
            }
        }
    });
}

/**
 * Erstellt ein Radar-Diagramm mit Chart.js.
 * @param {string} chartId - Die ID des Canvas-Elements.
 * @param {string[]} labels - Die Labels für die Achsen des Diagramms.
 * @param {object[]} datasets - Die Datensätze für das Diagramm.
 * @param {string} title - Der Titel des Diagramms.
 */
function createRadarChart(chartId, labels, datasets, title) {
    const ctx = document.getElementById(chartId).getContext('2d');
    charts[chartId] = new Chart(ctx, {
        type: 'radar',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                r: {
                    angleLines: {
                        display: true
                    },
                    suggestedMin: 1,
                    suggestedMax: 6,
                    reverse: true,
                    ticks: {
                        stepSize: 1,
                        backdropColor: 'rgba(255, 255, 255, 0.75)'
                    },
                    pointLabels: {
                        font: {
                            size: 12,
                            weight: 'bold'
                        }
                    }
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: title,
                    font: {
                        size: 16,
                        weight: 'bold'
                    }
                },
                legend: {
                    display: true,
                    position: 'bottom'
                }
            }
        }
    });
}

/**
 * Hilfsfunktion zum Formatieren von Datumsangaben
 * @param {Date|string|number} date - Das zu formatierende Datum.
 * @returns {string} - Das formatierte Datum im Format 'dd.mm.yyyy'.
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
 * Verarbeitet den Excel-Datenimport und die -verarbeitung.
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
            showError("Keine Teilnehmerdaten gefunden. Bitte überprüfen Sie die Struktur der Excel-Datei.");
            return;
        }

        // UI aktualisieren
        uploadArea.classList.add('hidden');
        participantList.classList.remove('hidden');
        dropArea.classList.remove('hidden');
        mainNavigation.classList.remove('hidden');

        renderParticipantList();
        setupParticipantDragDrop();

    } catch (error) {
        console.error("Fehler beim Verarbeiten der Datei:", error);
        showError(`Fehler beim Verarbeiten der Datei: ${error.message}`);
    } finally {
        hideLoading();
    }
}

/**
 * Verarbeitet die Daten aus dem Excel-Workbook.
 * @param {XLSX.WorkBook} workbook - Das Excel-Workbook-Objekt.
 * @returns {object} - Die verarbeiteten Daten.
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

    console.log("Verfügbare Tabellenblätter:", workbook.SheetNames);

    // Tabellenblätter finden (mit Flexibilität für unterschiedliche Benennungen)
    const sheetNames = workbook.SheetNames;
    const teilnehmerSheet = findSheet(sheetNames, ['Teilnehmerliste', 'Teilnehmer', 'TN_Liste', 'TN Liste']);
    const schulischSheet = findSheet(sheetNames, ['Schulische_Basiskompetenzen', 'Schulisch', 'Basiskompetenzen', 'Schulische']);
    const selbstFremdSheet = findSheet(sheetNames, ['Selbst_Fremdeinschaetzung', 'Selbsteinschätzung', 'Kompetenzen', 'Selbst_Fremd']);
    const berufsorientierungSheet = findSheet(sheetNames, ['Berufsorientierung', 'Berufe', 'Berufswahl', 'Berufsfelder']);
    const anmerkungenSheet = findSheet(sheetNames, ['Anmerkungen_Ressourcen', 'Anmerkungen', 'Ressourcen', 'Faktoren']);

    if (!teilnehmerSheet) {
        throw new Error("Keine Teilnehmerliste gefunden. Bitte überprüfen Sie die Tabellenblätter.");
    }

    // Teilnehmerliste verarbeiten
    try {
        processTeilnehmerSheet(workbook.Sheets[teilnehmerSheet], result);

        // Weitere Tabellenblätter verarbeiten, falls vorhanden
        if (schulischSheet) {
            processSchulischSheet(workbook.Sheets[schulischSheet], result);
        }

        if (selbstFremdSheet) {
            processSelbstFremdSheet(workbook.Sheets[selbstFremdSheet], result);
        }

        if (berufsorientierungSheet) {
            processBerufsorientierungSheet(workbook.Sheets[berufsorientierungSheet], result);
        }

        if (anmerkungenSheet) {
            processAnmerkungenSheet(workbook.Sheets[anmerkungenSheet], result);
        }

        return result;
    } catch (error) {
        throw new Error("Fehler bei der Datenverarbeitung: " + error.message);
    }
}

// Hilfsfunktion zum Finden eines Tabellenblatts mit flexiblen Namen
function findSheet(sheetNames, possibleNames) {
    for (const name of possibleNames) {
        const foundSheet = sheetNames.find(sheet =>
            sheet.toLowerCase().includes(name.toLowerCase())
        );
        if (foundSheet) return foundSheet;
    }
    return -1;
}

/**
 * Verarbeitet das Teilnehmerliste-Tabellenblatt.
 * @param {XLSX.WorkSheet} sheet - Das Worksheet-Objekt für die Teilnehmerliste.
 * @param {object} result - Das Ergebnisobjekt, in das die Daten gespeichert werden.
 */
function processTeilnehmerSheet(sheet, result) {
    const data = XLSX.utils.sheet_to_json(sheet, {header: 1});

    if (data.length <= 1) {
        throw new Error("Teilnehmerliste enthält keine Daten.");
    }

    // Header-Zeile finden und Spaltenindizes bestimmen
    const headerRow = data[0];
    const idIndex = findColumnIndex(headerRow, ['TN-ID', 'TNID', 'ID', 'Teilnehmer-ID']);
    const nameIndex = findColumnIndex(headerRow, ['Name', 'Nachname', 'Teilnehmer']);
    const birthIndex = findColumnIndex(headerRow, ['Geburtsdatum', 'Geb.dat', 'Geboren']);
    const dateIndex = findColumnIndex(headerRow, ['Datum', 'Datum der Erstellung', 'Erstelldatum']);
    const authorIndex = findColumnIndex(headerRow, ['Erstellt von', 'Autor', 'Betreuer']);
    const measureIndex = findColumnIndex(headerRow, ['Maßnahme', 'Massnahme', 'Maßnahmenart']);

    if (idIndex === -1) {
        throw new Error("Spalte für TN-ID nicht gefunden.");
    }

    // Teilnehmerdaten extrahieren
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.length > idIndex && row[idIndex]) {
            result.teilnehmer.push({
                id: row[idIndex],
                name: nameIndex !== -1 ? row[nameIndex] || '' : '',
                birthdate: birthIndex !== -1 ? formatDate(row[birthIndex]) : '',
                date: dateIndex !== -1 ? formatDate(row[dateIndex]) : '',
                author: authorIndex !== -1 ? row[authorIndex] || '' : '',
                measure: measureIndex !== -1 ? row[measureIndex] || '' : ''
            });
        }
    }
}

/**
 * Hilfsfunktion zum Finden des Spaltenindex
 */
function findColumnIndex(headerRow, possibleNames) {
    for (const name of possibleNames) {
        const index = headerRow.findIndex(header =>
            header && typeof header === 'string' &&
            header.toLowerCase().includes(name.toLowerCase())
        );
        if (index !== -1) return index;
    }
    return -1;
}

/**
 * Verarbeitet das Schulische Basiskompetenzen-Tabellenblatt.
 */
function processSchulischSheet(sheet, result) {
    const data = XLSX.utils.sheet_to_json(sheet);

    data.forEach(row => {
        const tnId = row['TN-ID'] || row['TNID'] || row['ID'];
        if (!tnId) return;

        if (!result.schulisch[tnId]) {
            result.schulisch[tnId] = {
                kompetenzen: {},
                anmerkungen: ''
            };
        }

        // Alle anderen Spalten als Kompetenzen betrachten
        Object.keys(row).forEach(key => {
            if (key !== 'TN-ID' && key !== 'TNID' && key !== 'ID' && key !== 'Anmerkungen') {
                result.schulisch[tnId].kompetenzen[key] = row[key];
            } else if (key === 'Anmerkungen') {
                result.schulisch[tnId].anmerkungen = row[key] || '';
            }
        });
    });
}

/**
 * Verarbeitet das Selbst- und Fremdeinschätzung-Tabellenblatt.
 */
function processSelbstFremdSheet(sheet, result) {
    const data = XLSX.utils.sheet_to_json(sheet);

    data.forEach(row => {
        const tnId = row['TN-ID'] || row['TNID'] || row['ID'];
        if (!tnId) return;

        const bereich = row['Kompetenzbereich'] || row['Bereich'] || '';
        const kategorie = row['Kategorie'] || row['Kompetenz'] || '';

        if (!bereich || !kategorie) return;

        if (!result.selbstFremd[tnId]) {
            result.selbstFremd[tnId] = {};
        }

        if (!result.selbstFremd[tnId][bereich]) {
            result.selbstFremd[tnId][bereich] = {};
        }

        result.selbstFremd[tnId][bereich][kategorie] = {
            selbst: row['Selbsteinschätzung'] || row['Selbst'] || 0,
            fremd: row['Fremdeinschätzung'] || row['Fremd'] || 0
        };
    });
}

/**
 * Verarbeitet das Berufsorientierung-Tabellenblatt.
 */
function processBerufsorientierungSheet(sheet, result) {
    const data = XLSX.utils.sheet_to_json(sheet);

    data.forEach(row => {
        const tnId = row['TN-ID'] || row['TNID'] || row['ID'];
        if (!tnId) return;

        const bereich = row['Bereich'] || '';
        const kategorie = row['Kategorie'] || '';

        if (!kategorie) return;

        // Berufswahlkompetenz
        if (bereich.toLowerCase().includes('berufswahl') || row['Interesssen'] || row['Vorkenntnisse']) {
            if (!result.berufswahl[tnId]) {
                result.berufswahl[tnId] = {};
            }

            const kategorieName = kategorie || (row['Interesssen'] ? 'Interessen' : 'Vorkenntnisse');

            result.berufswahl[tnId][kategorieName] = {
                handel: row['Handel'] || '',
                lager: row['Lager/Logistik'] || row['Lager'] || '',
                metall: row['Metall'] || '',
                elektro: row['Elektro'] || ''
            };
        }
        // Erprobung
        else if (bereich.toLowerCase().includes('erprobung') || row['Erprobung']) {
            if (!result.erprobung[tnId]) {
                result.erprobung[tnId] = {};
            }

            result.erprobung[tnId][kategorie] = {
                handel: row['Handel'] || 0,
                lager: row['Lager/Logistik'] || row['Lager'] || 0,
                metall: row['Metall'] || 0,
                elektro: row['Elektro'] || 0
            };
        }
        // Selbsteinschätzung
        else if (bereich.toLowerCase().includes('selbst') || row['Selbsteinschätzung'] || row['Selbst_Handel']) {
            if (!result.selbstBerufe[tnId]) {
                result.selbstBerufe[tnId] = {};
            }

            result.selbstBerufe[tnId][kategorie] = {
                handel: row['Handel'] || row['Selbst_Handel'] || 0,
                lager: row['Lager/Logistik'] || row['Lager'] || row['Selbst_Lager'] || 0,
                metall: row['Metall'] || row['Selbst_Metall'] || 0,
                elektro: row['Elektro'] || row['Selbst_Elektro'] || 0
            };
        }
    });
}

/**
 * Verarbeitet das Anmerkungen und Ressourcen-Tabellenblatt.
 */
function processAnmerkungenSheet(sheet, result) {
    const data = XLSX.utils.sheet_to_json(sheet);

    data.forEach(row => {
        const tnId = row['TN-ID'] || row['TNID'] || row['ID'];
        if (!tnId) return;

        result.ressourcen[tnId] = {
            anmerkungen: row['Anmerkungen Berufsorientierung'] || row['Anmerkungen'] || '',
            ressourcen: row['Ressourcen des TN'] || row['Ressourcen'] || '',
            faktoren: row['Hinderliche Faktoren'] || row['Faktoren'] || ''
        };
    });
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

        const displayName = participant.name || 'Unbenannt';
        const displayMeasure = participant.measure || 'Keine Maßnahme';
        const displayDate = participant.birthdate || 'Kein Datum';

        item.innerHTML = `
            <strong>${displayName}</strong>
            <span style="margin-left: 10px; font-size: 0.9em; color: #666;">
                (${displayMeasure}, ${displayDate})
            </span>
        `;

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
        const participant = participants.find(p => p.id === tnId);

        if (participant) {
            showLoading();
            currentParticipant = participant;

            //Kurze Verzögerung, um den Ladezustand anzuzeigen
            setTimeout(() => {
                loadParticipantData(participant);

                dropText.textContent = `Teilnehmer: ${participant.name || 'Unbenannt'}`;
                analysisSection.classList.remove('hidden');

                // Scroll zum Analysebereich
                analysisSection.scrollIntoView({ behavior: 'smooth' });

            }, 100);

            hideLoading();
        }
    });
}

/**
 * Lädt die Daten für den ausgewählten Teilnehmer und aktualisiert die UI.
 * @param {object} participant - Das Teilnehmerobjekt.
 */
function loadParticipantData(participant) {
    // Teilnehmerinfo aktualisieren
    document.getElementById('participantID').textContent = participant.id || '';
    document.getElementById('participantName').textContent = participant.name || '';
    document.getElementById('participantBirthdate').textContent = participant.birthdate || '';
    document.getElementById('participantMeasure').textContent = participant.measure || '';
    document.getElementById('assessmentDate').textContent = participant.date || '';
    document.getElementById('assessmentAuthor').textContent = participant.author || '';

    // Vorhandene Diagramme zerstören, um Fehler zu vermeiden
    Object.keys(charts).forEach(chartId => {
        destroyChart(chartId);
    });

    // Daten für die verschiedenen Tabs laden
    loadSchulischeDaten(participant.id);
    loadSelbstFremdDaten(participant.id);
    loadBerufswahlDaten(participant.id);
    loadErprobungDaten(participant.id);
    loadSelbstBerufeDaten(participant.id);
    loadRessourcenDaten(participant.id);
}

/**
 * Lädt die Schulische Basiskompetenzen-Daten und zeigt sie an.
 * @param {string} tnId - Die Teilnehmer-ID.
 */
function loadSchulischeDaten(tnId) {
    const tableBody = document.querySelector('#schulischTable tbody');
    tableBody.innerHTML = '';

    const schulischData = excelData.schulisch[tnId];
    if (!schulischData) {
        tableBody.innerHTML = '<tr><td colspan="2">Keine Daten verfügbar</td></tr>';
        destroyChart('schulischChart');
        return;
    }

    const chartData = {
        labels: [],
        datasets: [{
            label: 'Bewertung',
            data: [],
            borderColor: '#FF6700',
            backgroundColor: 'rgba(255, 103, 0, 0.1)',
            tension: 0.4,
            pointBackgroundColor: '#FF6700',
            pointRadius: 5,
            borderWidth: 2
        }]
    };

    // Anmerkungen anzeigen
    document.getElementById('schulischAnmerkungen').textContent = schulischData.anmerkungen || 'Keine Anmerkungen vorhanden';

    // Kompetenzen anzeigen
    Object.keys(schulischData.kompetenzen).forEach(key => {
        const value = schulischData.kompetenzen[key];

        // Tabelle aktualisieren
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${key}</td>
            <td class="score-${value}">${value}</td>
        `;
        tableBody.appendChild(row);

        // Diagrammdaten aktualisieren
        chartData.labels.push(key);
        chartData.datasets[0].data.push(value);
    });

    // Diagramm erstellen
    createResponsiveLineChart('schulischChart', chartData, 'Schulische Basiskompetenzen');
}

/**
 * Lädt die Selbst- und Fremdeinschätzung-Daten und zeigt sie an.
 * @param {string} tnId - Die Teilnehmer-ID.
 */
function loadSelbstFremdDaten(tnId) {
    const tableBody = document.querySelector('#selbstfremdTable tbody');
    tableBody.innerHTML = '';

    const selbstFremdData = excelData.selbstFremd[tnId];
    if (!selbstFremdData) {
        tableBody.innerHTML = '<tr><td colspan="3">Keine Daten verfügbar</td></tr>';
        destroyChart('selbstfremdChart');
        return;
    }

    const chartData = {
        labels: [],
        datasets: [
            {
                label: 'Selbsteinschätzung',
                data: [],
                borderColor: '#FF6700',
                backgroundColor: 'rgba(255, 103, 0, 0.1)',
                tension: 0.4,
                pointBackgroundColor: '#FF6700',
                pointRadius: 5,
                borderWidth: 2
            },
            {
                label: 'Fremdeinschätzung',
                data: [],
                borderColor: '#4a6da7',
                backgroundColor: 'rgba(74, 109, 167, 0.1)',
                tension: 0.4,
                pointBackgroundColor: '#4a6da7',
                pointRadius: 5,
                borderWidth: 2
            }
        ]
    };

    // Bereiche durchlaufen
    const bereiche = Object.keys(selbstFremdData);
    bereiche.forEach(bereich => {
        // Bereichsüberschrift
        const headerRow = document.createElement('tr');
        headerRow.className = 'category-header';
        headerRow.innerHTML = `<td colspan="3">${bereich}</td>`;
        tableBody.appendChild(headerRow);

        // Kategorien im Bereich
        const kategorien = Object.keys(selbstFremdData[bereich]);
        kategorien.forEach(kategorie => {
            const data = selbstFremdData[bereich][kategorie];

            // Tabelle aktualisieren
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${kategorie}</td>
                <td class="score-${data.selbst}">${data.selbst}</td>
                <td class="score-${data.fremd}">${data.fremd}</td>
            `;
            tableBody.appendChild(row);

            // Diagrammdaten aktualisieren
            chartData.labels.push(kategorie);
            chartData.datasets[0].data.push(data.selbst);
            chartData.datasets[1].data.push(data.fremd);
        });
    });

    // Diagramm erstellen
    createResponsiveLineChart('selbstfremdChart', chartData, 'Selbst- und Fremdeinschätzung im Vergleich');
}

/**
 * Lädt die Berufswahlkompetenz-Daten und zeigt sie an.
 * @param {string} tnId - Die Teilnehmer-ID.
 */
function loadBerufswahlDaten(tnId) {
    const tableBody = document.querySelector('#berufswahlTable tbody');
    tableBody.innerHTML = '';

    const berufswahlData = excelData.berufswahl[tnId];
    if (!berufswahlData) {
        tableBody.innerHTML = '<tr><td colspan="5">Keine Daten verfügbar</td></tr>';
        return;
    }

    // Berufswahlkompetenz anzeigen
    Object.keys(berufswahlData).forEach(kategorie => {
        const data = berufswahlData[kategorie];

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${kategorie}</td>
            <td>${data.handel}</td>
            <td>${data.lager}</td>
            <td>${data.metall}</td>
            <td>${data.elektro}</td>
        `;
        tableBody.appendChild(row);
    });

    // Anmerkungen anzeigen, falls vorhanden
    if (excelData.ressourcen[tnId] && excelData.ressourcen[tnId].anmerkungen) {
        document.getElementById('berufswahlAnmerkungen').textContent = excelData.ressourcen[tnId].anmerkungen;
    } else {
        document.getElementById('berufswahlAnmerkungen').textContent = 'Keine Anmerkungen vorhanden';
    }
}

/**
 * Lädt die Erprobung Berufsbereiche-Daten und zeigt sie an.
 * @param {string} tnId - Die Teilnehmer-ID.
 */
function loadErprobungDaten(tnId) {
    const tableBody = document.querySelector('#erprobungTable tbody');
    tableBody.innerHTML = '';

    const erprobungData = excelData.erprobung[tnId];
    if (!erprobungData) {
        tableBody.innerHTML = '<tr><td colspan="5">Keine Daten verfügbar</td></tr>';
        destroyChart('erprobungChart');
        return;
    }

    // Radardiagramm-Daten vorbereiten
    const chartLabels = Object.keys(erprobungData);
    const chartDatasets = [
        {
            label: 'Handel',
            data: [],
            backgroundColor: 'rgba(255, 103, 0, 0.2)',
            borderColor: 'rgba(255, 103, 0, 1)',
            borderWidth: 1,
            pointBackgroundColor: 'rgba(255, 103, 0, 1)'
        },
        {
            label: 'Lager/Logistik',
            data: [],
            backgroundColor: 'rgba(74, 109, 167, 0.2)',
            borderColor: 'rgba(74, 109, 167, 1)',
            borderWidth: 1,
            pointBackgroundColor: 'rgba(74, 109, 167, 1)'
        },
        {
            label: 'Metall',
            data: [],
            backgroundColor: 'rgba(104, 104, 104, 0.2)',
            borderColor: 'rgba(104, 104, 104, 1)',
            borderWidth: 1,
            pointBackgroundColor: 'rgba(104, 104, 104, 1)'
        },
        {
            label: 'Elektro',
            data: [],
            backgroundColor: 'rgba(75, 192, 192, 0.2)',
            borderColor: 'rgba(75, 192, 192, 1)',
            borderWidth: 1,
            pointBackgroundColor: 'rgba(75, 192, 192, 1)'
        }
    ];

    // Daten für Tabelle und Diagramm aufbereiten
    chartLabels.forEach(kategorie => {
        const data = erprobungData[kategorie];

        // Tabelle aktualisieren
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${kategorie}</td>
            <td class="score-${data.handel}">${data.handel}</td>
            <td class="score-${data.lager}">${data.lager}</td>
            <td class="score-${data.metall}">${data.metall}</td>
            <td class="score-${data.elektro}">${data.elektro}</td>
        `;
        tableBody.appendChild(row);

        // Diagrammdaten aktualisieren
        chartDatasets[0].data.push(data.handel);
        chartDatasets[1].data.push(data.lager);
        chartDatasets[2].data.push(data.metall);
        chartDatasets[3].data.push(data.elektro);
    });

    // Radar-Diagramm erstellen
    createRadarChart('erprobungChart', chartLabels, chartDatasets, 'Erprobung Berufsbereiche');
}

/**
 * Lädt die Selbsteinschätzung Berufsbereiche-Daten und zeigt sie an.
 * @param {string} tnId - Die Teilnehmer-ID.
 */
function loadSelbstBerufeDaten(tnId) {
    const tableBody = document.querySelector('#selbstberufeTable tbody');
    tableBody.innerHTML = '';

    const selbstBerufeData = excelData.selbstBerufe[tnId];
    if (!selbstBerufeData) {
        tableBody.innerHTML = '<tr><td colspan="5">Keine Daten verfügbar</td></tr>';
        destroyChart('selbstberufeChart');
        return;
    }

    // Radardiagramm-Daten vorbereiten
    const chartLabels = Object.keys(selbstBerufeData);
    const chartDatasets = [
        {
            label: 'Handel',
            data: [],
            backgroundColor: 'rgba(255, 103, 0, 0.2)',
            borderColor: 'rgba(255, 103, 0, 1)',
            borderWidth: 1,
            pointBackgroundColor: 'rgba(255, 103, 0, 1)'
        },
        {
            label: 'Lager/Logistik',
            data: [],
            backgroundColor: 'rgba(74, 109, 167, 0.2)',
            borderColor: 'rgba(74, 109, 167, 1)',
            borderWidth: 1,
            pointBackgroundColor: 'rgba(74, 109, 167, 1)'
        },
        {
            label: 'Metall',
            data: [],
            backgroundColor: 'rgba(104, 104, 104, 0.2)',
            borderColor: 'rgba(104, 104, 104, 1)',
            borderWidth: 1            
            pointBackgroundColor: 'rgba(104, 104, 104, 1)'
        },
        {
            label: 'Elektro',
            data: [],
            backgroundColor: 'rgba(75, 192, 192, 0.2)',
            borderColor: 'rgba(75, 192, 192, 1)',
            borderWidth: 1,
            pointBackgroundColor: 'rgba(75, 192, 192, 1)'
        }
    ];

    // Daten für Tabelle und Diagramm aufbereiten
    chartLabels.forEach(kategorie => {
        const data = selbstBerufeData[kategorie];

        // Tabelle aktualisieren
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${kategorie}</td>
            <td class="score-${data.handel}">${data.handel}</td>
            <td class="score-${data.lager}">${data.lager}</td>
            <td class="score-${data.metall}">${data.metall}</td>
            <td class="score-${data.elektro}">${data.elektro}</td>
        `;
        tableBody.appendChild(row);

        // Diagrammdaten aktualisieren
        chartDatasets[0].data.push(data.handel);
        chartDatasets[1].data.push(data.lager);
        chartDatasets[2].data.push(data.metall);
        chartDatasets[3].data.push(data.elektro);
    });

    // Radar-Diagramm erstellen
    createRadarChart('selbstberufeChart', chartLabels, chartDatasets, 'Selbsteinschätzung Berufsbereiche');

    // Anmerkungen anzeigen, falls vorhanden
    if (excelData.ressourcen[tnId] && excelData.ressourcen[tnId].anmerkungen) {
        document.getElementById('berufsbereicheAnmerkungen').textContent = excelData.ressourcen[tnId].anmerkungen;
    } else {
        document.getElementById('berufsbereicheAnmerkungen').textContent = 'Keine Anmerkungen vorhanden';
    }
}

/**
 * Lädt die Ressourcen und hinderliche Faktoren-Daten und zeigt sie an.
 * @param {string} tnId - Die Teilnehmer-ID.
 */
function loadRessourcenDaten(tnId) {
    const ressourcenData = excelData.ressourcen[tnId];

    if (!ressourcenData) {
        document.getElementById('ressourcenText').textContent = 'Keine Daten verfügbar';
        document.getElementById('faktorenText').textContent = 'Keine Daten verfügbar';
        return;
    }

    document.getElementById('ressourcenText').textContent = ressourcenData.ressourcen || 'Keine Angaben';
    document.getElementById('faktorenText').textContent = ressourcenData.faktoren || 'Keine Angaben';
}

/**
 * Erzeugt HTML für Drucken
 */

/**
 * Bericht drucken
 */
function printReport() {
    if (!currentParticipant) {
        alert('Bitte wählen Sie zuerst einen Teilnehmer aus.');
        return;
    }

    // Print-Ansicht vorbereiten
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Kompetenzanalyse - Stärken-Schwächen Profil</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    margin: 20px;
                }
                h1, h2, h3 {
                    color: #4a6da7;
                }
                .header {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 20px;
                    border-bottom: 1px solid #ddd;
                    padding-bottom: 10px;
                }
                .logo-container {
                    display: flex;
                    align-items: center;
                }
                .logo {
                    max-height: 50px;
                    margin-right: 15px;
                }
                .title {
                    font-size: 24px;
                    font-weight: bold;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 20px;
                }
                th, td {
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: left;
                }
                th {
                    background-color: #f2f2f2;
                }
                .section {
                    margin-bottom: 30px;
                    page-break-inside: avoid;
                }
                .participant-info {
                    display: grid;
                    grid-template-columns: 1fr 1fr;
                    gap: 10px;
                    margin-bottom: 20px;
                }
                .participant-info-item {
                    display: flex;
                }
                .participant-info-label {
                    font-weight: bold;
                    min-width: 150px;
                }
                .category-header {
                    background-color: #e9ecef;
                    font-weight: bold;
                }
                .score-1, .score-2 {
                    background-color: #90EE90;
                }
                .score-3, .score-4 {
                    background-color: #FFD580;
                }
                .score-5, .score-6 {
                    background-color: #FF9999;
                }
                @media print {
                    .page-break {
                        page-break-before: always;
                    }
                }
            </style>
        </head>
        <body>
            <div class="header">
                <div class="logo-container">
                    <img src="https://github.com/Thimeus/BBB/blob/main/FO%CC%88G%20LOGO.png?raw=true" alt="Logo 1" class="logo">
                    <img src="https://github.com/Thimeus/BBB/blob/main/logo_bbb_1.png?raw=true" alt="Logo 2" class="logo">
                </div>
                <div class="title">Kompetenzanalyse Stärken-Schwächen Profil</div>
            </div>

            <div class="participant-info">
                <div class="participant-info-item">
                    <span class="participant-info-label">TN-ID:</span>
                    <span>${currentParticipant.id || '-'}</span>
                </div>
                <div class="participant-info-item">
                    <span class="participant-info-label">Name:</span>
                    <span>${currentParticipant.name || '-'}</span>
                </div>
                <div class="participant-info-item">
                    <span class="participant-info-label">Maßnahme:</span>
                    <span>${currentParticipant.measure || '-'}</span>
                </div>
                <div class="participant-info-item">
                    <span class="participant-info-label">Geburtsdatum:</span>
                    <span>${currentParticipant.birthdate || '-'}</span>
                </div>
                <div class="participant-info-item">
                    <span class="participant-info-label">Datum der Erstellung:</span>
                    <span>${currentParticipant.date || '-'}</span>
                </div>
                <div class="participant-info-item">
                    <span class="participant-info-label">Erstellt von:</span>
                    <span>${currentParticipant.author || '-'}</span>
                </div>
            </div>
    `);

    // Schulische Basiskompetenzen
    printWindow.document.write(`
        <div class="section">
            <h2>I. Schulische Basiskompetenzen</h2>
            <table>
                <tr>
                    <th>Kompetenzen</th>
                    <th>Bewertung</th>
                </tr>
                ${generateSchulischTableHTML(currentParticipant.id)}
            </table>
            <p><strong>Anmerkungen:</strong> ${getSchulischAnmerkungen(currentParticipant.id)}</p>
        </div>
    `);

    // Selbst- und Fremdeinschätzung
    printWindow.document.write(`
        <div class="section page-break">
            <h2>II. Selbst- und Fremdeinschätzung</h2>
            <table>
                <tr>
                    <th>Kompetenzen</th>
                    <th>Selbsteinschätzung</th>
                    <th>Fremdeinschätzung</th>
                </tr>
                ${generateSelbstFremdTableHTML(currentParticipant.id)}
            </table>
        </div>
    `);

    // Berufsorientierung
    printWindow.document.write(`
        <div class="section page-break">
            <h2>III. Berufsorientierung/Berufswahl</h2>
            <h3>Berufswahlkompetenz</h3>
            <table>
                <tr>
                    <th></th>
                    <th>Handel</th>
                    <th>Lager/Logistik</th>
                    <th>Metall</th>
                    <th>Elektro</th>
                </tr>
                ${generateBerufswahlTableHTML(currentParticipant.id)}
            </table>

            <h3>Erprobung Berufsbereiche</h3>
            <table>
                <tr>
                    <th>Kompetenzen</th>
                    <th>Handel</th>
                    <th>Lager/Logistik</th>
                    <th>Metall</th>
                    <th>Elektro</th>
                </tr>
                ${generateErprobungTableHTML(currentParticipant.id)}
            </table>

            <h3>Selbsteinschätzung Berufsbereiche</h3>
            <table>
                <tr>
                    <th>Kompetenzen</th>
                    <th>Handel</th>
                    <th>Lager/Logistik</th>
                    <th>Metall</th>
                    <th>Elektro</th>
                </tr>
                ${generateSelbstBerufeTableHTML(currentParticipant.id)}
            </table>
        </div>
    `);

    // Ressourcen und Faktoren
    printWindow.document.write(`
        <div class="section page-break">
            <h2>IV. Ressourcen und hinderliche Faktoren</h2>
            <table>
                <tr>
                    <th>Ressourcen des TN:</th>
                    <td>${getRessourcenText(currentParticipant.id)}</td>
                </tr>
                <tr>
                    <th>Hinderliche Faktoren:</th>
                    <td>${getFaktorenText(currentParticipant.id)}</td>
                </tr>
            </table>
        </div>

        <div class="section">
            <p style="text-align: center; margin-top: 50px;">
                <i>Kompetenzanalyse-Tool für Berufsvorbereitende Bildungsmaßnahmen</i>
            </p>
        </div>
    `);

    printWindow.document.write(`
        </body>
        </html>
    `);

    printWindow.document.close();
    setTimeout(() => {
        printWindow.print();
    }, 1000);
}

/**
 * Hilfsfunktionen zum Abrufen der Texte
 */
function getSchulischAnmerkungen(tnId) {
    const schulischData = excelData.schulisch[tnId];
    return schulischData && schulischData.anmerkungen ? schulischData.anmerkungen : 'Keine Anmerkungen vorhanden';
}

function getRessourcenText(tnId) {
    const ressourcenData = excelData.ressourcen[tnId];
    return ressourcenData && ressourcenData.ressourcen ? ressourcenData.ressourcen : 'Keine Angaben';
}

function getFaktorenText(tnId) {
    const ressourcenData = excelData.ressourcen[tnId];
    return ressourcenData && ressourcenData.faktoren ? ressourcenData.faktoren : 'Keine Angaben';
}

/**
 * PDF Export
 */
function exportToPdf() {
    // Nutzer informieren, dass die PDF-Export-Funktion über den Druckdialog genutzt werden kann
    alert('Bitte nutzen Sie die Druckfunktion und wählen Sie "Als PDF speichern" im Druckdialog, um ein PDF zu erstellen.');

    // Druckfunktion aufrufen
    printReport();
}

/**
 * UI Table funktionen für Druck
 */
function generateSelbstFremdTableHTML(tnId) {
    const selbstFremdData = excelData.selbstFremd[tnId];
    if (!selbstFremdData) return '<tr><td colspan="3">Keine Daten verfügbar</td></tr>';

    let html = '';

    // Bereiche durchlaufen
    const bereiche = Object.keys(selbstFremdData);
    bereiche.forEach(bereich => {
        // Bereichsüberschrift
        html += `<tr class="category-header"><td colspan="3">${bereich}</td></tr>`;

        // Kategorien im Bereich
        const kategorien = Object.keys(selbstFremdData[bereich]);
        kategorien.forEach(kategorie => {
            const data = selbstFremdData[bereich][kategorie];
            html += `
                <tr>
                    <td>${kategorie}</td>
                    <td class="score-${data.selbst}">${data.selbst}</td>
                    <td class="score-${data.fremd}">${data.fremd}</td>
                </tr>
            `;
        });
    });

    return html;
}

function generateBerufswahlTableHTML(tnId) {
    const berufswahlData = excelData.berufswahl[tnId];
    if (!berufswahlData) return '<tr><td colspan="5">Keine Daten verfügbar</td></tr>';

    let html = '';
    Object.keys(berufswahlData).forEach(kategorie => {
        const data = berufswahlData[kategorie];
        html += `
            <tr>
                <td>${kategorie}</td>
                <td>${data.handel}</td>
                <td>${data.lager}</td>
                <td>${data.metall}</td>
                <td>${data.elektro}</td>
            </tr>
        `;
    });

    return html;
}

function generateSchulischTableHTML(tnId) {
    const schulischData = excelData.schulisch[tnId];
    if (!schulischData) return '<tr><td colspan="2">Keine Daten verfügbar</td></tr>';

    let html = '';
    Object.keys(schulischData.kompetenzen).forEach(key => {
        const value = schulischData.kompetenzen[key];
        html += `
            <tr>
                <td>${key}</td>
                <td class="score-${value}">${value}</td>
            </tr>
        `;
    });

    return html;
}

function generateErprobungTableHTML(tnId) {
    const erprobungData = excelData.erprobung[tnId];
    if (!erprobungData) return '<tr><td colspan="5">Keine Daten verfügbar</td></tr>';

    let html = '';
    Object.keys(erprobungData).forEach(kategorie => {
        const data = erprobungData[kategorie];
        html += `
            <tr>
                <td>${kategorie}</td>
                <td class="score-${data.handel}">${data.handel}</td>
                <td class="score-${data.lager}">${data.lager}</td>
                <td class="score-${data.metall}">${data.metall}</td>
                <td class="score-${data.elektro}">${data.elektro}</td>
            </tr>
        `;
    });

    return html;
}

function generateSelbstBerufeTableHTML(tnId) {
    const selbstBerufeData = excelData.selbstBerufe[tnId];
    if (!selbstBerufeData) return '<tr><td colspan="5">Keine Daten verfügbar</td></tr>';

    let html = '';
    Object.keys(selbstBerufeData).forEach(kategorie => {
        const data = selbstBerufeData[kategorie];
        html += `
            <tr>
                <td>${kategorie}</td>
                <td class="score-${data.handel}">${data.handel}</td>
                <td class="score-${data.lager}">${data.lager}</td>
                <td class="score-${data.metall}">${data.metall}</td>
                <td class="score-${data.elektro}">${data.elektro}</td>
            </tr>
        `;
    });

    return html;
}

/**
 * Event Listeners
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
 * Fügt Event Listener TabNavigation hinzu.
 */
function setupTabNavigation() {
    document.querySelectorAll('.tab-btn').forEach(button => {
        button.addEventListener('click', () => {
            // Aktive Tab-Klasse entfernen
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });

            // Aktive Tab-Klasse hinzufügen
            button.classList.add('active');
            const tabId = button.getAttribute('data-tab') + 'Tab';
            document.getElementById(tabId).classList.add('active');

            // Nach dem Anzeigen des Tabs die Diagramme neu erstellen
            if (currentParticipant) {
                const tabName = button.getAttribute('data-tab');
                switch (tabName) {
                    case 'schulisch':
                        loadSchulischeDaten(currentParticipant.id);
                        break;
                    case 'selbstfremd':
                        loadSelbstFremdDaten(currentParticipant.id);
                        break;
                    case 'erprobung':
                        loadErprobungDaten(currentParticipant.id);
                        break;
                    case 'selbstberufe':
                        loadSelbstBerufeDaten(currentParticipant.id);
                        break;
                }
            }
        });
    });
}

document.addEventListener('DOMContentLoaded', setupTabNavigation);

/**
 * Fügt Event Listener MainNavigation hinzu.
 */
function setupMainNavigation() {
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            // Aktive Tab-Klasse entfernen
            document.querySelectorAll('.nav-tab').forEach(navTab => {
                navTab.classList.remove('active');
            });

            // Aktive Tab-Klasse hinzufügen
            tab.classList.add('active');

            // Entsprechenden Tab-Inhalt aktivieren
            const tabName = tab.getAttribute('data-tab');

            // Klick auf den entsprechenden Tab-Button simulieren
            document.querySelector(`.tab-btn[data-tab="${tabName}"]`).click();
        });
    });
}

document.addEventListener('DOMContentLoaded', setupMainNavigation);

/**
 * Fügt Event Listener Button hinzu.
 */
document.getElementById('printReportBtn').addEventListener('click', printReport);
document.getElementById('exportPdfBtn').addEventListener('click', exportToPdf);