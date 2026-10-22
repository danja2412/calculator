// Конфігурація зон та тарифних сіток
const ZONES = ['1', '2', '2.1', '3', '4', '4.1', '5', '5.1', '5.2', '6'];

// Глобальний об'єкт для збереження розпарсених даних тарифів
// Структура: { "ТС1": { "A1": { "1": 1200, "2": 1100, ... } } }
let tariffData = {};

// Поточний стан вибору користувача
let selectedGrid = '';
let selectedVehicle = '';
let selectedZone = '';

// DOM Елементи
const excelFileInput = document.getElementById('excelFile');
const fileNameDisplay = document.getElementById('fileName');
const statusMessage = document.getElementById('statusMessage');
const downloadTemplateBtn = document.getElementById('downloadTemplateBtn');

const gridSelect = document.getElementById('gridSelect');
const vehicleSelect = document.getElementById('vehicleSelect');
const zonesContainer = document.getElementById('zonesContainer');
const totalPriceDisplay = document.getElementById('totalPrice');
const calcDetails = document.getElementById('calcDetails');

// Ініціалізація сторінки
document.addEventListener('DOMContentLoaded', () => {
    renderZoneChips();
    initEventListeners();
});

// Малювання кнопок зон
function renderZoneChips() {
    zonesContainer.innerHTML = '';
    ZONES.forEach(zone => {
        const chip = document.createElement('div');
        chip.className = 'zone-chip disabled';
        chip.dataset.zone = zone;
        chip.innerText = `Зона ${zone}`;
        chip.addEventListener('click', () => selectZone(zone));
        zonesContainer.appendChild(chip);
    });
}

// Події
function initEventListeners() {
    excelFileInput.addEventListener('change', handleFileUpload);
    downloadTemplateBtn.addEventListener('click', generateAndDownloadTemplate);
    gridSelect.addEventListener('change', handleGridChange);
    vehicleSelect.addEventListener('change', handleVehicleChange);
}

// 1. Обробка завантаження Excel файлу
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    fileNameDisplay.innerText = file.name;
    showStatus('Зчитування файлу...', 'info');

    const reader = new FileReader();
    reader.onload = function (evt) {
        try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            parseWorkbook(workbook);
            showStatus('Тарифну сітку успішно завантажено!', 'success');
            
            // Активуємо вибір сітки
            gridSelect.disabled = false;
            updateVehicleDropdown();
            calculateTariff();
        } catch (err) {
            console.error(err);
            showStatus('Помилка зчитування Excel файлу. Перевірте формат!', 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Очищення та нормалізація назв аркушів
function normalizeSheetName(name) {
    return name.trim().replace(',', '.');
}

// Парсинг усіх аркушів Excel у внутрішню структуру
function parseWorkbook(workbook) {
    tariffData = {};

    workbook.SheetNames.forEach(sheetName => {
        const normalizedName = normalizeSheetName(sheetName);
        const worksheet = workbook.Sheets[sheetName];
        
        // Преобразуємо аркуш у масив масивів (rows)
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (rows.length < 2) return; // Порожній або лише заголовок

        tariffData[normalizedName] = {};

        // Рядок 0 — заголовки (Тип ТЗ, Зона 1, Зона 2...)
        const headerRow = rows[0].map(val => String(val || '').trim());
        
        // Знаходимо індекси колонок для кожної зони
        const zoneColIndexes = {};
        ZONES.forEach(zone => {
            const colIdx = headerRow.findIndex(h => h === zone || h === `Зона ${zone}` || h === `Зона ${zone.replace('.', ',')}`);
            if (colIdx !== -1) {
                zoneColIndexes[zone] = colIdx;
            }
        });

        // Обробка рядочків даних (з 1-го рядка)
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const vehicleType = String(row[0] || '').trim();
            if (!vehicleType) continue;

            tariffData[normalizedName][vehicleType] = {};

            ZONES.forEach(zone => {
                const colIdx = zoneColIndexes[zone];
                let price = 0;
                if (colIdx !== undefined && row[colIdx] !== undefined) {
                    price = parseFloat(String(row[colIdx]).replace(',', '.')) || 0;
                } else if (i < row.length) {
                    // Якщо назви зон у заголовоку не збіглися, використовуємо позиційний порядковий номер (1..10)
                    const fallbackIndex = ZONES.indexOf(zone) + 1;
                    price = parseFloat(String(row[fallbackIndex] || 0).replace(',', '.')) || 0;
                }
                tariffData[normalizedName][vehicleType][zone] = price;
            });
        }
    });
}

// 2. Обробка вибору сітки
function handleGridChange() {
    selectedGrid = gridSelect.value;
    updateVehicleDropdown();
    calculateTariff();
}

// Оновлення списку типів ТЗ на основі обраної сітки
function updateVehicleDropdown() {
    vehicleSelect.innerHTML = '';
    
    if (!selectedGrid || !tariffData[selectedGrid]) {
        vehicleSelect.disabled = true;
        vehicleSelect.innerHTML = '<option value="">-- Оберіть доступну сітку --</option>';
        disableZoneChips(true);
        return;
    }

    const vehicles = Object.keys(tariffData[selectedGrid]);

    if (vehicles.length === 0) {
        vehicleSelect.disabled = true;
        vehicleSelect.innerHTML = '<option value="">-- Немає даних для цієї сітки --</option>';
        disableZoneChips(true);
        return;
    }

    vehicleSelect.disabled = false;
    vehicleSelect.innerHTML = '<option value="">-- Оберіть тип ТЗ --</option>';
    
    vehicles.forEach(veh => {
        const opt = document.createElement('option');
        opt.value = veh;
        opt.textContent = veh;
        vehicleSelect.appendChild(opt);
    });

    disableZoneChips(false);
}

// 3. Обробка вибору типу ТЗ
function handleVehicleChange() {
    selectedVehicle = vehicleSelect.value;
    calculateTariff();
}

// 4. Обробка вибору Зони
function selectZone(zone) {
    if (gridSelect.disabled || !vehicleSelect.value) return;

    selectedZone = zone;
    
    // Оновлення стилів кнопок зон
    document.querySelectorAll('.zone-chip').forEach(chip => {
        if (chip.dataset.zone === zone) {
            chip.classList.add('selected');
        } else {
            chip.classList.remove('selected');
        }
    });

    calculateTariff();
}

function disableZoneChips(disabled) {
    document.querySelectorAll('.zone-chip').forEach(chip => {
        if (disabled) {
            chip.classList.add('disabled');
            chip.classList.remove('selected');
        } else {
            chip.classList.remove('disabled');
        }
    });
}

// 5. Розрахунок підсумкового тарифу
function calculateTariff() {
    if (!selectedGrid || !selectedVehicle || !selectedZone) {
        totalPriceDisplay.innerHTML = `0.00 <span class="currency">грн</span>`;
        calcDetails.innerText = 'Оберіть тарифну сітку, тип ТЗ та зону для розрахунку.';
        return;
    }

    if (tariffData[selectedGrid] &&
        tariffData[selectedGrid][selectedVehicle] &&
        tariffData[selectedGrid][selectedVehicle][selectedZone] !== undefined) {
        
        const price = tariffData[selectedGrid][selectedVehicle][selectedZone];
        
        totalPriceDisplay.innerHTML = `${price.toLocaleString('uk-UA', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} <span class="currency">грн</span>`;
        calcDetails.innerText = `Сітка: ${selectedGrid} | ТЗ: ${selectedVehicle} | Зона: ${selectedZone}`;
    } else {
        totalPriceDisplay.innerHTML = `0.00 <span class="currency">грн</span>`;
        calcDetails.innerText = 'Тариф за обраними параметрами не знайдено у файлі.';
    }
}

// Статусні повідомлення
function showStatus(msg, type) {
    statusMessage.innerText = msg;
    statusMessage.className = `status-message ${type}`;
}

// 6. Генерація та скачування зразкового шаблону Excel
function generateAndDownloadTemplate() {
    const grids = ['ТС1', 'ТС2', 'ТС1 таксі', 'ТС2 таксі', 'ТС1 МАН', 'ТС2 МАН', 'Знижка 0.9', 'Знижка 0.95'];
    const sampleVehicles = [
        'A1 (Моторолери та мотоцикли до 300 см3)',
        'A2 (Мотоцикли понад 300 см3)',
        'B1 (Легкові до 1600 см3)',
        'B2 (Легкові 1601-2000 см3)',
        'B3 (Легкові 2001-3000 см3)',
        'B4 (Легкові понад 3000 см3)',
        'C1 (Вантажні до 2 т)',
        'C2 (Вантажні понад 2 т)',
        'D1 (Автобуси до 20 місць)',
        'D2 (Автобуси понад 20 місць)',
        'E (Причепи до легкових)',
        'F (Причепи до вантажних)'
    ];

    const wb = XLSX.utils.book_new();

    grids.forEach(gridName => {
        const header = ['Тип ТЗ', ...ZONES];
        const rows = [header];

        sampleVehicles.forEach((veh, idx) => {
            const basePrice = (idx + 1) * 250;
            const row = [veh];
            
            ZONES.forEach((z, zIdx) => {
                // Генерація тестових розрахованих цін залежно від зони
                const zoneMultiplier = 1.5 - (zIdx * 0.08);
                const price = Math.round(basePrice * zoneMultiplier);
                row.push(price);
            });
            
            rows.push(row);
        });

        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, gridName);
    });

    XLSX.writeFile(wb, 'Тарифна_сітка_страхування.xlsx');
}