// Конфігурація 10 зон
const ZONES = ['1', '2', '2.1', '3', '4', '4.1', '5', '5.1', '5.2', '6'];

// Конфігурація доступних сіток залежно від типу особи
const PERSON_GRID_MAP = {
    physical: ['ТС1', 'ТС2', 'ТС1 МАН', 'ТС2 МАН', 'Знижка 0.9', 'Знижка 0.95'],
    legal: ['ТС1', 'ТС2', 'ТС1 таксі', 'ТС2 таксі', 'ТС1 МАН', 'ТС2 МАН', 'Знижка 0.9', 'Знижка 0.95']
};

// Збереження тарифних даних
let tariffData = {};

// Поточний стан вибору
let selectedPersonType = 'physical';
let selectedGrid = '';
let selectedVehicle = '';
let selectedZone = '';

// DOM елементи
const dataStatus = document.getElementById('dataStatus');
const gridSelect = document.getElementById('gridSelect');
const vehicleSelect = document.getElementById('vehicleSelect');
const zonesContainer = document.getElementById('zonesContainer');
const totalPriceDisplay = document.getElementById('totalPrice');
const calcDetails = document.getElementById('calcDetails');

document.addEventListener('DOMContentLoaded', () => {
    renderZoneChips();
    initEventListeners();
    autoLoadTariffs();
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

function initEventListeners() {
    // Подія зміни Фіз/Юр особи
    document.querySelectorAll('input[name="personType"]').forEach(radio => {
        radio.addEventListener('change', handlePersonTypeChange);
    });

    gridSelect.addEventListener('change', handleGridChange);
    vehicleSelect.addEventListener('change', handleVehicleChange);
}

// 1. Автоматичне фонове завантаження `tariffs.xlsx`
async function autoLoadTariffs() {
    try {
        const response = await fetch('tariffs.xlsx');
        if (!response.ok) throw new Error('Файл tariffs.xlsx не знайдено');

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        
        parseWorkbook(workbook);
        updateStatus('Тарифи актуальні (з файлу)', 'success');
    } catch (err) {
        console.warn('Не вдалося завантажити tariffs.xlsx. Підключаємо базові рішення:', err.message);
        loadFallbackTariffs();
        updateStatus('Тестові тарифи (покладіть tariffs.xlsx у репозиторій)', 'warning');
    }

    populateGridDropdown();
}

// Парсинг Excel книги
function parseWorkbook(workbook) {
    tariffData = {};

    workbook.SheetNames.forEach(sheetName => {
        // Нормалізуємо назву аркуша (наприклад, "Знижка 0,95" -> "Знижка 0.95")
        const cleanSheetName = sheetName.trim().replace(',', '.');
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (rows.length < 2) return;

        tariffData[cleanSheetName] = {};
        const headerRow = rows[0].map(val => String(val || '').trim());

        // Пошук колонок за зоною
        const zoneColIndexes = {};
        ZONES.forEach(zone => {
            const colIdx = headerRow.findIndex(h => 
                h === zone || 
                h === `Зона ${zone}` || 
                h === `Зона ${zone.replace('.', ',')}`
            );
            if (colIdx !== -1) zoneColIndexes[zone] = colIdx;
        });

        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const vehicleType = String(row[0] || '').trim();
            if (!vehicleType) continue;

            tariffData[cleanSheetName][vehicleType] = {};

            ZONES.forEach(zone => {
                const colIdx = zoneColIndexes[zone];
                let price = 0;

                if (colIdx !== undefined && row[colIdx] !== undefined) {
                    price = parseFloat(String(row[colIdx]).replace(',', '.')) || 0;
                } else {
                    const fallbackIndex = ZONES.indexOf(zone) + 1;
                    price = parseFloat(String(row[fallbackIndex] || 0).replace(',', '.')) || 0;
                }
                tariffData[cleanSheetName][vehicleType][zone] = price;
            });
        }
    });
}

// Обробка перемикання Фізособа / Юрособа
function handlePersonTypeChange(e) {
    selectedPersonType = e.target.value;
    selectedGrid = '';
    selectedVehicle = '';
    selectedZone = '';

    populateGridDropdown();
    updateVehicleDropdown();
    resetZoneSelection();
    calculateTariff();
}

// Заповнення списку сіток залежно від обраного типу особи
function populateGridDropdown() {
    gridSelect.innerHTML = '<option value="">-- Оберіть тарифну сітку --</option>';
    
    // Отримуємо масив дозволених сіток для даного типу особи
    const allowedGrids = PERSON_GRID_MAP[selectedPersonType] || [];
    
    // Фільтруємо лише ті сітки, які реально є у розпарсеному файлі Excel
    const availableInFile = Object.keys(tariffData);
    const validGrids = allowedGrids.filter(g => availableInFile.includes(g));

    if (validGrids.length === 0) {
        // Якщо точних назв з крапкою не знайшли, виводимо всі дозволені для цієї категорії
        allowedGrids.forEach(grid => {
            const opt = document.createElement('option');
            opt.value = grid;
            opt.textContent = grid;
            gridSelect.appendChild(opt);
        });
        return;
    }

    validGrids.forEach(grid => {
        const opt = document.createElement('option');
        opt.value = grid;
        opt.textContent = grid;
        gridSelect.appendChild(opt);
    });
}

// Обробка зміни сітки
function handleGridChange() {
    selectedGrid = gridSelect.value;
    selectedVehicle = '';
    selectedZone = '';
    updateVehicleDropdown();
    resetZoneSelection();
    calculateTariff();
}

// Оновлення списку ТЗ
function updateVehicleDropdown() {
    vehicleSelect.innerHTML = '';
    
    if (!selectedGrid || !tariffData[selectedGrid]) {
        vehicleSelect.disabled = true;
        vehicleSelect.innerHTML = '<option value="">-- Спочатку оберіть сітку --</option>';
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

function handleVehicleChange() {
    selectedVehicle = vehicleSelect.value;
    calculateTariff();
}

function selectZone(zone) {
    if (gridSelect.disabled || !vehicleSelect.value) return;

    selectedZone = zone;
    document.querySelectorAll('.zone-chip').forEach(chip => {
        chip.classList.toggle('selected', chip.dataset.zone === zone);
    });

    calculateTariff();
}

function resetZoneSelection() {
    selectedZone = '';
    document.querySelectorAll('.zone-chip').forEach(chip => {
        chip.classList.remove('selected');
    });
}

function disableZoneChips(disabled) {
    document.querySelectorAll('.zone-chip').forEach(chip => {
        chip.classList.toggle('disabled', disabled);
        if (disabled) chip.classList.remove('selected');
    });
}

// Підсумковий розрахунок
function calculateTariff() {
    if (!selectedGrid || !selectedVehicle || !selectedZone) {
        totalPriceDisplay.innerHTML = `0.00 <span class="currency">грн</span>`;
        calcDetails.innerText = 'Оберіть тарифну сітку, тип ТЗ та зону для розрахунку.';
        return;
    }

    if (tariffData[selectedGrid]?.[selectedVehicle]?.[selectedZone] !== undefined) {
        const price = tariffData[selectedGrid][selectedVehicle][selectedZone];
        const personTitle = selectedPersonType === 'physical' ? 'Фіз. особа' : 'Юр. особа';

        totalPriceDisplay.innerHTML = `${price.toLocaleString('uk-UA', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} <span class="currency">грн</span>`;
        calcDetails.innerText = `${personTitle} | Сітка: ${selectedGrid} | ТЗ: ${selectedVehicle} | Зона: ${selectedZone}`;
    } else {
        totalPriceDisplay.innerHTML = `0.00 <span class="currency">грн</span>`;
        calcDetails.innerText = 'Тариф за обраними параметрами не знайдено.';
    }
}

function updateStatus(text, className) {
    dataStatus.innerText = text;
    dataStatus.className = `status-badge ${className}`;
}

// Резервний варіант тарифів
function loadFallbackTariffs() {
    tariffData = {};
    const allGrids = ['ТС1', 'ТС2', 'ТС1 таксі', 'ТС2 таксі', 'ТС1 МАН', 'ТС2 МАН', 'Знижка 0.9', 'Знижка 0.95'];
    const fallbackVehicles = [
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

    allGrids.forEach(grid => {
        tariffData[grid] = {};
        fallbackVehicles.forEach((veh, idx) => {
            tariffData[grid][veh] = {};
            const basePrice = (idx + 1) * 300;
            ZONES.forEach((z, zIdx) => {
                tariffData[grid][veh][z] = Math.round(basePrice * (1.6 - zIdx * 0.07));
            });
        });
    });
}