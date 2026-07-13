const fileInput = document.getElementById('excel-file');
const listContainer = document.getElementById('rank-list');
let itemsData = [];

// DOM Selectors for the Modal Elements
const exportBtn = document.getElementById('export-btn');
const modal = document.getElementById('export-modal');
const filenameInput = document.getElementById('filename-input');
const downloadBtn = document.getElementById('modal-download-btn');
const emailBtn = document.getElementById('modal-email-btn');
const whatsappBtn = document.getElementById('modal-wa-btn');
const closeBtn = document.getElementById('modal-close-btn');

// File Upload Event Listener
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        
        itemsData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        itemsData = itemsData.filter(row => row.length > 0);
        
        renderList();
    };
    reader.readAsArrayBuffer(file);
});

// Render the Table List
function renderList() {
    listContainer.innerHTML = '';
    if (itemsData.length === 0) return;

    // 1. Separate the first row as the static column header
    const headerRowData = itemsData[0];
    const headerEl = document.createElement('div');
    headerEl.classList.add('list-row', 'header-row');
    
    headerEl.innerHTML = `
        <div class="rank-number-wrapper">
            <div style="width: 12px;"></div>
            <div class="rank-index">Rank</div>
        </div>
        <div class="row-content">
            <span class="code-id">${headerRowData[0] || 'Code'}</span>
            <span class="name-text">${headerRowData[1] || 'Institute Name'}</span>
        </div>
    `;
    listContainer.appendChild(headerEl);

    // 2. Render all remaining rows as actual interactive draggable items
    const draggableItems = itemsData.slice(1);

    draggableItems.forEach((row, index) => {
        const rowEl = document.createElement('div');
        rowEl.classList.add('list-row', 'draggable-row');
        rowEl.setAttribute('draggable', 'true');

        const handleHtml = `
            <div class="drag-handle">
                <div class="dots-pair"><span></span><span></span></div>
                <div class="dots-pair"><span></span><span></span></div>
                <div class="dots-pair"><span></span><span></span></div>
            </div>
        `;

        rowEl.innerHTML = `
            <div class="rank-number-wrapper">
                ${handleHtml}
                <div class="rank-index">${index + 1}</div>
            </div>
            <div class="row-content">
                <span class="code-id">${row[0] || ''}</span>
                <span class="name-text">${row[1] || ''}</span>
            </div>
        `;

        rowEl.addEventListener('dragstart', () => rowEl.classList.add('dragging'));
        rowEl.addEventListener('dragend', () => rowEl.classList.remove('dragging'));

        listContainer.appendChild(rowEl);
    });
}

// Drag over layout calculations
listContainer.addEventListener('dragover', (e) => {
    e.preventDefault();
    const draggingEl = document.querySelector('.dragging');
    if (!draggingEl) return;

    const afterElement = getDragAfterElement(listContainer, e.clientY);
    
    if (afterElement == null) {
        listContainer.appendChild(draggingEl);
    } else if (afterElement.classList.contains('header-row')) {
        const firstDraggable = listContainer.querySelector('.draggable-row');
        if (firstDraggable && firstDraggable !== draggingEl) {
            listContainer.insertBefore(draggingEl, firstDraggable);
        }
    } else {
        listContainer.insertBefore(draggingEl, afterElement);
    }
    
    updateRankNumbers();
});

function getDragAfterElement(container, y) {
    const elements = [...container.querySelectorAll('.list-row:not(.dragging)')];

    return elements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;
        if (offset < 0 && offset > closest.offset) {
            return { offset: offset, element: child };
        } else {
            return closest;
        }
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

function updateRankNumbers() {
    const rows = listContainer.querySelectorAll('.draggable-row');
    rows.forEach((row, index) => {
        row.querySelector('.rank-index').innerText = index + 1;
    });
}

// Open Modal Interface
exportBtn.addEventListener('click', () => {
    const rows = listContainer.querySelectorAll('.list-row');
    if (rows.length <= 1) {
        alert("Please load an Excel sheet with choice preferences first!");
        return;
    }
    modal.style.display = 'flex';
});

// Close Modal Interface
closeBtn.addEventListener('click', () => modal.style.display = 'none');

// Extract current ranked data out of DOM
function getRankedDataMatrix() {
    const rows = listContainer.querySelectorAll('.list-row');
    const matrix = [];
    rows.forEach((row) => {
        const isHeader = row.classList.contains('header-row');
        const code = row.querySelector('.code-id').innerText;
        const name = row.querySelector('.name-text').innerText;
        
        if (isHeader) {
            matrix.push(["Final Rank", code, name]);
        } else {
            const rank = row.querySelector('.rank-index').innerText;
            matrix.push([parseInt(rank, 10), code, name]);
        }
    });
    return matrix;
}

// Generate text payload summary for updates
function getTextSummary(matrix) {
    let summaryText = "Here is my final preference choice list:\n\n";
    matrix.slice(1).forEach(item => {
        summaryText += `Rank ${item[0]}: [${item[1]}] ${item[2]}\n`;
    });
    return summaryText;
}

// Action 1: Download File with Custom Name
downloadBtn.addEventListener('click', () => {
    const matrix = getRankedDataMatrix();
    let name = filenameInput.value.trim().replace(/[/\\?%*:|"<>]/g, '-');
    if (!name) name = "Final_Ranked_Choices";
    
    const ws = XLSX.utils.aoa_to_sheet(matrix);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Final Ranks");
    
    XLSX.writeFile(wb, `${name}.xlsx`);
    modal.style.display = 'none';
});

// Action 2: Share Text Summary via Email
emailBtn.addEventListener('click', () => {
    const matrix = getRankedDataMatrix();
    const bodyText = getTextSummary(matrix);
    const subject = encodeURIComponent("My Final Preference Choice Ranks");
    const body = encodeURIComponent(bodyText);
    
    window.location.href = `mailto:?subject=${subject}&body=${body}`;
    modal.style.display = 'none';
});

// Action 3: Share Text Summary via WhatsApp Web/App
whatsappBtn.addEventListener('click', () => {
    const matrix = getRankedDataMatrix();
    const bodyText = getTextSummary(matrix);
    const encodedText = encodeURIComponent(bodyText);
    
    window.open(`https://api.whatsapp.com/send?text=${encodedText}`, '_blank');
    modal.style.display = 'none';
});