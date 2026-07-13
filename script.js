const fileInput = document.getElementById('excel-file');
const listContainer = document.getElementById('rank-list');
const exportBtn = document.getElementById('export-btn');
let itemsData = [];

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

exportBtn.addEventListener('click', exportRankedList);
function exportRankedList() {
    const rows = listContainer.querySelectorAll('.list-row');
    if (rows.length <= 1) {
        alert("Please load an Excel sheet with choice preferences first!");
        return;
    }

    const exportedData = [];

    // 1. Re-compile the grid layout matrix tracking updated indices
    rows.forEach((row) => {
        const isHeader = row.classList.contains('header-row');
        const code = row.querySelector('.code-id').innerText;
        const name = row.querySelector('.name-text').innerText;
        
        if (isHeader) {
            // Re-apply explicit headers to the output file structure
            exportedData.push(["Final Rank", code, name]);
        } else {
            const rank = row.querySelector('.rank-index').innerText;
            exportedData.push([parseInt(rank, 10), code, name]);
        }
    });

    // 2. Build the updated binary spreadsheet workbook via SheetJS tools
    const newWorksheet = XLSX.utils.aoa_to_sheet(exportedData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Final Ranks");

    // 3. Trigger immediate client-side system download
    const filename = "Final_Ranked_Choices.xlsx";
    XLSX.writeFile(newWorkbook, filename);

    // 4. Draft email draft notification sharing configuration
    // Because file attachments require dedicated server architectures, this builds a direct 
    // mailto configuration link summarizing your updated selections for sharing.
    setTimeout(() => {
        const confirmShare = confirm("Spreadsheet downloaded successfully! Do you want to share a text summary of these ranks via email right now?");
        if (confirmShare) {
            let emailBody = "Here is my final preference choice list:\n\n";
            exportedData.slice(1).forEach(item => {
                emailBody += `Rank ${item[0]}: [${item[1]}] ${item[2]}\n`;
            });
            
            const subject = encodeURIComponent("My Final Preference Choice Ranks");
            const body = encodeURIComponent(emailBody);
            
            window.location.href = `mailto:?subject=${subject}&body=${body}`;
        }
    }, 500);
}