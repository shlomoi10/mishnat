document.addEventListener('DOMContentLoaded', () => {
    const pasteBtn = document.getElementById('pasteBtn');
    const outputContainer = document.getElementById('outputContainer');
    const modal = document.getElementById('infoModal');
    const modalCloseBtn = document.getElementById('modalCloseBtn');

    pasteBtn.addEventListener('click', handlePaste);

    // Check for 'paste' query parameter on page load
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.has('paste')) {
        handlePaste();
    }

    async function handlePaste() {
        try {
            const jsonText = await navigator.clipboard.readText();
            const apiData = JSON.parse(jsonText);
            await renderOrder(apiData); // Ensure we wait for the entire process
        } catch (err) {
            showModal('שגיאה', 'ההדבקה נכשלה. אנא ודא שהעתקת קובץ JSON תקין.');
            console.error("Error processing pasted data:", err);
        }
    }

    async function renderOrder(apiData) {
        const contentHtml = createPrintContent(apiData);
        outputContainer.innerHTML = contentHtml;
        addEventListeners(apiData);
    }

    function addEventListeners(apiData) {
        document.getElementById('exportExcelBtn').addEventListener('click', () => exportToExcel(apiData));
        document.getElementById('printBtnNew').addEventListener('click', () => window.print());
        document.getElementById('showAdminBtn').addEventListener('click', () => showAdminInfo(apiData));
        
        // Event delegation for station manager buttons
        outputContainer.addEventListener('click', (e) => {
            if (e.target && e.target.classList.contains('station-manager-btn')) {
                const managerName = e.target.dataset.manager;
                const productName = e.target.dataset.product;
                showModal(`אחראי תחנה: ${productName}`, `שם האחראי: <strong>${managerName}</strong>`);
            }
        });
    }

    // --- Modal Logic ---
    function showModal(title, body) {
        document.getElementById('modalTitle').textContent = title;
        document.getElementById('modalBody').innerHTML = body;
        modal.classList.add('is-visible');
    }

    function closeModal() {
        modal.classList.remove('is-visible');
    }

    modalCloseBtn.addEventListener('click', closeModal);
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            closeModal();
        }
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === "Escape" && modal.classList.contains('is-visible')) {
            closeModal();
        }
    });

    // --- Data Processing & HTML Generation ---
    class Product {
        constructor(data, item) {
            this.data = data;
            this.item = item;
        }
        isCut() { return this.data.src_amount && this.data.src_amount > 0; }
        getUnitsText() {
            if (this.data.units_type === 10) return 'גרם';
            if (this.data.units_type === 16) return 'ליטר';
            return 'יחידה';
        }
        getPricePerUnit() {
            if (!this.data.units || this.data.units === 0) return '';
            if (this.data.units_type === 10 && this.data.units < 1) return `(₪${(this.data.price / this.data.units / 10).toFixed(2)} ל-100 גרם)`;
            if (this.data.units_type === 16 && this.data.units < 1) return `(₪${(this.data.price / this.data.units / 10).toFixed(2)} ל-100 מ״ל)`;
            return `(₪${(this.data.price / this.data.units).toFixed(2)} ליחידה)`;
        }
        getStationManagerButton() {
            const manager = this.item.item_sale?.product?.station_manager;
            if (!manager) return '';
            return `<button class="station-manager-btn" data-manager="${manager}" data-product="${this.data.full_name}">פרטים</button>`;
        }
        getQuantityBadge() {
            if (this.data.amount <= 1) return this.data.amount;
            return `<span class="quantity-badge">${this.data.amount}</span>`;
        }
    }

    function createPrintContent(apiData) {
        const products = apiData.order?.products || [];
        const items = apiData.order?.items || [];
        const sale = apiData.sale || {};

        const headerHtml = `
            <div class="print-controls">
                <button id="exportExcelBtn">ייצא לאקסל</button>
                <button id="printBtnNew" class="print-btn">הדפס</button>
                <button id="showAdminBtn" class="admin-btn">הצג פרטי מנהל</button>
            </div>
            <div class="order-header">
                <h2>פרטי הזמנה - ${sale.name || ''}</h2>
                <div class="order-header-grid">
                    <div class="order-header-item"><strong>לקוח:</strong> <span>${apiData.order?.customer?.name || 'לא צוין'}</span></div>
                    <div class="order-header-item"><strong>תאריך איסוף:</strong> <span>${new Date(sale.pickup_date).toLocaleDateString('he-IL') || 'לא צוין'}</span></div>
                    <div class="order-header-item"><strong>שעת איסוף:</strong> <span>${sale.pickup_time_start ? `${sale.pickup_time_start} - ${sale.pickup_time_end}` : 'לא צוין'}</span></div>
                    <div class="order-header-item"><strong>מיקום:</strong> <span>${sale.site?.name || 'לא צוין'}</span></div>
                    <div class="order-header-item"><strong>סה"כ:</strong> <span>${apiData.order?.total ? `₪${apiData.order.total}` : 'לא צוין'}</span></div>
                </div>
            </div>
        `;

        let tableHtml = `
          <table>
            <thead>
                <tr>
                  <th style="width: 5%">#</th>
                  <th style="width: 10%">תמונה</th>
                  <th style="width: 35%">שם המוצר</th>
                  <th style="width: 15%">מחיר</th>
                  <th style="width: 10%">כמות</th>
                  <th style="width: 15%">סה"כ</th>
                  <th style="width: 10%">הערות</th>
                </tr>
            </thead>
            <tbody>
        `;

        const groupedProducts = {};
        products.forEach(product => {
            const item = items.find(i => i.item_salesID === product.item_salesID);
            if (!item) return;
            const mainCategory = item.item_sale?.product?.category?.main_category?.name || 'ללא קטגוריה ראשית';
            const subCategory = item.item_sale?.product?.category?.name || 'ללא קטגוריה משנית';
            if (!groupedProducts[mainCategory]) groupedProducts[mainCategory] = {};
            if (!groupedProducts[mainCategory][subCategory]) groupedProducts[mainCategory][subCategory] = [];
            groupedProducts[mainCategory][subCategory].push(new Product(product, item));
        });

        let index = 1;
        const sortedMainCategories = Object.keys(groupedProducts).sort();
        sortedMainCategories.forEach(mainCategory => {
            tableHtml += `<tr class="category-header"><th colspan="7">${mainCategory}</th></tr>`;
            const sortedSubCategories = Object.keys(groupedProducts[mainCategory]).sort();
            sortedSubCategories.forEach(subCategory => {
                tableHtml += `<tr class="subcategory-header"><th colspan="7">${subCategory}</th></tr>`;
                groupedProducts[mainCategory][subCategory].forEach(product => {
                    const imageUrl = product.data.featured_image ? `https://images.mishnatyosef.org/images/items/${product.data.featured_image}` : '';
                    const isCut = product.isCut();
                    tableHtml += `
                        <tr class="product-row ${isCut ? 'cut-row' : ''}">
                          <td>${index++}</td>
                          <td>${imageUrl ? `<img src="${imageUrl}" alt="${product.data.full_name}" class="product-image">` : ''}</td>
                          <td>${product.data.full_name}</td>
                          <td>₪${product.data.price}</td>
                          <td>${product.getQuantityBadge()}</td>
                          <td>₪${(product.data.price * product.data.amount).toFixed(2)}</td>
                          <td>
                            <div>${product.getPricePerUnit()}</div>
                            ${product.item.item_sale?.product?.note ? `<div style="font-size: 0.875rem; margin-top: 4px;">${product.item.item_sale.product.note}</div>` : ''}
                            ${product.getStationManagerButton()}
                            ${isCut ? `<div style="color: #dc3545; font-weight: 500; margin-top: 4px;">מוצר מקוצץ</div>` : ''}
                          </td>
                        </tr>
                    `;
                });
            });
        });

        tableHtml += '</tbody></table>';
        return headerHtml + tableHtml;
    }

    function showAdminInfo(apiData) {
        const siteAdmin = apiData.sale?.site;
        if (siteAdmin && siteAdmin.admin_name) {
            const adminInfoBody = `
                <p><strong>שם:</strong> ${siteAdmin.admin_name}</p>
                <p><strong>טלפון:</strong> ${siteAdmin.admin_phone || 'לא צוין'}</p>
                <p><strong>אימייל:</strong> ${siteAdmin.admin_email || 'לא צוין'}</p>
            `;
            showModal('פרטי מנהל אתר', adminInfoBody);
        } else {
            showModal('פרטי מנהל אתר', 'לא נמצאו פרטי מנהל אתר בהזמנה זו.');
        }
    }

    function exportToExcel(apiData) {
        const products = apiData.order?.products || [];
        const items = apiData.order?.items || [];
        const dataForExcel = [];
        products.forEach(p => {
            const item = items.find(i => i.item_salesID === p.item_salesID);
            if (!item) return;
            dataForExcel.push({
                'קטגוריה ראשית': item.item_sale?.product?.category?.main_category?.name || '',
                'קטגורית משנה': item.item_sale?.product?.category?.name || '',
                'שם המוצר': p.full_name,
                'מחיר': p.price,
                'כמות': p.amount,
                'סה"כ': (p.price * p.amount).toFixed(2),
                'הערות מוצר': item.item_sale?.product?.note || '',
                'אחראי תחנה': item.item_sale?.product?.station_manager || '',
                'מוצר מקוצץ': (p.src_amount && p.src_amount > 0) ? 'כן' : 'לא'
            });
        });

        const worksheet = XLSX.utils.json_to_sheet(dataForExcel);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Order");

        const cols = Object.keys(dataForExcel[0]);
        const colWidths = cols.map(col => ({ wch: Math.max(...dataForExcel.map(row => row[col]?.toString().length || 0), col.length) }));
        worksheet['!cols'] = colWidths;

        XLSX.writeFile(workbook, `MishnatYosef_Order_${apiData.order?.id || 'export'}.xlsx`);
    }
});
