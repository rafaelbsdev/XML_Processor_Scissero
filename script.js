const inputXML = document.querySelector(".inputXML");
const message = document.querySelector(".message");
const dataTable = document.querySelector(".dataTable");
const btnPreview = document.querySelector(".btnPreviewData");
const btnExport = document.querySelector(".btnExportExcel");
const productTypeFilter = document.getElementById('productTypeFilter');

btnPreview.addEventListener("click", previewExtractedXML);
btnExport.addEventListener("click", exportExcel);

productTypeFilter.addEventListener('change', (event) => {
    const selectedType = event.target.value;
    const tableRows = document.querySelectorAll('.dataTable tbody tr');
    tableRows.forEach(row => {
        const rowProductType = row.dataset.productType;
        if (selectedType === 'all' || rowProductType === selectedType) {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
});

const setDate = (strDate) => {
    if (!strDate) return "";
    const date = new Date(strDate + 'T00:00:00Z');
    if (isNaN(date.getTime())) return "";

    const day = date.getUTCDate().toString().padStart(2, '0');
    const month = date.toLocaleString('en-US', { month: 'short', timeZone: 'UTC' });
    const year = date.getUTCFullYear().toString().slice(-2);
    
    return `${day}-${month}-${year}`;
};

const uniqueArray = (array, objKey) => {
    const uniqValues = new Set();
    array.forEach((item) => {
        if (item && typeof item === 'object' && item[objKey]) {
            uniqValues.add(item[objKey]);
        }
    });
    return [...uniqValues];
};

const setBuffer = (value) => 100 - parseFloat(value);

const findFirstContent = (node, selectors) => {
    if (!node) return "";
    for (const selector of selectors) {
        const element = node.querySelector(selector);
        if (element && element.textContent) {
            return element.textContent.trim();
        }
    }
    return "";
};

const getUnderlying = (assetNode, key) => {
    const assets = assetNode.querySelectorAll("assets");
    if (typeof key === "number") {
        return findFirstContent(assets[key], ["bloombergTickerSuffix"]);
    } else {
        if (assets.length === 0) return "N/A";
        const assetType = findFirstContent(assets[0], ["assetType"]).replace("Exchange_Traded_Fund", "ETF");
        const basketType = findFirstContent(assetNode, ["basketType"]) || "Multiple";
        return assets.length === 1 ? `Single ${assetType}` : `${basketType} ${assetType}`;
    }
};
const getProducts = (productNode, type) => {
    if (type === "tenor") {
        const tenorMonths = findFirstContent(productNode, ["tenor > months"]);
        if (!tenorMonths) return "";
        const months = parseInt(tenorMonths, 10);
        return months > 12 && months % 12 === 0 ? `${months / 12}Y` : `${months}M`;
    }
    return findFirstContent(productNode, ["bufferedReturnEnhancedNote > productType", "reverseConvertible > description", "productName"]);
};
const upsideLeverage = (productNode) => {
    const leverage = findFirstContent(productNode, ["bufferedReturnEnhancedNote > upsideLeverage"]);
    return (leverage && `${leverage}X`) || "N/A";
};
const upsideCap = (productNode) => {
    const cap = findFirstContent(productNode, ["bufferedReturnEnhancedNote > upsideCap"]);
    return (cap && `${cap}%`) || "N/A";
};
const getCoupon = (productNode, type) => {
    const couponSchedule = productNode.querySelector("reverseConvertible > couponSchedule");
    if (!couponSchedule) return "N/A";
    switch (type) {
        case "frequency": return findFirstContent(couponSchedule, ["frequency"]) || "N/A";
        case "level":
            const level = findFirstContent(couponSchedule, ["coupons > contingentLevelLegal > level", "contigentLevelLegal > level"]);
            return (level && `${level}%`) || "N/A";
        case "memory":
            const hasMemory = findFirstContent(couponSchedule, ["hasMemory"]);
            return hasMemory ? (hasMemory === "false" ? "N" : "Y") : "N/A";
        default: return "N/A";
    }
};
const getEarlyStrike = (xmlNode) => {
    const strikeRaw = findFirstContent(xmlNode, ["securitized > issuance > prospectusStartDate", "strikeDate > date"]);
    const pricingRaw = findFirstContent(xmlNode, ["securitized > issuance > clientOrderTradeDate"]);
    return strikeRaw && pricingRaw && strikeRaw !== pricingRaw ? "Y" : "N";
};
const detectClient = (xmlNode) => {
    const cp = findFirstContent(xmlNode, ["counterparty > name"]);
    const dealer = findFirstContent(xmlNode, ["dealer > name"]);
    const blob = (cp + " " + dealer).toLowerCase();
    if (blob.includes("jpm")) return "JPM PB";
    if (blob.includes("goldman")) return "GS";
    if (blob.includes("bauble") || blob.includes("ubs")) return "UBS";
    return "3P";
};
const detectXmlType = (xmlNode) => {
    const docType = findFirstContent(xmlNode, ["documentType"]).toUpperCase();
    return {
        termsheet: docType.includes("TERMSHEET") ? "Y" : "N",
        finalPS: docType.includes("PRICING_SUPPLEMENT") ? "Y" : "N",
        factSheet: docType.includes("FACT_SHEET") ? "Y" : "N",
    };
};
const getDetails = (productNode, type, tradableFormNode) => {
    const productType = getProducts(productNode);
    if (productType === "BREN" || productType === "REN") {
        switch (type) {
            case "strikelevel": return "Buffer";
            case "bufferlevel":
                const buffer = findFirstContent(productNode, ["bufferedReturnEnhancedNote > buffer"]);
                return buffer ? `${100 - parseFloat(buffer)}%` : "";
            case "frequency": return "At Maturity";
            case "noncall": return "N/A";
        }
    }
    const phoenixType = productNode.querySelector("reverseConvertible > issuerCallable") ? "issuerCallable" : "autocallSchedule";
    switch (type) {
        case "strikelevel":
            const strikelevel = findFirstContent(productNode, ["reverseConvertible > strike > level"]);
            return parseFloat(strikelevel) < 100 ? "Buffer" : "KI Barrier";
        case "bufferlevel":
             const kiBarrier = findFirstContent(productNode, ["knockInBarrier > barrierSchedule > barrierLevel > level"]);
             return kiBarrier ? `${kiBarrier}%` : "";
        case "frequency": return findFirstContent(productNode, [`reverseConvertible > ${phoenixType} > barrierSchedule > frequency`]);
        case "noncall":
            const issueDateStr = findFirstContent(tradableFormNode, ["securitized > issuance > issueDate"]);
            const firstCallDateStr = findFirstContent(productNode, [`reverseConvertible > ${phoenixType} > barrierSchedule > firstDate`]);
            if (!issueDateStr || !firstCallDateStr) return "";
            const issueDate = new Date(issueDateStr);
            const firstCallDate = new Date(firstCallDateStr);
            if (isNaN(issueDate.getTime()) || isNaN(firstCallDate.getTime())) return "";
            const months = (firstCallDate.getFullYear() - issueDate.getFullYear()) * 12 + (firstCallDate.getMonth() - issueDate.getMonth());
            return months > 0 ? `${months}M` : "";
        default:
            const barrierLevel = parseFloat(findFirstContent(productNode, ["reverseConvertible > couponSchedule > contigentLevelLegal > level"]));
            const kiLevel = parseFloat(findFirstContent(productNode, ["knockInBarrier > barrierSchedule > barrierLevel > level"]));
            if (isNaN(barrierLevel) || isNaN(kiLevel)) return "N/A";
            const comparison = barrierLevel > kiLevel ? ">" : barrierLevel < kiLevel ? "<" : "=";
            return `Interest Barrier ${comparison} KI Barrier`;
    }
};
const findIdentifier = (tradableFormNode, type) => {
    const identifiers = tradableFormNode.querySelectorAll('identifiers');
    for (const idNode of identifiers) {
        const typeNode = idNode.querySelector('type');
        if (typeNode && typeNode.textContent.trim().toUpperCase() === type) {
            const codeNode = idNode.querySelector('code');
            return codeNode ? codeNode.textContent.trim() : "";
        }
    }
    return "";
};

async function previewExtractedXML() {
    const files = Array.from(inputXML.files).filter((file) => file.name.endsWith(".xml"));
    if (files.length === 0) {
        message.innerHTML = "Please select at least one XML file.";
        return;
    }

    message.innerHTML = `Processing ${files.length} file(s)...`;
    dataTable.innerHTML = "";
    btnExport.style.display = "none";
    document.querySelector('.filter-container').style.display = 'none';
    btnPreview.disabled = true;

    try {
        const productMap = new Map();

        for (const file of files) {
            const parser = new DOMParser();
            const content = await file.text();
            const xmlNode = parser.parseFromString(content, "application/xml");
            const parseError = xmlNode.querySelector('parsererror');
            if (parseError) {
                throw new Error(`Error: The file ${file.name} is corrupt or not a valid XML.`);
            }
            const product = xmlNode.querySelector("product");
            const tradableForm = xmlNode.querySelector("tradableForm");
            const asset = xmlNode.querySelector("asset");
            const product_type = getProducts(product);
            const { termsheet, finalPS, factSheet } = detectXmlType(xmlNode);
            const cusip = findIdentifier(tradableForm, 'CUSIP');
            const isin = findIdentifier(tradableForm, 'ISIN');
            const identifier = cusip || isin;

            if (!identifier) continue;

            if (productMap.has(identifier)) { // Se existir ele da merge nos id
                const existingItem = productMap.get(identifier);
                existingItem.termSheet = (existingItem.termSheet === 'Y' || termsheet === 'Y') ? 'Y' : 'N';
                existingItem.finalPS   = (existingItem.finalPS === 'Y' || finalPS === 'Y') ? 'Y' : 'N';
                existingItem.factSheet = (existingItem.factSheet === 'Y' || factSheet === 'Y') ? 'Y' : 'N';
            } else {
                const item = {
                    prodCusip: cusip,
                    prodIsin: isin,
                    underlyingAssetType: getUnderlying(asset),
                    assets: Array.from(asset.querySelectorAll("assets")).map(node => findFirstContent(node, ["bloombergTickerSuffix"])),
                    productType: product_type,
                    productClient: detectClient(xmlNode),
                    productTenor: getProducts(product, "tenor"),
                    couponFrequency: getCoupon(product, "frequency"),
                    couponBarrierLevel: getCoupon(product, "level"),
                    couponMemory: getCoupon(product, "memory"),
                    upsideCap: (product_type === "BREN" || product_type === "REN") ? upsideCap(product) : "N/A",
                    upsideLeverage: (product_type === "BREN" || product_type === "REN") ? upsideLeverage(product) : "N/A",
                    detailBufferKIBarrier: getDetails(product, "strikelevel"),
                    detailBufferBarrierLevel: getDetails(product, "bufferlevel"),
                    detailFrequency: getDetails(product, "frequency", tradableForm),
                    detailNonCallPerid: getDetails(product, "noncall", tradableForm),
                    detailInterestBarrierTriggerValue: getDetails(product),
                    dateBookingStrikeDate: setDate(findFirstContent(xmlNode, ["securitized > issuance > prospectusStartDate", "strikeDate > date"])),
                    dateBookingPricingDate: setDate(findFirstContent(tradableForm, ["securitized > issuance > clientOrderTradeDate"])),
                    maturityDate: setDate(findFirstContent(product, ["redemptionDate", "settlementDate"])),
                    valuationDate: setDate(findFirstContent(product, ["finalObservation"])),
                    earlyStrike: getEarlyStrike(xmlNode),
                    termSheet: termsheet,
                    finalPS: finalPS,
                    factSheet: factSheet,
                };
                productMap.set(identifier, item);
            }
        }

        const consolidatedData = Array.from(productMap.values());

        const maxAssets = consolidatedData.reduce((max, item) => Math.max(max, item.assets.length), 0);
        renderTable(consolidatedData, maxAssets);
        message.innerHTML = "";
        btnExport.style.display = "inline-block";
    } catch (error) {
        console.error("An unexpected error occurred:", error);
        message.innerHTML = error.message || "An unexpected error occurred. Check the console for details.";
    } finally {
        btnPreview.disabled = false;
    }
}

function renderTable(data, maxAssets) {
    if (data.length === 0) {
        message.innerHTML = `No data found in the selected XML files.`;
        return;
    }
    const filterContainer = document.querySelector('.filter-container');
    productTypeFilter.innerHTML = ''; 
    const productTypes = uniqueArray(data, 'productType');
    const allOption = document.createElement('option');
    allOption.value = 'all';
    allOption.textContent = 'All Types';
    productTypeFilter.appendChild(allOption);
    productTypes.forEach(type => {
        if (type) {
            const option = document.createElement('option');
            option.value = type;
            option.textContent = type;
            productTypeFilter.appendChild(option);
        }
    });
    filterContainer.style.display = 'inline-block';
    
    let assetHeaders;
    if (maxAssets === 1) {
        assetHeaders = ["Asset"];
    } else {
        assetHeaders = Array.from({ length: maxAssets }, (_, i) => `Asset ${i + 1}`);
    }
    
    const headers = [
        { title: "Prod CUSIP" },
        { title: "ISIN" },
        { title: "Underlying", children: ["Asset Type", ...assetHeaders] },
        { title: "Product Details", children: ["Product Type", "Client", "Tenor"] },
        { title: "Coupons", children: ["Frequency", "Barrier Level", "Memory"] },
        { title: "Details", children: ["Upside Cap", "Upside Leverage", "Buffer Threshold / KI Barrier", "Barrier/Buffer Level", "Frequency", "Non-call period", "Interest Barrier vs KI"] },
        { title: "DATES IN BOOKINGS", children: ["Strike", "Pricing", "Maturity", "Valuation", "Early Strike"] },
        { title: "Doc Type", children: ["Term Sheet", "Final PS", "Fact Sheet"] }
    ];

    let htmlTable = "<table><thead><tr class='first-tr'>";
    headers.forEach(({ title, children }) => {
        const span = children ? `colspan="${children.length}" class="category"` : `rowspan="2"`;
        htmlTable += `<th ${span}>${title.toUpperCase()}</th>`;
    });
    htmlTable += "</tr><tr class='second-tr'>";
    headers.forEach(({ children }) => {
        if (children) children.forEach(child => (htmlTable += `<th>${child}</th>`));
    });
    htmlTable += "</tr></thead><tbody>";
    data.forEach(row => {
        htmlTable += `<tr data-product-type="${row.productType || ''}">`;
        htmlTable += `<td>${row.prodCusip || ""}</td>`;
        htmlTable += `<td>${row.prodIsin || ""}</td>`;
        htmlTable += `<td>${row.underlyingAssetType || ""}</td>`;
        for (let i = 0; i < maxAssets; i++) {
            htmlTable += `<td>${row.assets[i] || ""}</td>`;
        }
        htmlTable += `<td>${row.productType || ""}</td>`;
        htmlTable += `<td>${row.productClient || ""}</td>`;
        htmlTable += `<td>${row.productTenor || ""}</td>`;
        htmlTable += `<td>${row.couponFrequency || ""}</td>`;
        htmlTable += `<td>${row.couponBarrierLevel || ""}</td>`;
        htmlTable += `<td>${row.couponMemory || ""}</td>`;
        htmlTable += `<td>${row.upsideCap || ""}</td>`;
        htmlTable += `<td>${row.upsideLeverage || ""}</td>`;
        htmlTable += `<td>${row.detailBufferKIBarrier || ""}</td>`;
        htmlTable += `<td>${row.detailBufferBarrierLevel || ""}</td>`;
        htmlTable += `<td>${row.detailFrequency || ""}</td>`;
        htmlTable += `<td>${row.detailNonCallPerid || ""}</td>`;
        htmlTable += `<td>${row.detailInterestBarrierTriggerValue || ""}</td>`;
        htmlTable += `<td>${row.dateBookingStrikeDate || ""}</td>`;
        htmlTable += `<td>${row.dateBookingPricingDate || ""}</td>`;
        htmlTable += `<td>${row.maturityDate || ""}</td>`;
        htmlTable += `<td>${row.valuationDate || ""}</td>`;
        htmlTable += `<td>${row.earlyStrike || ""}</td>`;
        htmlTable += `<td>${row.termSheet || ""}</td>`;
        htmlTable += `<td>${row.finalPS || ""}</td>`;
        htmlTable += `<td>${row.factSheet || ""}</td>`;
        htmlTable += "</tr>";
    });
    htmlTable += "</tbody></table>";
    dataTable.innerHTML = htmlTable;
}

function exportExcel() {
    const originalTable = dataTable.querySelector("table");
    if (!originalTable) {
        message.innerHTML = "There is no data to export.";
        return;
    }
    const tempTable = document.createElement('table');
    const thead = originalTable.querySelector('thead');
    if (thead) {
        tempTable.appendChild(thead.cloneNode(true));
    }
    const visibleRows = originalTable.querySelectorAll('tbody tr:not([style*="display: none"])');
    const tbody = document.createElement('tbody');
    visibleRows.forEach(row => {
        tbody.appendChild(row.cloneNode(true));
    });
    tempTable.appendChild(tbody);
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.table_to_sheet(tempTable);
    const colWidths = [];
    for (const cellAddress in worksheet) {
        if (cellAddress[0] === '!') continue;
        const cell = XLSX.utils.decode_cell(cellAddress);
        worksheet[cellAddress].s = {
            font: { name: "Arial", sz: 10 },
            alignment: { vertical: "top", horizontal: "left" },
            border: { top: { style: "thin", color: "000000" }, right: { style: "thin", color: "000000" }, bottom: { style: "thin", color: "000000" }, left: { style: "thin", color: "000000" } }
        };
        if (cell.r < 2) {
            worksheet[cellAddress].s.font.bold = true;
            worksheet[cellAddress].s.alignment.horizontal = "center";
            worksheet[cellAddress].s.fill = { fgColor: { rgb: "DDEBF7" } };
        }
        const value = worksheet[cellAddress].v;
        const width = value ? value.toString().length + 2 : 12;
        if (!colWidths[cell.c] || width > colWidths[cell.c].wch) {
            colWidths[cell.c] = { wch: width };
        }
    }
    worksheet['!cols'] = colWidths;
    XLSX.utils.book_append_sheet(workbook, worksheet, "Scenario List");
    
    XLSX.writeFile(workbook, "ExtractedXML_Data.xlsx");
}