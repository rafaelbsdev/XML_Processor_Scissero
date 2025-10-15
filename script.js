const inputXML = document.querySelector(".inputXML");
const message = document.querySelector(".message");
const dataTable = document.querySelector(".dataTable");
const btnPreview = document.querySelector(".btnPreviewData");
const btnExport = document.querySelector(".btnExportExcel");
const productTypeFilter = document.getElementById('productTypeFilter');

let consolidatedData = [];
let headerStructure = [];
let maxAssetsForExport = 0;

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

const sanitizeSheetName = (name) => {
    return name.replace(/[\\/*?\[\]:]/g, "").substring(0, 31);
};

const formatAsPercentage = (value) => {
    if (!value || value === "N/A" || value === "") {
        return value;
    }
    const num = parseFloat(value);
    if (isNaN(num)) {
        return value;
    }

    return `${Number(num.toFixed(4))}%`;
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
    return cap ? formatAsPercentage(cap) : "N/A";
};

const getCoupon = (productNode, type) => {
    const couponSchedule = productNode.querySelector("reverseConvertible > couponSchedule");
    if (!couponSchedule) return "N/A";
    switch (type) {
        case "frequency": return findFirstContent(couponSchedule, ["frequency"]) || "N/A";
        case "level":
            const level = findFirstContent(couponSchedule, ["coupons > contingentLevelLegal > level", "contigentLevelLegal > level"]);
            return level ? formatAsPercentage(level) : "N/A";
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
                return buffer ? formatAsPercentage(100 - parseFloat(buffer)) : "";
            case "frequency": return "At Maturity";
            case "noncall": return "N/A";
        }
    }
    const phoenixType = productNode.querySelector("reverseConvertible > issuerCallable") ? "issuerCallable" : "autocallSchedule";
    switch (type) {
        case "strikelevel":
            const strikelevelValue = findFirstContent(productNode, ["reverseConvertible > strike > level"]);
            return parseFloat(strikelevelValue) < 100 ? "Buffer" : "KI Barrier";

        case "bufferlevel":
            const strikelevel = findFirstContent(productNode, ["reverseConvertible > strike > level"]);
            const isBuffer = parseFloat(strikelevel) < 100;

            if (isBuffer) {
                const bufferLevel = findFirstContent(productNode, ["reverseConvertible > buffer > level"]);
                return formatAsPercentage(bufferLevel);
            } else {
                const kiBarrier = findFirstContent(productNode, ["knockInBarrier > barrierSchedule > barrierLevel > level"]);
                return formatAsPercentage(kiBarrier);
            }
        
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
            const identifier = cusip;

            if (!identifier) continue;
            
            const cappedValue = findFirstContent(product, ['bufferedReturnEnhancedNote > capped', 'capped']);

            if (productMap.has(identifier)) {
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
                    upsideCap: upsideCap(product),
                    upsideLeverage: upsideLeverage(product),
                    detailCappedUncapped: (product_type === 'BREN' || product_type === 'REN')
                        ? (cappedValue === 'true' ? 'Capped' : (cappedValue === 'false' ? 'Uncapped' : ''))
                        : '',
                    detailBufferKIBarrier: getDetails(product, "strikelevel", tradableForm),
                    detailBufferBarrierLevel: getDetails(product, "bufferlevel", tradableForm),
                    detailFrequency: getDetails(product, "frequency", tradableForm),
                    detailNonCallPerid: getDetails(product, "noncall", tradableForm),
                    detailInterestBarrierTriggerValue: getDetails(product, "default", tradableForm),
                    dateBookingStrikeDate: setDate(findFirstContent(xmlNode, ["securitized > issuance > prospectusStartDate", "strikeDate > date"])),
                    dateBookingPricingDate: setDate(findFirstContent(tradableForm, ["securitized > issuance > clientOrderTradeDate"])),
                    maturityDate: setDate(findFirstContent(product, ["maturity > date", "redemptionDate", "settlementDate"])),
                    valuationDate: setDate(findFirstContent(product, ["finalObsDate > date", "finalObservation"])),
                    earlyStrike: getEarlyStrike(xmlNode),
                    termSheet: termsheet,
                    finalPS: finalPS,
                    factSheet: factSheet,
                };
                productMap.set(identifier, item);
            }
        }

        consolidatedData = Array.from(productMap.values());
        maxAssetsForExport = consolidatedData.reduce((max, item) => Math.max(max, item.assets.length), 0);
        renderTable(consolidatedData, maxAssetsForExport);
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
    
    const hasBrenRenProducts = data.some(row => row.productType === 'BREN' || row.productType === 'REN');
    
    const detailsChildren = ["Upside Cap", "Upside Leverage"];
    if (hasBrenRenProducts) {
        detailsChildren.push("Capped / Uncapped");
    }
    detailsChildren.push("Buffer / Barrier", "Barrier/Buffer Level", "Interest Barrier vs KI");

    headerStructure = [
        { title: "Prod CUSIP" }, { title: "ISIN" },
        { title: "Underlying", children: ["Asset Type", ...assetHeaders] },
        { title: "Product Details", children: ["Product Type", "Client", "Tenor"] },
        { title: "Coupons", children: ["Frequency", "Barrier Level", "Memory"] },
        { title: "CALL", children: ["Frequency", "Non-call period"] },
        { title: "Details", children: detailsChildren },
        { title: "DATES IN BOOKINGS", children: ["Strike", "Pricing", "Maturity", "Valuation", "Early Strike"] },
        { title: "Doc Type", children: ["Term Sheet", "Final PS", "Fact Sheet"] }
    ];

    let htmlTable = "<table><thead><tr class='first-tr'>";
    headerStructure.forEach(({ title, children }) => {
        const span = children ? `colspan="${children.length}" class="category"` : `rowspan="2"`;
        htmlTable += `<th ${span}>${title.toUpperCase()}</th>`;
    });
    htmlTable += "</tr><tr class='second-tr'>";
    headerStructure.forEach(({ children }) => {
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
        htmlTable += `<td>${row.detailFrequency || ""}</td>`;
        htmlTable += `<td>${row.detailNonCallPerid || ""}</td>`;
        htmlTable += `<td>${row.upsideCap || ""}</td>`;
        htmlTable += `<td>${row.upsideLeverage || ""}</td>`;
        if (hasBrenRenProducts) {
            htmlTable += `<td>${row.detailCappedUncapped || ""}</td>`;
        }
        htmlTable += `<td>${row.detailBufferKIBarrier || ""}</td>`;
        htmlTable += `<td>${row.detailBufferBarrierLevel || ""}</td>`;
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
    const selectedType = productTypeFilter.value;
    let filteredData = consolidatedData;

    if (selectedType !== 'all') {
        filteredData = consolidatedData.filter(row => row.productType === selectedType);
    }

    if (filteredData.length === 0) {
        message.innerHTML = "No data to export for the selected filter.";
        return;
    }

    const hasBrenRenProductsInExport = filteredData.some(row => row.productType === 'BREN' || row.productType === 'REN');

    const groupedData = filteredData.reduce((acc, row) => {
        const key = row.productType || 'Uncategorized';
        if (!acc[key]) acc[key] = [];
        acc[key].push(row);
        return acc;
    }, {});

    const workbook = XLSX.utils.book_new();

    for (const productType in groupedData) {
        const sheetData = groupedData[productType];
        const sheetName = sanitizeSheetName(productType);
        
        const isBrenRenSheet = productType === 'BREN' || productType === 'REN';
        
        let localHeaderStructure = JSON.parse(JSON.stringify(headerStructure));
        
        if (hasBrenRenProductsInExport && !isBrenRenSheet) {
            const details = localHeaderStructure.find(h => h.title === 'Details');
            if (details) {
                details.children = details.children.filter(c => c !== 'Capped / Uncapped');
            }
        }
        
        const headerRow1 = [];
        const headerRow2 = [];
        localHeaderStructure.forEach(h => {
            headerRow1.push(h.title.toUpperCase());
            if (h.children) {
                headerRow2.push(...h.children);
                for (let i = 1; i < h.children.length; i++) {
                    headerRow1.push(null);
                }
            } else {
                headerRow2.push(null);
            }
        });

        const dataAoA = sheetData.map(row => {
            const rowAsArray = [];
            rowAsArray.push(row.prodCusip, row.prodIsin, row.underlyingAssetType);
            for (let i = 0; i < maxAssetsForExport; i++) {
                rowAsArray.push(row.assets[i] || "");
            }
            rowAsArray.push(
                row.productType, row.productClient, row.productTenor, row.couponFrequency,
                row.couponBarrierLevel, row.couponMemory, row.detailFrequency, row.detailNonCallPerid,
                row.upsideCap, row.upsideLeverage
            );
            
            if (isBrenRenSheet) {
                rowAsArray.push(row.detailCappedUncapped);
            }
            
            rowAsArray.push(
                row.detailBufferKIBarrier, row.detailBufferBarrierLevel, row.detailInterestBarrierTriggerValue,
                row.dateBookingStrikeDate, row.dateBookingPricingDate, row.maturityDate, 
                row.valuationDate, row.earlyStrike, row.termSheet, row.finalPS, row.factSheet
            );
            return rowAsArray;
        });

        const worksheet = XLSX.utils.aoa_to_sheet([headerRow1, headerRow2, ...dataAoA]);

        const merges = [];
        let colIndex = 0;
        localHeaderStructure.forEach(header => {
            if (header.children) {
                if (header.children.length > 1) {
                    merges.push({ s: { r: 0, c: colIndex }, e: { r: 0, c: colIndex + header.children.length - 1 } });
                }
                colIndex += header.children.length;
            } else {
                merges.push({ s: { r: 0, c: colIndex }, e: { r: 1, c: colIndex } });
                colIndex++;
            }
        });
        worksheet['!merges'] = merges;
        
        const colWidths = [];
        Object.keys(worksheet).forEach(cellAddress => {
            if (cellAddress[0] === '!') return;
            const cell = XLSX.utils.decode_cell(cellAddress);
            worksheet[cellAddress].s = {
                font: { name: "Arial", sz: 10 },
                alignment: { vertical: "center", horizontal: "left" },
                border: { top: { style: "thin" }, right: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" } }
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
        });
        worksheet['!cols'] = colWidths;

        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    }
    
    XLSX.writeFile(workbook, "ExtractedXML_Data.xlsx");
}