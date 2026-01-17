/**
 * Excel Sync Module
 * Handles API calls, modal management, file upload/download, and data synchronization
 */

(function() {
    'use strict';

    // =============================================================================
    // State
    // =============================================================================
    
    const state = {
        hasUnsavedChanges: false,
        isSaving: false,
        currentEditNodeId: null,
        currentDeleteNodeId: null,
        departments: new Set()
    };

    // =============================================================================
    // DOM Elements
    // =============================================================================
    
    let elements = {};

    function initElements() {
        elements = {
            // File operations
            btnUpload: document.getElementById('btnUpload'),
            btnDownloadTemplate: document.getElementById('btnDownloadTemplate'),
            fileInput: document.getElementById('fileInput'),
            btnSave: document.getElementById('btnSave'),
            saveStatus: document.getElementById('saveStatus'),
            
            // Add root
            btnAddRoot: document.getElementById('btnAddRoot'),
            
            // Export
            btnOpenExport: document.getElementById('btnOpenExport'),
            exportModal: document.getElementById('exportModal'),
            exportBackground: document.getElementById('exportBackground'),
            exportBgColor: document.getElementById('exportBgColor'),
            customBgGroup: document.getElementById('customBgGroup'),
            exportPadding: document.getElementById('exportPadding'),
            jpegQuality: document.getElementById('jpegQuality'),
            jpegQualityValue: document.getElementById('jpegQualityValue'),
            jpegQualityGroup: document.getElementById('jpegQualityGroup'),
            exportPreviewBox: document.getElementById('exportPreviewBox'),
            exportDimensions: document.getElementById('exportDimensions'),
            exportEstSize: document.getElementById('exportEstSize'),
            exportResDimensions: document.getElementById('exportResDimensions'),
            btnDoExport: document.getElementById('btnDoExport'),
            
            // Edit Modal
            editModal: document.getElementById('editModal'),
            editModalTitle: document.getElementById('editModalTitle'),
            editForm: document.getElementById('editForm'),
            editNodeId: document.getElementById('editNodeId'),
            editName: document.getElementById('editName'),
            editTitle: document.getElementById('editTitle'),
            editDepartment: document.getElementById('editDepartment'),
            editManager: document.getElementById('editManager'),
            editColor: document.getElementById('editColor'),
            editColorHex: document.getElementById('editColorHex'),
            editAvatarUrl: document.getElementById('editAvatarUrl'),
            editWhatsapp: document.getElementById('editWhatsapp'),
            departmentSuggestions: document.getElementById('departmentSuggestions'),
            
            // Add Modal
            addModal: document.getElementById('addModal'),
            addForm: document.getElementById('addForm'),
            addManagerId: document.getElementById('addManagerId'),
            addName: document.getElementById('addName'),
            addTitle: document.getElementById('addTitle'),
            addDepartment: document.getElementById('addDepartment'),
            addManager: document.getElementById('addManager'),
            addColor: document.getElementById('addColor'),
            addColorHex: document.getElementById('addColorHex'),
            addWhatsapp: document.getElementById('addWhatsapp'),
            
            // Delete Modal
            deleteModal: document.getElementById('deleteModal'),
            deleteNodeId: document.getElementById('deleteNodeId'),
            deleteNodeName: document.getElementById('deleteNodeName'),
            deleteReassignSection: document.getElementById('deleteReassignSection'),
            deleteReassignTo: document.getElementById('deleteReassignTo'),
            btnConfirmDelete: document.getElementById('btnConfirmDelete'),
            
            // Toast
            toastContainer: document.getElementById('toastContainer')
        };
    }

    // =============================================================================
    // API Calls
    // =============================================================================
    
    const API = {
        /**
         * Fetch organization data from server
         */
        async getOrgData() {
            const response = await fetch('/get_org_data');
            const result = await response.json();
            
            if (!result.success) {
                throw new Error(result.error || 'Failed to fetch organization data');
            }
            
            return result;
        },
        
        /**
         * Upload Excel file
         */
        async uploadExcel(file) {
            const formData = new FormData();
            formData.append('file', file);
            
            const response = await fetch('/upload_excel', {
                method: 'POST',
                body: formData
            });
            
            const result = await response.json();
            
            if (!result.success) {
                throw new Error(result.error || 'Failed to upload file');
            }
            
            return result;
        },
        
        /**
         * Update a node
         */
        async updateNode(nodeData) {
            const response = await fetch('/update_node', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(nodeData)
            });
            
            const result = await response.json();
            
            if (!result.success) {
                throw new Error(result.error || 'Failed to update node');
            }
            
            return result;
        },
        
        /**
         * Add a new node
         */
        async addNode(nodeData) {
            const response = await fetch('/add_node', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(nodeData)
            });
            
            const result = await response.json();
            
            if (!result.success) {
                throw new Error(result.error || 'Failed to add node');
            }
            
            return result;
        },
        
        /**
         * Delete a node
         */
        async deleteNode(nodeId, reassignTo = null) {
            const response = await fetch('/delete_node', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: nodeId, reassign_to: reassignTo })
            });
            
            const result = await response.json();
            
            if (!result.success) {
                throw new Error(result.error || 'Failed to delete node');
            }
            
            return result;
        },
        
        /**
         * Save to Excel file
         */
        async saveExcel(filename = null) {
            const response = await fetch('/save_excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(filename ? { filename } : {})
            });
            
            const result = await response.json();
            
            if (!result.success) {
                throw new Error(result.error || 'Failed to save file');
            }
            
            return result;
        },
        
        /**
         * Export chart
         */
        async exportChart(format, imageData, pdfOptions = null) {
            const body = { format, image_data: imageData };
            if (pdfOptions) {
                body.pdf_options = pdfOptions;
            }
            
            const response = await fetch('/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body)
            });
            
            if (!response.ok) {
                const result = await response.json();
                throw new Error(result.error || 'Failed to export chart');
            }
            
            return response.blob();
        }
    };

    // =============================================================================
    // Data Loading
    // =============================================================================
    
    async function loadOrgData() {
        try {
            showLoading(true);
            
            const result = await API.getOrgData();
            
            if (result.data) {
                window.OrgChart.setData(result.data);
                updateDepartmentsList(result.data.roots);
                
                state.hasUnsavedChanges = result.has_unsaved_changes;
                updateSaveStatus();
                
                if (result.source_file) {
                    window.OrgChart.setSourceFile(result.source_file);
                }
            }
            
            showLoading(false);
            
        } catch (error) {
            console.error('Error loading org data:', error);
            showToast('Failed to load organization data: ' + error.message, 'error');
            showLoading(false);
        }
    }
    
    function showLoading(show) {
        const orgChart = document.getElementById('orgChart');
        if (show) {
            orgChart.innerHTML = `
                <div class="loading-placeholder">
                    <div class="spinner"></div>
                    <p>Loading organization chart...</p>
                </div>
            `;
        }
    }
    
    /**
     * Extract unique departments from tree for autocomplete
     */
    function updateDepartmentsList(roots) {
        state.departments.clear();
        
        function traverse(nodes) {
            nodes.forEach(node => {
                if (node.department) {
                    state.departments.add(node.department);
                }
                if (node.children) {
                    traverse(node.children);
                }
            });
        }
        
        traverse(roots);
    }

    // =============================================================================
    // File Operations
    // =============================================================================
    
    function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        uploadFile(file);
        e.target.value = ''; // Reset input
    }
    
    async function uploadFile(file) {
        try {
            showToast('Uploading file...', 'info');
            
            const result = await API.uploadExcel(file);
            
            if (result.data) {
                window.OrgChart.setData(result.data);
                updateDepartmentsList(result.data.roots);
                
                state.hasUnsavedChanges = false;
                updateSaveStatus();
                
                showToast(result.message || 'File uploaded successfully', 'success');
            }
            
        } catch (error) {
            console.error('Upload error:', error);
            showToast('Upload failed: ' + error.message, 'error');
        }
    }
    
    function downloadTemplate() {
        window.location.href = '/download_template';
    }
    
    async function saveToExcel() {
        if (state.isSaving) return;
        
        try {
            state.isSaving = true;
            updateSaveStatus('Saving...');
            
            const result = await API.saveExcel();
            
            state.hasUnsavedChanges = false;
            state.isSaving = false;
            updateSaveStatus();
            
            showToast('Changes saved successfully', 'success');
            
            if (result.filename) {
                window.OrgChart.setSourceFile(result.filename);
            }
            
        } catch (error) {
            console.error('Save error:', error);
            state.isSaving = false;
            updateSaveStatus();
            showToast('Failed to save: ' + error.message, 'error');
        }
    }
    
    function updateSaveStatus(customMessage = null) {
        const status = elements.saveStatus;
        
        if (customMessage) {
            status.textContent = customMessage;
            status.className = 'save-status saving';
        } else if (state.hasUnsavedChanges) {
            status.textContent = 'Unsaved changes';
            status.className = 'save-status unsaved';
        } else {
            status.textContent = 'All changes saved';
            status.className = 'save-status saved';
        }
    }
    
    function markAsChanged() {
        state.hasUnsavedChanges = true;
        updateSaveStatus();
    }

    // =============================================================================
    // Export Modal
    // =============================================================================
    
    let exportPreviewData = null;
    
    // Page sizes in points (72 points per inch)
    const PAGE_SIZES_PT = {
        'a4': { width: 595, height: 842 },      // 210 x 297 mm
        'a3': { width: 842, height: 1191 },     // 297 x 420 mm
        'a2': { width: 1191, height: 1684 },    // 420 x 594 mm
        'a1': { width: 1684, height: 2384 },    // 594 x 841 mm
        'a0': { width: 2384, height: 3370 },    // 841 x 1189 mm
        'letter': { width: 612, height: 792 },  // 8.5 x 11 in
        'legal': { width: 612, height: 1008 },  // 8.5 x 14 in
        'tabloid': { width: 792, height: 1224 } // 11 x 17 in
    };
    
    function openExportModal() {
        openModal(elements.exportModal);
        updateExportPreview();
    }
    
    async function updateExportPreview() {
        const format = document.querySelector('input[name="exportFormat"]:checked').value;
        const quality = parseInt(document.querySelector('input[name="exportQuality"]:checked')?.value || 2);
        const dpi = document.querySelector('input[name="exportQuality"]:checked')?.dataset.dpi || 150;
        const padding = parseInt(elements.exportPadding.value) || 50;
        
        // Get element references for new groups
        const imageQualityGroup = document.getElementById('imageQualityGroup');
        const pdfSettingsGroup = document.getElementById('pdfSettingsGroup');
        const pdfOrientationGroup = document.getElementById('pdfOrientationGroup');
        const backgroundGroup = document.getElementById('backgroundGroup');
        const paddingGroup = document.getElementById('paddingGroup');
        const pdfPageSize = document.getElementById('pdfPageSize');
        
        // Determine if format needs image-related options
        const isImageFormat = ['png', 'jpeg', 'svg'].includes(format);
        const isPdf = format === 'pdf';
        const isJson = format === 'json';
        
        // Show/hide format-specific options
        elements.jpegQualityGroup.style.display = format === 'jpeg' ? 'block' : 'none';
        
        // Hide quality for SVG, PDF, and JSON
        if (imageQualityGroup) {
            imageQualityGroup.style.display = (format === 'png' || format === 'jpeg') ? 'block' : 'none';
        }
        
        // PDF settings
        if (pdfSettingsGroup) {
            pdfSettingsGroup.style.display = isPdf ? 'block' : 'none';
        }
        
        // Show/hide orientation based on page size selection
        if (pdfOrientationGroup && pdfPageSize) {
            pdfOrientationGroup.style.display = (isPdf && pdfPageSize.value !== 'auto') ? 'flex' : 'none';
        }
        
        // Hide padding for PDF and JSON
        if (paddingGroup) {
            paddingGroup.style.display = (isPdf || isJson) ? 'none' : 'block';
        }
        
        // Hide background for JSON
        if (backgroundGroup) {
            backgroundGroup.style.display = isJson ? 'none' : 'block';
        }
        
        // Handle transparent option
        const bgSelect = elements.exportBackground;
        const transparentOption = bgSelect.querySelector('option[value="transparent"]');
        if (transparentOption) {
            // Transparent only for PNG and SVG
            transparentOption.disabled = (format !== 'png' && format !== 'svg');
            if (format !== 'png' && format !== 'svg' && bgSelect.value === 'transparent') {
                bgSelect.value = 'white';
            }
        }
        
        // Get chart dimensions
        const chartInfo = window.OrgChart.getChartInfo ? window.OrgChart.getChartInfo() : { width: 1500, height: 1000 };
        const baseWidth = chartInfo.width + (padding * 2);
        const baseHeight = chartInfo.height + (padding * 2);
        const scaledWidth = Math.round(baseWidth * quality);
        const scaledHeight = Math.round(baseHeight * quality);
        
        // Update dimensions display
        if (format === 'json') {
            const nodeCount = chartInfo.nodeCount || 'N/A';
            elements.exportDimensions.textContent = `${nodeCount} nodes`;
            elements.exportResDimensions.textContent = 'Complete data export';
        } else if (format === 'svg') {
            elements.exportDimensions.textContent = `${baseWidth} × ${baseHeight} px`;
            elements.exportResDimensions.textContent = 'High quality image (PNG embedded)';
        } else if (format === 'pdf') {
            const pageSizeVal = pdfPageSize ? pdfPageSize.value : 'auto';
            const orientation = document.querySelector('input[name="pdfOrientation"]:checked')?.value || 'landscape';
            const pdfMargin = parseInt(document.getElementById('pdfMargin')?.value || 10);
            
            if (pageSizeVal === 'auto') {
                elements.exportDimensions.textContent = `Chart size: ${chartInfo.width} × ${chartInfo.height} px`;
                elements.exportResDimensions.textContent = 'Page fits chart content';
            } else {
                // Calculate actual page dimensions
                const pageInfo = PAGE_SIZES_PT[pageSizeVal] || PAGE_SIZES_PT['a4'];
                let pageW = orientation === 'landscape' ? pageInfo.height : pageInfo.width;
                let pageH = orientation === 'landscape' ? pageInfo.width : pageInfo.height;
                // Convert points to mm for display
                const pageWmm = Math.round(pageW / 72 * 25.4);
                const pageHmm = Math.round(pageH / 72 * 25.4);
                elements.exportDimensions.textContent = `Page: ${pageWmm} × ${pageHmm} mm`;
                elements.exportResDimensions.textContent = `${pageSizeVal.toUpperCase()} ${orientation} (${pdfMargin}mm margin)`;
            }
        } else {
            elements.exportDimensions.textContent = `${baseWidth} × ${baseHeight} px (base)`;
            elements.exportResDimensions.textContent = `${scaledWidth} × ${scaledHeight} px`;
        }
        
        // Estimate file size (rough approximation)
        let estSize;
        if (format === 'png') {
            estSize = (scaledWidth * scaledHeight * 0.5) / 1024 / 1024;
        } else if (format === 'jpeg') {
            const jpegQ = parseInt(elements.jpegQuality.value) / 100;
            estSize = (scaledWidth * scaledHeight * 0.2 * jpegQ) / 1024 / 1024;
        } else if (format === 'svg') {
            // SVG with embedded PNG
            estSize = (baseWidth * baseHeight * 0.6) / 1024 / 1024;
        } else if (format === 'json') {
            // JSON is small, estimate based on node count
            const nodeCount = chartInfo.nodeCount || 50;
            estSize = (nodeCount * 500) / 1024 / 1024;
        } else {
            // PDF - estimate based on content
            estSize = (chartInfo.width * chartInfo.height * 0.3) / 1024 / 1024;
        }
        
        if (estSize < 1) {
            elements.exportEstSize.textContent = `~${Math.round(estSize * 1024)} KB`;
        } else {
            elements.exportEstSize.textContent = `~${estSize.toFixed(1)} MB`;
        }
        
        // Generate preview thumbnail
        try {
            await generatePreview(format);
        } catch (e) {
            console.error('Preview generation failed:', e);
        }
    }
    
    async function generatePreview(format) {
        const bgColor = getBackgroundColor();
        const padding = parseInt(elements.exportPadding.value) || 50;
        const pdfPageSize = document.getElementById('pdfPageSize')?.value || 'auto';
        const orientation = document.querySelector('input[name="pdfOrientation"]:checked')?.value || 'landscape';
        const pdfMargin = parseInt(document.getElementById('pdfMargin')?.value || 10);
        
        // Get chart dimensions
        const chartInfo = window.OrgChart.getChartInfo ? window.OrgChart.getChartInfo() : { width: 1500, height: 1000 };
        
        // JSON format - show JSON preview
        if (format === 'json') {
            const jsonPreview = `<div style="background: #1e1e1e; color: #d4d4d4; padding: 12px; border-radius: 8px; font-family: 'Consolas', monospace; font-size: 10px; text-align: left; max-height: 160px; overflow: auto;">
<span style="color: #ce9178;">{</span>
  <span style="color: #9cdcfe;">"exportDate"</span>: <span style="color: #ce9178;">"..."</span>,
  <span style="color: #9cdcfe;">"version"</span>: <span style="color: #b5cea8;">1.0</span>,
  <span style="color: #9cdcfe;">"nodes"</span>: <span style="color: #ce9178;">[</span>
    <span style="color: #6a9955;">// ${chartInfo.nodeCount || 'N'} employees</span>
    { <span style="color: #9cdcfe;">id, name, title...</span> }
  <span style="color: #ce9178;">]</span>,
  <span style="color: #9cdcfe;">"hierarchy"</span>: <span style="color: #ce9178;">{...}</span>
<span style="color: #ce9178;">}</span>
</div>`;
            elements.exportPreviewBox.innerHTML = jsonPreview;
            exportPreviewData = null;
            return;
        }
        
        // Generate base image for image formats
        const previewData = await window.OrgChart.exportAsImage(1, bgColor, format === 'pdf' ? 20 : padding);
        
        if (format === 'pdf' && pdfPageSize !== 'auto') {
            // Create a preview showing how chart fits on page
            const previewCanvas = document.createElement('canvas');
            const ctx = previewCanvas.getContext('2d');
            
            // Get page dimensions
            const pageInfo = PAGE_SIZES_PT[pdfPageSize] || PAGE_SIZES_PT['a4'];
            let pageW = orientation === 'landscape' ? pageInfo.height : pageInfo.width;
            let pageH = orientation === 'landscape' ? pageInfo.width : pageInfo.height;
            
            // Scale to fit preview box (180px wide max)
            const previewScale = 160 / Math.max(pageW, pageH);
            const canvasW = Math.round(pageW * previewScale);
            const canvasH = Math.round(pageH * previewScale);
            
            previewCanvas.width = canvasW;
            previewCanvas.height = canvasH;
            
            // Draw page background
            ctx.fillStyle = bgColor === 'transparent' ? '#ffffff' : bgColor;
            ctx.fillRect(0, 0, canvasW, canvasH);
            
            // Draw page border
            ctx.strokeStyle = '#ccc';
            ctx.lineWidth = 1;
            ctx.strokeRect(0.5, 0.5, canvasW - 1, canvasH - 1);
            
            // Calculate margin in preview
            const marginPx = (pdfMargin / 25.4 * 72) * previewScale;
            
            // Draw margin guides (dotted)
            ctx.strokeStyle = '#e0e0e0';
            ctx.setLineDash([2, 2]);
            ctx.strokeRect(marginPx, marginPx, canvasW - 2 * marginPx, canvasH - 2 * marginPx);
            ctx.setLineDash([]);
            
            // Load the chart image and draw it scaled to fit
            const img = new Image();
            img.src = previewData;
            await new Promise(resolve => img.onload = resolve);
            
            // Calculate how chart fits in available area
            const availW = canvasW - 2 * marginPx;
            const availH = canvasH - 2 * marginPx;
            const chartW = chartInfo.width + 40; // with some padding
            const chartH = chartInfo.height + 40;
            
            const scale = Math.min(availW / chartW, availH / chartH);
            const drawW = chartW * scale;
            const drawH = chartH * scale;
            const drawX = marginPx + (availW - drawW) / 2;
            const drawY = marginPx + (availH - drawH) / 2;
            
            ctx.drawImage(img, drawX, drawY, drawW, drawH);
            
            elements.exportPreviewBox.innerHTML = `<img src="${previewCanvas.toDataURL()}" alt="PDF Preview" style="border: 1px solid #ddd; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">`;
        } else {
            elements.exportPreviewBox.innerHTML = `<img src="${previewData}" alt="Chart Preview">`;
        }
        
        exportPreviewData = previewData;
    }
    
    function getBackgroundColor() {
        const bgType = elements.exportBackground.value;
        if (bgType === 'transparent') return 'transparent';
        if (bgType === 'custom') return elements.exportBgColor.value;
        return '#ffffff';
    }
    
    // Calculate optimal quality for PDF based on page size
    function calculatePdfQuality(pageSize, orientation, chartWidth, chartHeight, margin) {
        const TARGET_DPI = 300; // Print quality DPI
        const POINTS_PER_INCH = 72;
        
        // For auto mode, use standard 2x quality
        if (pageSize === 'auto') {
            return 2;
        }
        
        // Get page dimensions in points
        const pageInfo = PAGE_SIZES_PT[pageSize] || PAGE_SIZES_PT['a4'];
        let pageW = orientation === 'landscape' ? pageInfo.height : pageInfo.width;
        let pageH = orientation === 'landscape' ? pageInfo.width : pageInfo.height;
        
        // Calculate available space after margins (margin in mm, convert to points)
        const marginPt = (margin / 25.4) * POINTS_PER_INCH;
        const availW = pageW - 2 * marginPt;
        const availH = pageH - 2 * marginPt;
        
        // Calculate what image resolution we need for TARGET_DPI print quality
        // Available width in inches * DPI = required pixels
        const requiredWidthPx = (availW / POINTS_PER_INCH) * TARGET_DPI;
        const requiredHeightPx = (availH / POINTS_PER_INCH) * TARGET_DPI;
        
        // Calculate quality multiplier needed
        const qualityForWidth = requiredWidthPx / chartWidth;
        const qualityForHeight = requiredHeightPx / chartHeight;
        
        // Use the larger of the two to ensure we have enough resolution
        // But cap at reasonable values (1x to 4x)
        let quality = Math.max(qualityForWidth, qualityForHeight);
        quality = Math.max(1, Math.min(4, Math.ceil(quality)));
        
        return quality;
    }
    
    async function doExport() {
        const format = document.querySelector('input[name="exportFormat"]:checked').value;
        const quality = parseInt(document.querySelector('input[name="exportQuality"]:checked')?.value || 2);
        const padding = parseInt(elements.exportPadding.value) || 50;
        const bgColor = getBackgroundColor();
        const jpegQuality = parseInt(elements.jpegQuality.value) / 100;
        
        try {
            showToast(`Generating ${format.toUpperCase()}...`, 'info');
            closeModal(elements.exportModal);
            
            let blob;
            let extension = format;
            
            if (format === 'json') {
                // Export complete JSON data
                const jsonData = await window.OrgChart.exportAsJSON();
                const jsonString = JSON.stringify(jsonData, null, 2);
                blob = new Blob([jsonString], { type: 'application/json' });
            } else if (format === 'svg') {
                // Generate high-quality PNG and embed in SVG for reliable rendering
                const svgData = await window.OrgChart.exportAsSVG(bgColor, padding);
                blob = new Blob([svgData], { type: 'image/svg+xml' });
            } else if (format === 'pdf') {
                // PDF with page settings
                const pdfPageSize = document.getElementById('pdfPageSize')?.value || 'auto';
                const pdfOrientation = document.querySelector('input[name="pdfOrientation"]:checked')?.value || 'landscape';
                const pdfMargin = parseInt(document.getElementById('pdfMargin')?.value || 10);
                
                // Get chart dimensions to calculate optimal quality
                const chartInfo = window.OrgChart.getChartInfo ? window.OrgChart.getChartInfo() : { width: 1500, height: 1000 };
                
                // Calculate optimal quality for this page size
                const pdfQuality = calculatePdfQuality(
                    pdfPageSize, 
                    pdfOrientation, 
                    chartInfo.width, 
                    chartInfo.height, 
                    pdfMargin
                );
                
                console.log(`PDF Export: Page=${pdfPageSize}, Quality=${pdfQuality}x, Chart=${chartInfo.width}x${chartInfo.height}`);
                
                const imageData = await window.OrgChart.exportAsImage(pdfQuality, bgColor, 20);
                
                blob = await API.exportChart('pdf', imageData, {
                    pageSize: pdfPageSize,
                    orientation: pdfOrientation,
                    margin: pdfMargin
                });
            } else {
                // Generate high quality image for PNG/JPEG
                const imageData = await window.OrgChart.exportAsImage(quality, bgColor, padding);
                
                if (format === 'jpeg') {
                    // Convert to JPEG with quality
                    const img = new Image();
                    img.src = imageData;
                    await new Promise(resolve => img.onload = resolve);
                    
                    const canvas = document.createElement('canvas');
                    canvas.width = img.width;
                    canvas.height = img.height;
                    const ctx = canvas.getContext('2d');
                    
                    // Fill background for JPEG (no transparency)
                    ctx.fillStyle = bgColor === 'transparent' ? '#ffffff' : bgColor;
                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                    ctx.drawImage(img, 0, 0);
                    
                    blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', jpegQuality));
                } else {
                    // PNG - use API to get proper blob
                    blob = await API.exportChart('png', imageData);
                }
            }
            
            // Download
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `org_chart_${formatDate(new Date())}.${extension}`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            
            showToast(`${format.toUpperCase()} exported successfully`, 'success');
            
        } catch (error) {
            console.error('Export error:', error);
            showToast('Export failed: ' + error.message, 'error');
        }
    }
    
    function formatDate(date) {
        return date.toISOString().slice(0, 10).replace(/-/g, '');
    }

    // =============================================================================
    // Modal Management
    // =============================================================================
    
    function openModal(modal) {
        modal.classList.add('show');
        document.body.style.overflow = 'hidden';
        
        // Focus first input
        const firstInput = modal.querySelector('input:not([type="hidden"]), select');
        if (firstInput) {
            setTimeout(() => firstInput.focus(), 100);
        }
    }
    
    function closeModal(modal) {
        modal.classList.remove('show');
        document.body.style.overflow = '';
    }
    
    function closeAllModals() {
        document.querySelectorAll('.modal.show').forEach(modal => {
            closeModal(modal);
        });
    }
    
    /**
     * Populate manager dropdown with all employees except specified node and its descendants
     */
    function populateManagerDropdown(selectElement, excludeNodeId = null) {
        const nodes = window.OrgChart.getAllNodes();
        
        // Get descendants of excluded node
        const excludeIds = new Set();
        if (excludeNodeId) {
            excludeIds.add(excludeNodeId);
            
            function addDescendants(nodeId) {
                const node = window.OrgChart.getNode(nodeId);
                if (node && node.children) {
                    node.children.forEach(child => {
                        excludeIds.add(child.id);
                        addDescendants(child.id);
                    });
                }
            }
            addDescendants(excludeNodeId);
        }
        
        // Build options
        selectElement.innerHTML = '<option value="">None (Root Node)</option>';
        
        nodes
            .filter(node => !excludeIds.has(node.id))
            .sort((a, b) => a.name.localeCompare(b.name))
            .forEach(node => {
                const option = document.createElement('option');
                option.value = node.id;
                option.textContent = `${node.name} (${node.title})`;
                selectElement.appendChild(option);
            });
    }

    // =============================================================================
    // Edit Modal
    // =============================================================================
    
    function openEditModal(nodeId) {
        const node = window.OrgChart.getNode(nodeId);
        if (!node) {
            showToast('Node not found', 'error');
            return;
        }
        
        state.currentEditNodeId = nodeId;
        
        // Populate form
        elements.editNodeId.value = nodeId;
        elements.editName.value = node.name || '';
        elements.editTitle.value = node.title || '';
        elements.editDepartment.value = node.department || '';
        elements.editColor.value = node.color || '#1a73e8';
        elements.editColorHex.textContent = node.color || '#1a73e8';
        elements.editAvatarUrl.value = node.avatar_url || '';
        elements.editWhatsapp.value = node.whatsapp || '';
        
        // Populate manager dropdown
        populateManagerDropdown(elements.editManager, nodeId);
        elements.editManager.value = node.manager_id || '';
        
        elements.editModalTitle.textContent = `Edit: ${node.name}`;
        
        openModal(elements.editModal);
    }
    
    async function handleEditSubmit(e) {
        e.preventDefault();
        
        const nodeId = elements.editNodeId.value;
        
        const nodeData = {
            id: nodeId,
            name: elements.editName.value.trim(),
            title: elements.editTitle.value.trim(),
            department: elements.editDepartment.value.trim(),
            manager_id: elements.editManager.value || null,
            color: elements.editColor.value,
            avatar_url: elements.editAvatarUrl.value.trim() || null,
            whatsapp: elements.editWhatsapp.value.trim() || null
        };
        
        // Validate
        if (!nodeData.name || !nodeData.title || !nodeData.department) {
            showToast('Please fill in all required fields', 'warning');
            return;
        }
        
        try {
            const result = await API.updateNode(nodeData);
            
            if (result.data) {
                window.OrgChart.setData(result.data);
                updateDepartmentsList(result.data.roots);
            }
            
            markAsChanged();
            closeModal(elements.editModal);
            showToast('Employee updated successfully', 'success');
            
        } catch (error) {
            console.error('Update error:', error);
            showToast('Failed to update: ' + error.message, 'error');
        }
    }

    // =============================================================================
    // Add Modal
    // =============================================================================
    
    function openAddModal(managerId = null) {
        // Reset form
        elements.addForm.reset();
        elements.addManagerId.value = managerId || '';
        elements.addColor.value = '#1a73e8';
        elements.addColorHex.textContent = '#1a73e8';
        elements.addWhatsapp.value = '';
        
        // Populate manager dropdown
        populateManagerDropdown(elements.addManager);
        elements.addManager.value = managerId || '';
        
        // Pre-fill department if adding subordinate
        if (managerId) {
            const manager = window.OrgChart.getNode(managerId);
            if (manager) {
                elements.addDepartment.value = manager.department || '';
                elements.addColor.value = manager.color || '#1a73e8';
                elements.addColorHex.textContent = manager.color || '#1a73e8';
            }
        }
        
        openModal(elements.addModal);
    }
    
    async function handleAddSubmit(e) {
        e.preventDefault();
        
        const nodeData = {
            name: elements.addName.value.trim(),
            title: elements.addTitle.value.trim(),
            department: elements.addDepartment.value.trim(),
            manager_id: elements.addManager.value || null,
            color: elements.addColor.value,
            whatsapp: elements.addWhatsapp.value.trim() || null
        };
        
        // Validate
        if (!nodeData.name || !nodeData.title || !nodeData.department) {
            showToast('Please fill in all required fields', 'warning');
            return;
        }
        
        try {
            const result = await API.addNode(nodeData);
            
            if (result.data) {
                window.OrgChart.setData(result.data);
                updateDepartmentsList(result.data.roots);
            }
            
            markAsChanged();
            closeModal(elements.addModal);
            showToast('Employee added successfully', 'success');
            
            // Center on new node
            if (result.node) {
                setTimeout(() => {
                    window.OrgChart.centerOnNode(result.node.id);
                }, 100);
            }
            
        } catch (error) {
            console.error('Add error:', error);
            showToast('Failed to add employee: ' + error.message, 'error');
        }
    }

    // =============================================================================
    // Delete Modal
    // =============================================================================
    
    function openDeleteModal(nodeId) {
        const node = window.OrgChart.getNode(nodeId);
        if (!node) {
            showToast('Node not found', 'error');
            return;
        }
        
        state.currentDeleteNodeId = nodeId;
        elements.deleteNodeId.value = nodeId;
        elements.deleteNodeName.textContent = node.name;
        
        // Show reassignment options if node has children
        const hasChildren = node.children && node.children.length > 0;
        elements.deleteReassignSection.style.display = hasChildren ? 'block' : 'none';
        
        if (hasChildren) {
            // Populate reassignment dropdown (exclude this node and its descendants)
            populateManagerDropdown(elements.deleteReassignTo, nodeId);
            
            // Default to parent
            if (node.manager_id) {
                elements.deleteReassignTo.value = node.manager_id;
            }
        }
        
        openModal(elements.deleteModal);
    }
    
    async function handleDelete() {
        const nodeId = elements.deleteNodeId.value;
        const reassignTo = elements.deleteReassignTo.value || null;
        
        try {
            const result = await API.deleteNode(nodeId, reassignTo);
            
            if (result.data) {
                window.OrgChart.setData(result.data);
                updateDepartmentsList(result.data.roots);
            }
            
            markAsChanged();
            closeModal(elements.deleteModal);
            showToast('Employee deleted successfully', 'success');
            
        } catch (error) {
            console.error('Delete error:', error);
            showToast('Failed to delete: ' + error.message, 'error');
        }
    }

    // =============================================================================
    // Department Autocomplete
    // =============================================================================
    
    function showDepartmentSuggestions(input, suggestionsElement) {
        const value = input.value.toLowerCase().trim();
        
        if (!value) {
            suggestionsElement.classList.remove('show');
            return;
        }
        
        const matches = Array.from(state.departments)
            .filter(dept => dept.toLowerCase().includes(value))
            .slice(0, 5);
        
        if (matches.length === 0) {
            suggestionsElement.classList.remove('show');
            return;
        }
        
        suggestionsElement.innerHTML = matches.map(dept => 
            `<div class="suggestion-item" data-value="${escapeHtml(dept)}">${escapeHtml(dept)}</div>`
        ).join('');
        
        suggestionsElement.classList.add('show');
    }
    
    function handleSuggestionClick(e, input, suggestionsElement) {
        const item = e.target.closest('.suggestion-item');
        if (item) {
            input.value = item.dataset.value;
            suggestionsElement.classList.remove('show');
        }
    }

    // =============================================================================
    // Color Picker
    // =============================================================================
    
    function handleColorChange(colorInput, colorHexDisplay) {
        colorHexDisplay.textContent = colorInput.value;
    }
    
    function handleColorPresetClick(e, colorInput, colorHexDisplay) {
        const preset = e.target.closest('.color-preset');
        if (preset) {
            const color = preset.dataset.color;
            colorInput.value = color;
            colorHexDisplay.textContent = color;
        }
    }

    // =============================================================================
    // Toast Notifications
    // =============================================================================
    
    function showToast(message, type = 'info') {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        
        const iconSvg = getToastIcon(type);
        toast.innerHTML = `${iconSvg}<span>${escapeHtml(message)}</span>`;
        
        elements.toastContainer.appendChild(toast);
        
        // Auto remove after 4 seconds
        setTimeout(() => {
            toast.classList.add('toast-out');
            setTimeout(() => toast.remove(), 200);
        }, 4000);
    }
    
    function getToastIcon(type) {
        switch (type) {
            case 'success':
                return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polyline points="20 6 9 17 4 12"/>
                </svg>`;
            case 'error':
                return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="15" y1="9" x2="9" y2="15"/>
                    <line x1="9" y1="9" x2="15" y2="15"/>
                </svg>`;
            case 'warning':
                return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
                    <line x1="12" y1="9" x2="12" y2="13"/>
                    <line x1="12" y1="17" x2="12.01" y2="17"/>
                </svg>`;
            default:
                return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="12" y1="16" x2="12" y2="12"/>
                    <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>`;
        }
    }

    // =============================================================================
    // Manager Update (from drag-drop)
    // =============================================================================
    
    async function updateNodeManager(nodeId, newManagerId) {
        try {
            const node = window.OrgChart.getNode(nodeId);
            if (!node) return;
            
            const result = await API.updateNode({
                id: nodeId,
                manager_id: newManagerId
            });
            
            if (result.data) {
                window.OrgChart.setData(result.data);
                updateDepartmentsList(result.data.roots);
            }
            
            markAsChanged();
            showToast('Manager updated successfully', 'success');
            
        } catch (error) {
            console.error('Update manager error:', error);
            showToast('Failed to update manager: ' + error.message, 'error');
        }
    }

    // =============================================================================
    // Utilities
    // =============================================================================
    
    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // =============================================================================
    // Event Binding
    // =============================================================================
    
    function bindEvents() {
        // File operations
        elements.btnUpload.addEventListener('click', () => elements.fileInput.click());
        elements.fileInput.addEventListener('change', handleFileUpload);
        elements.btnDownloadTemplate.addEventListener('click', downloadTemplate);
        elements.btnSave.addEventListener('click', saveToExcel);
        
        // Add root node
        elements.btnAddRoot.addEventListener('click', () => openAddModal());
        
        // Export modal
        elements.btnOpenExport.addEventListener('click', openExportModal);
        elements.btnDoExport.addEventListener('click', doExport);
        
        // Export settings listeners
        document.querySelectorAll('input[name="exportFormat"]').forEach(radio => {
            radio.addEventListener('change', updateExportPreview);
        });
        document.querySelectorAll('input[name="exportQuality"]').forEach(radio => {
            radio.addEventListener('change', updateExportPreview);
        });
        elements.exportBackground.addEventListener('change', (e) => {
            elements.customBgGroup.style.display = e.target.value === 'custom' ? 'block' : 'none';
            // Disable transparent for non-PNG/SVG
            const format = document.querySelector('input[name="exportFormat"]:checked').value;
            if (e.target.value === 'transparent' && format !== 'png' && format !== 'svg') {
                e.target.value = 'white';
                showToast('Transparent background only available for PNG and SVG', 'warning');
            }
            updateExportPreview();
        });
        elements.exportPadding.addEventListener('change', updateExportPreview);
        elements.jpegQuality.addEventListener('input', () => {
            elements.jpegQualityValue.textContent = elements.jpegQuality.value + '%';
        });
        
        // PDF settings listeners
        const pdfPageSize = document.getElementById('pdfPageSize');
        if (pdfPageSize) {
            pdfPageSize.addEventListener('change', updateExportPreview);
        }
        document.querySelectorAll('input[name="pdfOrientation"]').forEach(radio => {
            radio.addEventListener('change', updateExportPreview);
        });
        const pdfMargin = document.getElementById('pdfMargin');
        if (pdfMargin) {
            pdfMargin.addEventListener('change', updateExportPreview);
            pdfMargin.addEventListener('input', updateExportPreview);
        }
        
        // Edit modal
        elements.editForm.addEventListener('submit', handleEditSubmit);
        elements.editColor.addEventListener('input', () => {
            handleColorChange(elements.editColor, elements.editColorHex);
        });
        
        // Color presets in edit modal
        elements.editModal.querySelectorAll('.color-preset').forEach(preset => {
            preset.addEventListener('click', (e) => {
                handleColorPresetClick(e, elements.editColor, elements.editColorHex);
            });
        });
        
        // Department autocomplete in edit modal
        elements.editDepartment.addEventListener('input', () => {
            showDepartmentSuggestions(elements.editDepartment, elements.departmentSuggestions);
        });
        elements.departmentSuggestions.addEventListener('click', (e) => {
            handleSuggestionClick(e, elements.editDepartment, elements.departmentSuggestions);
        });
        
        // Add modal
        elements.addForm.addEventListener('submit', handleAddSubmit);
        elements.addColor.addEventListener('input', () => {
            handleColorChange(elements.addColor, elements.addColorHex);
        });
        
        // Delete modal
        elements.btnConfirmDelete.addEventListener('click', handleDelete);
        
        // Modal close buttons
        document.querySelectorAll('[data-close-modal]').forEach(btn => {
            btn.addEventListener('click', () => {
                const modal = btn.closest('.modal');
                if (modal) closeModal(modal);
            });
        });
        
        // Close modal on overlay click
        document.querySelectorAll('.modal-overlay').forEach(overlay => {
            overlay.addEventListener('click', () => {
                const modal = overlay.closest('.modal');
                if (modal) closeModal(modal);
            });
        });
        
        // Close modal on Escape
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                closeAllModals();
            }
        });
        
        // Warn before leaving with unsaved changes
        window.addEventListener('beforeunload', (e) => {
            if (state.hasUnsavedChanges) {
                e.preventDefault();
                e.returnValue = '';
            }
        });
        
        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            // Ctrl+S to save
            if (e.ctrlKey && e.key === 's') {
                e.preventDefault();
                saveToExcel();
            }
        });
    }

    // =============================================================================
    // Initialization
    // =============================================================================
    
    function init() {
        initElements();
        bindEvents();
        
        // Initialize org chart
        window.OrgChart.init();
        
        // Load initial data
        loadOrgData();
    }

    // =============================================================================
    // Public API
    // =============================================================================
    
    window.ExcelSync = {
        init,
        loadOrgData,
        openEditModal,
        openAddModal,
        openDeleteModal,
        showToast,
        updateNodeManager,
        markAsChanged
    };

    // Initialize when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

})();
