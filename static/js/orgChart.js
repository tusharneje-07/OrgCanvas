/**
 * Organization Chart Renderer
 * Handles tree rendering, zoom/pan, drag-drop, and node interactions
 */

(function() {
    'use strict';

    // =============================================================================
    // Configuration
    // =============================================================================
    
    const CONFIG = {
        // Node dimensions
        nodeWidth: 150,
        nodeMinHeight: 115,
        nodeGapX: 15,
        nodeGapY: 70,
        
        // Zoom settings
        minZoom: 0.1,
        maxZoom: 2,
        zoomStep: 0.1,
        
        // Animation
        animationDuration: 300,
        
        // Layout
        defaultLayout: 'vertical', // 'vertical' or 'horizontal'
        
        // Padding around chart
        chartPadding: 40
    };

    // =============================================================================
    // State Management
    // =============================================================================
    
    const state = {
        // Data
        orgData: null,
        flatNodes: new Map(), // id -> node data
        
        // View
        zoom: 1,
        panX: 0,
        panY: 0,
        layout: CONFIG.defaultLayout,
        
        // Settings
        showWhatsApp: true,
        closeness: 1, // 0.5 = compact, 1 = normal, 2 = wide
        borderLineColor: '#90a4ae', // Color for borders and connecting lines
        lineCornerStyle: 'hard', // 'hard' or 'curved'
        cardCornerStyle: 'curved', // 'hard' or 'curved'
        
        // Interaction
        isPanning: false,
        panStartX: 0,
        panStartY: 0,
        selectedNodeId: null,
        
        // Drag & Drop
        isDragging: false,
        dragNodeId: null,
        dragStartX: 0,
        dragStartY: 0,
        
        // Computed positions
        nodePositions: new Map(), // id -> {x, y, width, height}
        
        // Chart dimensions
        chartWidth: 0,
        chartHeight: 0
    };

    // =============================================================================
    // DOM Elements
    // =============================================================================
    
    let elements = {};

    function initElements() {
        elements = {
            chartContainer: document.getElementById('chartContainer'),
            orgChart: document.getElementById('orgChart'),
            searchInput: document.getElementById('searchInput'),
            searchResults: document.getElementById('searchResults'),
            contextMenu: document.getElementById('contextMenu'),
            dragPreview: document.getElementById('dragPreview'),
            zoomLevel: document.getElementById('zoomLevel'),
            minimap: document.getElementById('minimap'),
            minimapCanvas: document.getElementById('minimapCanvas'),
            minimapViewport: document.getElementById('minimapViewport'),
            
            // Info displays
            infoTotalEmployees: document.getElementById('infoTotalEmployees'),
            infoMaxDepth: document.getElementById('infoMaxDepth'),
            infoSourceFile: document.getElementById('infoSourceFile'),
            
            // Buttons
            btnLayoutVertical: document.getElementById('btnLayoutVertical'),
            btnLayoutHorizontal: document.getElementById('btnLayoutHorizontal'),
            btnZoomIn: document.getElementById('btnZoomIn'),
            btnZoomOut: document.getElementById('btnZoomOut'),
            btnZoomFit: document.getElementById('btnZoomFit'),
            btnZoomReset: document.getElementById('btnZoomReset')
        };
    }

    // =============================================================================
    // Tree Layout Algorithm
    // =============================================================================
    
    /**
     * Calculate positions for all nodes in the tree using a modified Reingold-Tilford algorithm.
     * This ensures proper spacing and centering of subtrees.
     * 
     * Algorithm:
     * 1. First pass (bottom-up): Calculate subtree widths
     * 2. Second pass (top-down): Assign x,y positions based on parent position and subtree width
     */
    function calculateLayout(roots) {
        if (!roots || roots.length === 0) return;
        
        state.nodePositions.clear();
        
        const isHorizontal = state.layout === 'horizontal';
        const nodeWidth = CONFIG.nodeWidth;
        // Node height is dynamic based on whether WhatsApp numbers are shown
        // With WhatsApp: header(24) + body(~91) = 115px
        // Without WhatsApp: header(24) + body(~71) = 95px
        const nodeHeight = state.showWhatsApp ? 115 : 95;
        // Apply closeness multiplier to horizontal gap
        const gapX = Math.round(CONFIG.nodeGapX * state.closeness);
        const gapY = CONFIG.nodeGapY;
        
        /**
         * First pass: Calculate the width of each subtree
         * Returns the total width needed for this node and all its descendants
         */
        function calculateSubtreeWidth(node) {
            if (!node.children || node.children.length === 0 || !node.expanded) {
                return isHorizontal ? nodeHeight : nodeWidth;
            }
            
            let totalWidth = 0;
            node.children.forEach((child, index) => {
                if (index > 0) totalWidth += isHorizontal ? gapY : gapX;
                totalWidth += calculateSubtreeWidth(child);
            });
            
            return Math.max(isHorizontal ? nodeHeight : nodeWidth, totalWidth);
        }
        
        /**
         * Second pass: Assign positions to each node
         * @param node - Current node
         * @param x - X position for this node's center
         * @param y - Y position for this node's top
         * @param subtreeWidth - Width available for this subtree
         */
        function assignPositions(node, x, y, subtreeWidth) {
            // Calculate actual height for this node based on whether it has WhatsApp and if it's shown
            const hasVisiblePhone = state.showWhatsApp && node.whatsapp;
            const actualNodeHeight = hasVisiblePhone ? 115 : 95;
            
            // Store position (center x, top y)
            state.nodePositions.set(node.id, {
                x: x,
                y: y,
                width: nodeWidth,
                height: actualNodeHeight,
                node: node
            });
            
            // Position children
            if (node.children && node.children.length > 0 && node.expanded !== false) {
                let childrenTotalWidth = 0;
                const childWidths = [];
                
                node.children.forEach((child, index) => {
                    const childWidth = calculateSubtreeWidth(child);
                    childWidths.push(childWidth);
                    if (index > 0) childrenTotalWidth += isHorizontal ? gapY : gapX;
                    childrenTotalWidth += childWidth;
                });
                
                // Starting position for first child
                let currentPos;
                if (isHorizontal) {
                    currentPos = y + nodeHeight / 2 - childrenTotalWidth / 2;
                } else {
                    currentPos = x - childrenTotalWidth / 2;
                }
                
                node.children.forEach((child, index) => {
                    const childWidth = childWidths[index];
                    const childCenter = currentPos + childWidth / 2;
                    
                    if (isHorizontal) {
                        assignPositions(
                            child,
                            x + nodeWidth + gapX,
                            childCenter - nodeHeight / 2,
                            childWidth
                        );
                    } else {
                        assignPositions(
                            child,
                            childCenter,
                            y + nodeHeight + gapY,
                            childWidth
                        );
                    }
                    
                    currentPos += childWidth + (isHorizontal ? gapY : gapX);
                });
            }
        }
        
        // Handle multiple roots
        let totalRootsWidth = 0;
        const rootWidths = [];
        
        roots.forEach((root, index) => {
            const rootWidth = calculateSubtreeWidth(root);
            rootWidths.push(rootWidth);
            if (index > 0) totalRootsWidth += isHorizontal ? gapY : gapX;
            totalRootsWidth += rootWidth;
        });
        
        // Position each root
        let currentPos = CONFIG.chartPadding;
        
        roots.forEach((root, index) => {
            const rootWidth = rootWidths[index];
            const rootCenter = currentPos + rootWidth / 2;
            
            if (isHorizontal) {
                assignPositions(
                    root,
                    CONFIG.chartPadding,
                    rootCenter,
                    rootWidth
                );
            } else {
                assignPositions(
                    root,
                    rootCenter,
                    CONFIG.chartPadding,
                    rootWidth
                );
            }
            
            currentPos += rootWidth + (isHorizontal ? gapY : gapX);
        });
        
        // Calculate chart dimensions
        let maxX = 0, maxY = 0;
        state.nodePositions.forEach(pos => {
            maxX = Math.max(maxX, pos.x + pos.width / 2);
            maxY = Math.max(maxY, pos.y + pos.height);
        });
        
        state.chartWidth = maxX + CONFIG.chartPadding;
        state.chartHeight = maxY + CONFIG.chartPadding;
    }

    // =============================================================================
    // Rendering
    // =============================================================================
    
    /**
     * Render the complete organization chart
     */
    function renderChart() {
        if (!state.orgData || !state.orgData.roots) {
            elements.orgChart.innerHTML = `
                <div class="loading-placeholder">
                    <p>No organization data loaded. Upload an Excel file to get started.</p>
                </div>
            `;
            return;
        }
        
        // Calculate layout
        calculateLayout(state.orgData.roots);
        
        // Create chart HTML
        const chartHTML = `
            <svg class="org-lines" id="orgLines"></svg>
            <div class="org-nodes" id="orgNodes"></div>
        `;
        elements.orgChart.innerHTML = chartHTML;
        
        // Render nodes
        const nodesContainer = document.getElementById('orgNodes');
        renderNodes(state.orgData.roots, nodesContainer);
        
        // Render connecting lines
        renderLines();
        
        // Set chart size
        elements.orgChart.style.width = state.chartWidth + 'px';
        elements.orgChart.style.height = state.chartHeight + 'px';
        
        // Update minimap
        updateMinimap();
        
        // Update info display
        updateInfoDisplay();
    }
    
    /**
     * Render nodes recursively
     */
    function renderNodes(nodes, container) {
        nodes.forEach(node => {
            const pos = state.nodePositions.get(node.id);
            if (!pos) return;
            
            const nodeElement = createNodeElement(node, pos);
            container.appendChild(nodeElement);
            
            // Render children
            if (node.children && node.children.length > 0 && node.expanded !== false) {
                renderNodes(node.children, container);
            }
        });
    }
    
    /**
     * Create a single node DOM element
     */
    function createNodeElement(node, pos) {
        const div = document.createElement('div');
        div.className = 'org-node';
        div.dataset.nodeId = node.id;
        div.draggable = true;
        
        if (state.selectedNodeId === node.id) {
            div.classList.add('selected');
        }
        
        // Position the node
        div.style.position = 'absolute';
        div.style.left = (pos.x - CONFIG.nodeWidth / 2) + 'px';
        div.style.top = pos.y + 'px';
        div.style.width = CONFIG.nodeWidth + 'px';
        
        // Get initials for avatar
        const initials = getInitials(node.name);
        const avatarBg = node.color || '#757575';
        
        // Check if has children for expand button
        const hasChildren = node.children && node.children.length > 0;
        const isExpanded = node.expanded !== false;
        
        // Build WhatsApp display - plain text link
        let whatsappHtml = '';
        if (node.whatsapp && state.showWhatsApp) {
            const cleanNumber = node.whatsapp.replace(/[^0-9]/g, '');
            const displayNumber = '+' + cleanNumber;
            whatsappHtml = `<a href="https://wa.me/${cleanNumber}" class="node-phone" target="_blank" rel="noopener noreferrer" onclick="event.stopPropagation();" title="WhatsApp">${displayNumber}</a>`;
        }

        div.innerHTML = `
            <div class="node-header" style="background-color: ${node.color || '#757575'}"></div>
            <div class="node-body">
                <div class="node-avatar" style="background-color: ${avatarBg}">
                    ${node.avatar_url 
                        ? `<img src="${escapeHtml(node.avatar_url)}" alt="${escapeHtml(node.name)}" onerror="this.parentElement.innerHTML='${initials}'">`
                        : initials
                    }
                </div>
                <div class="node-name">${escapeHtml(node.name)}</div>
                <div class="node-title">${escapeHtml(node.title)}</div>
                ${whatsappHtml}
            </div>
            ${hasChildren ? `
                <button class="node-expand-btn ${isExpanded ? '' : 'collapsed'}" 
                        data-node-id="${node.id}" 
                        title="${isExpanded ? 'Collapse' : 'Expand'}">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="6 9 12 15 18 9"/>
                    </svg>
                </button>
            ` : ''}
        `;
        
        return div;
    }
    
    /**
     * Render SVG connecting lines between nodes
     */
    function renderLines() {
        const svg = document.getElementById('orgLines');
        if (!svg) return;
        
        svg.innerHTML = '';
        svg.setAttribute('width', state.chartWidth);
        svg.setAttribute('height', state.chartHeight);
        
        const isHorizontal = state.layout === 'horizontal';
        
        // Draw lines for each parent-child relationship
        state.nodePositions.forEach((pos, nodeId) => {
            const node = pos.node;
            if (!node.children || node.children.length === 0 || node.expanded === false) return;
            
            node.children.forEach(child => {
                const childPos = state.nodePositions.get(child.id);
                if (!childPos) return;
                
                const line = createConnectingLine(pos, childPos, isHorizontal);
                svg.appendChild(line);
            });
        });
    }
    
    /**
     * Create an SVG path for connecting two nodes
     */
    function createConnectingLine(parentPos, childPos, isHorizontal) {
        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.classList.add('org-line');
        
        let d;
        const cornerRadius = 12; // Radius for curved corners
        
        if (isHorizontal) {
            // Horizontal layout: lines go left to right
            const startX = parentPos.x + CONFIG.nodeWidth / 2;
            const startY = parentPos.y + parentPos.height / 2;
            const endX = childPos.x - CONFIG.nodeWidth / 2;
            const endY = childPos.y + childPos.height / 2;
            const midX = (startX + endX) / 2;
            
            if (state.lineCornerStyle === 'curved') {
                d = `M ${startX} ${startY} 
                     C ${midX} ${startY}, ${midX} ${endY}, ${endX} ${endY}`;
            } else {
                // Hard corners for horizontal
                d = `M ${startX} ${startY} 
                     L ${midX} ${startY} 
                     L ${midX} ${endY} 
                     L ${endX} ${endY}`;
            }
        } else {
            // Vertical layout: lines go top to bottom
            const startX = parentPos.x;
            const startY = parentPos.y + parentPos.height;
            const endX = childPos.x;
            const endY = childPos.y;
            const midY = startY + (endY - startY) / 2;
            
            if (state.lineCornerStyle === 'curved') {
                // Curved corners using quadratic bezier curves
                const r = Math.min(cornerRadius, Math.abs(endX - startX) / 2, (endY - startY) / 4);
                
                if (startX === endX) {
                    // Straight vertical line
                    d = `M ${startX} ${startY} L ${endX} ${endY}`;
                } else if (endX > startX) {
                    // Child is to the right
                    d = `M ${startX} ${startY} 
                         L ${startX} ${midY - r} 
                         Q ${startX} ${midY} ${startX + r} ${midY} 
                         L ${endX - r} ${midY} 
                         Q ${endX} ${midY} ${endX} ${midY + r} 
                         L ${endX} ${endY}`;
                } else {
                    // Child is to the left
                    d = `M ${startX} ${startY} 
                         L ${startX} ${midY - r} 
                         Q ${startX} ${midY} ${startX - r} ${midY} 
                         L ${endX + r} ${midY} 
                         Q ${endX} ${midY} ${endX} ${midY + r} 
                         L ${endX} ${endY}`;
                }
            } else {
                // Hard corners (original)
                d = `M ${startX} ${startY} 
                     L ${startX} ${midY} 
                     L ${endX} ${midY} 
                     L ${endX} ${endY}`;
            }
        }
        
        path.setAttribute('d', d);
        return path;
    }

    // =============================================================================
    // Zoom & Pan
    // =============================================================================
    
    function setZoom(newZoom, centerX, centerY) {
        const oldZoom = state.zoom;
        state.zoom = Math.max(CONFIG.minZoom, Math.min(CONFIG.maxZoom, newZoom));
        
        // Adjust pan to zoom towards the center point
        if (centerX !== undefined && centerY !== undefined) {
            state.panX = centerX - (centerX - state.panX) * (state.zoom / oldZoom);
            state.panY = centerY - (centerY - state.panY) * (state.zoom / oldZoom);
        }
        
        applyTransform();
        updateZoomDisplay();
        updateMinimap();
    }
    
    function zoomIn() {
        const rect = elements.chartContainer.getBoundingClientRect();
        setZoom(state.zoom + CONFIG.zoomStep, rect.width / 2, rect.height / 2);
    }
    
    function zoomOut() {
        const rect = elements.chartContainer.getBoundingClientRect();
        setZoom(state.zoom - CONFIG.zoomStep, rect.width / 2, rect.height / 2);
    }
    
    function zoomReset() {
        state.zoom = 1;
        state.panX = 0;
        state.panY = 0;
        applyTransform();
        updateZoomDisplay();
        updateMinimap();
    }
    
    function zoomFit() {
        if (!state.chartWidth || !state.chartHeight) return;
        
        const containerRect = elements.chartContainer.getBoundingClientRect();
        const scaleX = (containerRect.width - 40) / state.chartWidth;
        const scaleY = (containerRect.height - 40) / state.chartHeight;
        
        state.zoom = Math.min(scaleX, scaleY, 1);
        
        // Center the chart
        state.panX = (containerRect.width - state.chartWidth * state.zoom) / 2;
        state.panY = (containerRect.height - state.chartHeight * state.zoom) / 2;
        
        applyTransform();
        updateZoomDisplay();
        updateMinimap();
    }
    
    function applyTransform() {
        elements.orgChart.style.transform = 
            `translate(${state.panX}px, ${state.panY}px) scale(${state.zoom})`;
    }
    
    function updateZoomDisplay() {
        elements.zoomLevel.textContent = Math.round(state.zoom * 100) + '%';
    }
    
    // Pan handling
    function startPan(e) {
        if (e.target.closest('.org-node') || e.target.closest('.node-expand-btn')) return;
        
        state.isPanning = true;
        state.panStartX = e.clientX - state.panX;
        state.panStartY = e.clientY - state.panY;
        elements.chartContainer.classList.add('grabbing');
    }
    
    function doPan(e) {
        if (!state.isPanning) return;
        
        state.panX = e.clientX - state.panStartX;
        state.panY = e.clientY - state.panStartY;
        applyTransform();
        updateMinimap();
    }
    
    function endPan() {
        state.isPanning = false;
        elements.chartContainer.classList.remove('grabbing');
    }
    
    // Mouse wheel zoom
    function handleWheel(e) {
        e.preventDefault();
        
        const rect = elements.chartContainer.getBoundingClientRect();
        const mouseX = e.clientX - rect.left;
        const mouseY = e.clientY - rect.top;
        
        const delta = e.deltaY > 0 ? -CONFIG.zoomStep : CONFIG.zoomStep;
        setZoom(state.zoom + delta, mouseX, mouseY);
    }

    // =============================================================================
    // Minimap
    // =============================================================================
    
    function updateMinimap() {
        const canvas = elements.minimapCanvas;
        const ctx = canvas.getContext('2d');
        
        if (!state.chartWidth || !state.chartHeight) return;
        
        // Set canvas size
        const minimapRect = elements.minimap.getBoundingClientRect();
        canvas.width = minimapRect.width;
        canvas.height = minimapRect.height;
        
        // Calculate scale
        const scaleX = canvas.width / state.chartWidth;
        const scaleY = canvas.height / state.chartHeight;
        const scale = Math.min(scaleX, scaleY) * 0.9;
        
        // Clear canvas
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        
        // Draw nodes as rectangles
        ctx.fillStyle = '#90a4ae';
        state.nodePositions.forEach(pos => {
            const x = (pos.x - CONFIG.nodeWidth / 2) * scale + 5;
            const y = pos.y * scale + 5;
            const width = CONFIG.nodeWidth * scale;
            const height = pos.height * scale;
            
            ctx.fillRect(x, y, width, height);
        });
        
        // Update viewport indicator
        const containerRect = elements.chartContainer.getBoundingClientRect();
        const viewportWidth = containerRect.width / state.zoom * scale;
        const viewportHeight = containerRect.height / state.zoom * scale;
        const viewportX = -state.panX / state.zoom * scale + 5;
        const viewportY = -state.panY / state.zoom * scale + 5;
        
        elements.minimapViewport.style.width = viewportWidth + 'px';
        elements.minimapViewport.style.height = viewportHeight + 'px';
        elements.minimapViewport.style.left = viewportX + 'px';
        elements.minimapViewport.style.top = viewportY + 'px';
    }

    // =============================================================================
    // Node Interactions
    // =============================================================================
    
    function handleNodeClick(e) {
        const nodeElement = e.target.closest('.org-node');
        if (!nodeElement) return;
        
        const nodeId = nodeElement.dataset.nodeId;
        selectNode(nodeId);
    }
    
    function selectNode(nodeId) {
        // Deselect previous
        if (state.selectedNodeId) {
            const prevNode = document.querySelector(`[data-node-id="${state.selectedNodeId}"]`);
            if (prevNode) prevNode.classList.remove('selected');
        }
        
        state.selectedNodeId = nodeId;
        
        // Select new
        const newNode = document.querySelector(`[data-node-id="${nodeId}"]`);
        if (newNode) newNode.classList.add('selected');
    }
    
    function handleNodeDoubleClick(e) {
        const nodeElement = e.target.closest('.org-node');
        if (!nodeElement) return;
        
        const nodeId = nodeElement.dataset.nodeId;
        window.ExcelSync.openEditModal(nodeId);
    }
    
    function handleExpandCollapse(e) {
        const btn = e.target.closest('.node-expand-btn');
        if (!btn) return;
        
        e.stopPropagation();
        
        const nodeId = btn.dataset.nodeId;
        const node = state.flatNodes.get(nodeId);
        
        if (node) {
            node.expanded = node.expanded === false ? true : false;
            btn.classList.toggle('collapsed');
            renderChart();
        }
    }
    
    function centerOnNode(nodeId) {
        const pos = state.nodePositions.get(nodeId);
        if (!pos) return;
        
        const containerRect = elements.chartContainer.getBoundingClientRect();
        
        // Calculate pan to center the node
        state.panX = containerRect.width / 2 - pos.x * state.zoom;
        state.panY = containerRect.height / 2 - (pos.y + pos.height / 2) * state.zoom;
        
        applyTransform();
        updateMinimap();
        
        // Highlight the node
        const nodeElement = document.querySelector(`[data-node-id="${nodeId}"]`);
        if (nodeElement) {
            nodeElement.classList.add('highlight');
            setTimeout(() => nodeElement.classList.remove('highlight'), 2000);
        }
    }

    // =============================================================================
    // Context Menu
    // =============================================================================
    
    function showContextMenu(e, nodeId) {
        e.preventDefault();
        
        const node = state.flatNodes.get(nodeId);
        if (!node) return;
        
        selectNode(nodeId);
        
        // Update expand/collapse text
        const expandText = document.getElementById('expandCollapseText');
        const hasChildren = node.children && node.children.length > 0;
        
        if (hasChildren) {
            expandText.textContent = node.expanded === false ? 'Expand' : 'Collapse';
            expandText.parentElement.style.display = '';
        } else {
            expandText.parentElement.style.display = 'none';
        }
        
        // Position menu
        const menu = elements.contextMenu;
        menu.style.left = e.clientX + 'px';
        menu.style.top = e.clientY + 'px';
        menu.classList.add('show');
        menu.dataset.nodeId = nodeId;
        
        // Adjust if off screen
        const menuRect = menu.getBoundingClientRect();
        if (menuRect.right > window.innerWidth) {
            menu.style.left = (e.clientX - menuRect.width) + 'px';
        }
        if (menuRect.bottom > window.innerHeight) {
            menu.style.top = (e.clientY - menuRect.height) + 'px';
        }
    }
    
    function hideContextMenu() {
        elements.contextMenu.classList.remove('show');
    }
    
    function handleContextMenuAction(action) {
        const nodeId = elements.contextMenu.dataset.nodeId;
        hideContextMenu();
        
        switch (action) {
            case 'edit':
                window.ExcelSync.openEditModal(nodeId);
                break;
            case 'add-subordinate':
                window.ExcelSync.openAddModal(nodeId);
                break;
            case 'expand-collapse':
                const node = state.flatNodes.get(nodeId);
                if (node && node.children && node.children.length > 0) {
                    node.expanded = node.expanded === false ? true : false;
                    renderChart();
                }
                break;
            case 'center':
                centerOnNode(nodeId);
                break;
            case 'delete':
                window.ExcelSync.openDeleteModal(nodeId);
                break;
        }
    }

    // =============================================================================
    // Drag & Drop
    // =============================================================================
    
    function handleDragStart(e) {
        const nodeElement = e.target.closest('.org-node');
        if (!nodeElement) return;
        
        state.isDragging = true;
        state.dragNodeId = nodeElement.dataset.nodeId;
        
        nodeElement.classList.add('dragging');
        
        // Create drag preview
        const preview = elements.dragPreview;
        preview.innerHTML = nodeElement.outerHTML;
        preview.style.display = 'block';
        
        // Set drag image to transparent (we'll use our own preview)
        const dragImg = document.createElement('img');
        dragImg.src = 'data:image/gif;base64,R0lGODlhAQABAIAAAAUEBAAAACwAAAAAAQABAAACAkQBADs=';
        e.dataTransfer.setDragImage(dragImg, 0, 0);
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', state.dragNodeId);
    }
    
    function handleDrag(e) {
        if (!state.isDragging) return;
        
        const preview = elements.dragPreview;
        preview.style.left = e.clientX + 'px';
        preview.style.top = e.clientY + 'px';
    }
    
    function handleDragEnd(e) {
        state.isDragging = false;
        
        const nodeElement = document.querySelector(`[data-node-id="${state.dragNodeId}"]`);
        if (nodeElement) {
            nodeElement.classList.remove('dragging');
        }
        
        elements.dragPreview.style.display = 'none';
        state.dragNodeId = null;
        
        // Remove all drag-over classes
        document.querySelectorAll('.org-node.drag-over').forEach(el => {
            el.classList.remove('drag-over');
        });
    }
    
    function handleDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        
        const nodeElement = e.target.closest('.org-node');
        if (nodeElement && nodeElement.dataset.nodeId !== state.dragNodeId) {
            // Check if this would create a cycle
            if (!wouldCreateCycle(state.dragNodeId, nodeElement.dataset.nodeId)) {
                nodeElement.classList.add('drag-over');
            }
        }
    }
    
    function handleDragLeave(e) {
        const nodeElement = e.target.closest('.org-node');
        if (nodeElement) {
            nodeElement.classList.remove('drag-over');
        }
    }
    
    function handleDrop(e) {
        e.preventDefault();
        
        const targetNode = e.target.closest('.org-node');
        if (!targetNode) return;
        
        const targetNodeId = targetNode.dataset.nodeId;
        const sourceNodeId = e.dataTransfer.getData('text/plain');
        
        if (sourceNodeId === targetNodeId) return;
        
        // Check for cycle
        if (wouldCreateCycle(sourceNodeId, targetNodeId)) {
            window.ExcelSync.showToast('Cannot move: would create circular reference', 'error');
            return;
        }
        
        // Update the node's manager
        window.ExcelSync.updateNodeManager(sourceNodeId, targetNodeId);
        
        targetNode.classList.remove('drag-over');
    }
    
    /**
     * Check if moving sourceId under targetId would create a cycle
     */
    function wouldCreateCycle(sourceId, targetId) {
        const sourceNode = state.flatNodes.get(sourceId);
        if (!sourceNode) return false;
        
        // Check if target is a descendant of source
        function isDescendant(nodeId, ancestorId) {
            const node = state.flatNodes.get(nodeId);
            if (!node) return false;
            if (nodeId === ancestorId) return true;
            
            if (node.children) {
                for (const child of node.children) {
                    if (isDescendant(child.id, ancestorId)) return true;
                }
            }
            return false;
        }
        
        return isDescendant(sourceId, targetId);
    }

    // =============================================================================
    // Search
    // =============================================================================
    
    function handleSearch(e) {
        const query = e.target.value.toLowerCase().trim();
        
        if (!query) {
            elements.searchResults.classList.remove('show');
            elements.searchResults.innerHTML = '';
            return;
        }
        
        const results = [];
        state.flatNodes.forEach((node, id) => {
            if (node.name.toLowerCase().includes(query) ||
                node.title.toLowerCase().includes(query) ||
                node.department.toLowerCase().includes(query)) {
                results.push(node);
            }
        });
        
        if (results.length === 0) {
            elements.searchResults.innerHTML = `
                <div class="search-no-results">No employees found</div>
            `;
        } else {
            elements.searchResults.innerHTML = results.slice(0, 10).map(node => `
                <div class="search-result-item" data-node-id="${node.id}">
                    <div class="avatar" style="background-color: ${node.color || '#757575'}">
                        ${getInitials(node.name)}
                    </div>
                    <div class="info">
                        <div class="name">${escapeHtml(node.name)}</div>
                        <div class="title">${escapeHtml(node.title)} - ${escapeHtml(node.department)}</div>
                    </div>
                </div>
            `).join('');
        }
        
        elements.searchResults.classList.add('show');
    }
    
    function handleSearchResultClick(e) {
        const item = e.target.closest('.search-result-item');
        if (!item) return;
        
        const nodeId = item.dataset.nodeId;
        
        // Expand path to this node
        expandPathToNode(nodeId);
        
        // Clear search
        elements.searchInput.value = '';
        elements.searchResults.classList.remove('show');
        
        // Center and select
        selectNode(nodeId);
        setTimeout(() => centerOnNode(nodeId), 100);
    }
    
    /**
     * Expand all ancestor nodes to make the target node visible
     */
    function expandPathToNode(nodeId) {
        const node = state.flatNodes.get(nodeId);
        if (!node || !node.manager_id) return;
        
        let currentId = node.manager_id;
        let needsRerender = false;
        
        while (currentId) {
            const parent = state.flatNodes.get(currentId);
            if (parent) {
                if (parent.expanded === false) {
                    parent.expanded = true;
                    needsRerender = true;
                }
                currentId = parent.manager_id;
            } else {
                break;
            }
        }
        
        if (needsRerender) {
            renderChart();
        }
    }

    // =============================================================================
    // Layout Toggle
    // =============================================================================
    
    function setLayout(layout) {
        state.layout = layout;
        
        // Update button states
        elements.btnLayoutVertical.classList.toggle('active', layout === 'vertical');
        elements.btnLayoutHorizontal.classList.toggle('active', layout === 'horizontal');
        
        renderChart();
        zoomFit();
    }

    // =============================================================================
    // Utilities
    // =============================================================================
    
    function getInitials(name) {
        if (!name) return '?';
        const parts = name.trim().split(/\s+/);
        if (parts.length >= 2) {
            return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
        }
        return name.substring(0, 2).toUpperCase();
    }
    
    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    function updateInfoDisplay() {
        if (!state.orgData) return;
        
        elements.infoTotalEmployees.textContent = state.orgData.total_employees || 0;
        elements.infoMaxDepth.textContent = state.orgData.max_depth || 0;
    }
    
    /**
     * Build flat lookup map from tree structure
     */
    function buildFlatNodes(roots) {
        state.flatNodes.clear();
        
        function traverse(node) {
            state.flatNodes.set(node.id, node);
            if (node.children) {
                node.children.forEach(traverse);
            }
        }
        
        roots.forEach(traverse);
    }

    // =============================================================================
    // Public API
    // =============================================================================
    
    window.OrgChart = {
        /**
         * Initialize the org chart
         */
        init() {
            initElements();
            this.bindEvents();
            
            // Initial transform
            applyTransform();
            updateZoomDisplay();
        },
        
        /**
         * Set org data and render
         */
        setData(data) {
            state.orgData = data;
            buildFlatNodes(data.roots || []);
            renderChart();
            zoomFit();
        },
        
        /**
         * Refresh the chart with current data
         */
        refresh() {
            if (state.orgData) {
                buildFlatNodes(state.orgData.roots || []);
                renderChart();
            }
        },
        
        /**
         * Get a node by ID
         */
        getNode(nodeId) {
            return state.flatNodes.get(nodeId);
        },
        
        /**
         * Get all nodes as array
         */
        getAllNodes() {
            return Array.from(state.flatNodes.values());
        },
        
        /**
         * Select a node
         */
        selectNode(nodeId) {
            selectNode(nodeId);
        },
        
        /**
         * Center on a node
         */
        centerOnNode(nodeId) {
            expandPathToNode(nodeId);
            setTimeout(() => centerOnNode(nodeId), 100);
        },
        
        /**
         * Update source file display
         */
        setSourceFile(filename) {
            if (elements.infoSourceFile) {
                elements.infoSourceFile.textContent = filename || '-';
            }
        },
        
        /**
         * Get chart dimensions and info for export
         */
        getChartInfo() {
            // Calculate max depth
            let maxDepth = 0;
            const calculateDepth = (nodeId, depth) => {
                maxDepth = Math.max(maxDepth, depth);
                state.flatNodes.forEach(node => {
                    if (node.manager_id === nodeId) {
                        calculateDepth(node.id, depth + 1);
                    }
                });
            };
            state.flatNodes.forEach(node => {
                if (!node.manager_id) {
                    calculateDepth(node.id, 1);
                }
            });
            
            return {
                width: state.chartWidth,
                height: state.chartHeight,
                nodeCount: state.flatNodes.size,
                maxDepth: maxDepth
            };
        },
        
        /**
         * Export chart as image data URL using html2canvas for accurate rendering
         * @param {number} scale - Resolution multiplier (1=72dpi, 2=150dpi, 3=300dpi, 4=600dpi)
         * @param {string} bgColor - Background color ('transparent', '#ffffff', etc.)
         * @param {number} padding - Extra padding around chart in pixels
         */
        async exportAsImage(scale = 2, bgColor = '#ffffff', padding = 50) {
            return new Promise(async (resolve, reject) => {
                try {
                    const orgChart = elements.orgChart;
                    
                    // Store original styles
                    const originalTransform = orgChart.style.transform;
                    const originalWidth = orgChart.style.width;
                    const originalHeight = orgChart.style.height;
                    const containerOriginalOverflow = elements.chartContainer.style.overflow;
                    
                    // Reset transform for accurate capture
                    orgChart.style.transform = 'none';
                    
                    // Temporarily hide expand/collapse buttons for cleaner export
                    const expandBtns = orgChart.querySelectorAll('.node-expand-btn');
                    expandBtns.forEach(btn => btn.style.display = 'none');
                    
                    // Wait for any transitions to complete
                    await new Promise(r => setTimeout(r, 100));
                    
                    // Use html2canvas to capture the chart exactly as rendered
                    const canvas = await html2canvas(orgChart, {
                        scale: scale,
                        backgroundColor: bgColor === 'transparent' ? null : bgColor,
                        logging: false,
                        useCORS: true,
                        allowTaint: true,
                        // Ensure fonts are properly rendered
                        onclone: (clonedDoc) => {
                            const clonedChart = clonedDoc.getElementById('orgChart');
                            if (clonedChart) {
                                // Ensure the cloned chart has no transform
                                clonedChart.style.transform = 'none';
                                
                                // Make sure all text is fully visible (no truncation)
                                const allText = clonedChart.querySelectorAll('.node-name, .node-title, .node-phone');
                                allText.forEach(el => {
                                    el.style.overflow = 'visible';
                                    el.style.textOverflow = 'clip';
                                    el.style.whiteSpace = 'normal';
                                    el.style.wordBreak = 'break-word';
                                });
                                
                                // Hide expand buttons in clone too
                                const expandBtnsClone = clonedChart.querySelectorAll('.node-expand-btn');
                                expandBtnsClone.forEach(btn => btn.style.display = 'none');
                            }
                        }
                    });
                    
                    // Restore original styles
                    orgChart.style.transform = originalTransform;
                    orgChart.style.width = originalWidth;
                    orgChart.style.height = originalHeight;
                    elements.chartContainer.style.overflow = containerOriginalOverflow;
                    
                    // Restore expand buttons
                    expandBtns.forEach(btn => btn.style.display = '');
                    
                    // Create final canvas with padding
                    const finalCanvas = document.createElement('canvas');
                    const finalCtx = finalCanvas.getContext('2d');
                    
                    finalCanvas.width = canvas.width + (padding * 2 * scale);
                    finalCanvas.height = canvas.height + (padding * 2 * scale);
                    
                    // Fill background
                    if (bgColor !== 'transparent') {
                        finalCtx.fillStyle = bgColor;
                        finalCtx.fillRect(0, 0, finalCanvas.width, finalCanvas.height);
                    }
                    
                    // Draw the captured chart centered with padding
                    finalCtx.drawImage(canvas, padding * scale, padding * scale);
                    
                    resolve(finalCanvas.toDataURL('image/png'));
                } catch (err) {
                    console.error('Export error:', err);
                    reject(err);
                }
            });
        },
        
        /**
         * Export chart as true vector SVG string
         * Creates native SVG elements for nodes and connecting lines
         * @param {string} bgColor - Background color ('transparent', '#ffffff', etc.)
         * @param {number} padding - Extra padding around chart in pixels
         */
        async exportAsSVG(bgColor = '#ffffff', padding = 50) {
            return new Promise(async (resolve, reject) => {
                try {
                    const width = state.chartWidth + (padding * 2);
                    const height = state.chartHeight + (padding * 2);
                    const isHorizontal = state.layout === 'horizontal';
                    const borderRadius = state.cardCornerStyle === 'curved' ? 12 : 0;
                    const innerRadius = state.cardCornerStyle === 'curved' ? 10 : 0;
                    
                    // Helper to escape HTML entities
                    const escapeXml = (str) => {
                        if (!str) return '';
                        return String(str)
                            .replace(/&/g, '&amp;')
                            .replace(/</g, '&lt;')
                            .replace(/>/g, '&gt;')
                            .replace(/"/g, '&quot;')
                            .replace(/'/g, '&apos;');
                    };
                    
                    // Helper to get initials
                    const getInitialsForSvg = (name) => {
                        if (!name) return '?';
                        const words = name.trim().split(/\s+/);
                        if (words.length === 1) {
                            return words[0].substring(0, 2).toUpperCase();
                        }
                        return (words[0][0] + words[words.length - 1][0]).toUpperCase();
                    };
                    
                    // Build connecting lines
                    let linesContent = '';
                    state.nodePositions.forEach((pos, nodeId) => {
                        const node = pos.node;
                        if (!node.children || node.children.length === 0 || node.expanded === false) return;
                        
                        node.children.forEach(child => {
                            const childPos = state.nodePositions.get(child.id);
                            if (!childPos) return;
                            
                            let d;
                            const cornerRadius = 12;
                            
                            if (isHorizontal) {
                                const startX = pos.x + CONFIG.nodeWidth / 2;
                                const startY = pos.y + pos.height / 2;
                                const endX = childPos.x - CONFIG.nodeWidth / 2;
                                const endY = childPos.y + childPos.height / 2;
                                const midX = (startX + endX) / 2;
                                
                                if (state.lineCornerStyle === 'curved') {
                                    d = `M ${startX + padding} ${startY + padding} C ${midX + padding} ${startY + padding}, ${midX + padding} ${endY + padding}, ${endX + padding} ${endY + padding}`;
                                } else {
                                    d = `M ${startX + padding} ${startY + padding} L ${midX + padding} ${startY + padding} L ${midX + padding} ${endY + padding} L ${endX + padding} ${endY + padding}`;
                                }
                            } else {
                                const startX = pos.x;
                                const startY = pos.y + pos.height;
                                const endX = childPos.x;
                                const endY = childPos.y;
                                const midY = startY + (endY - startY) / 2;
                                
                                if (state.lineCornerStyle === 'curved') {
                                    const r = Math.min(cornerRadius, Math.abs(endX - startX) / 2, (endY - startY) / 4);
                                    
                                    if (startX === endX) {
                                        d = `M ${startX + padding} ${startY + padding} L ${endX + padding} ${endY + padding}`;
                                    } else if (endX > startX) {
                                        d = `M ${startX + padding} ${startY + padding} L ${startX + padding} ${midY - r + padding} Q ${startX + padding} ${midY + padding} ${startX + r + padding} ${midY + padding} L ${endX - r + padding} ${midY + padding} Q ${endX + padding} ${midY + padding} ${endX + padding} ${midY + r + padding} L ${endX + padding} ${endY + padding}`;
                                    } else {
                                        d = `M ${startX + padding} ${startY + padding} L ${startX + padding} ${midY - r + padding} Q ${startX + padding} ${midY + padding} ${startX - r + padding} ${midY + padding} L ${endX + r + padding} ${midY + padding} Q ${endX + padding} ${midY + padding} ${endX + padding} ${midY + r + padding} L ${endX + padding} ${endY + padding}`;
                                    }
                                } else {
                                    d = `M ${startX + padding} ${startY + padding} L ${startX + padding} ${midY + padding} L ${endX + padding} ${midY + padding} L ${endX + padding} ${endY + padding}`;
                                }
                            }
                            
                            linesContent += `<path d="${d}" fill="none" stroke="${state.borderLineColor}" stroke-width="2"/>\n`;
                        });
                    });
                    
                    // Helper function to wrap text into multiple lines
                    const wrapText = (text, maxWidth, fontSize) => {
                        if (!text) return [];
                        const words = text.split(' ');
                        const lines = [];
                        let currentLine = '';
                        
                        // Approximate character width (varies by font, this is an estimate)
                        const avgCharWidth = fontSize * 0.55;
                        const maxChars = Math.floor(maxWidth / avgCharWidth);
                        
                        words.forEach(word => {
                            const testLine = currentLine ? currentLine + ' ' + word : word;
                            if (testLine.length <= maxChars) {
                                currentLine = testLine;
                            } else {
                                if (currentLine) {
                                    lines.push(currentLine);
                                }
                                // If single word is too long, split it
                                if (word.length > maxChars) {
                                    while (word.length > maxChars) {
                                        lines.push(word.substring(0, maxChars - 1) + '-');
                                        word = word.substring(maxChars - 1);
                                    }
                                    currentLine = word;
                                } else {
                                    currentLine = word;
                                }
                            }
                        });
                        if (currentLine) {
                            lines.push(currentLine);
                        }
                        return lines;
                    };
                    
                    // Build node elements
                    let nodesContent = '';
                    state.nodePositions.forEach((pos, nodeId) => {
                        const node = pos.node;
                        const x = pos.x - CONFIG.nodeWidth / 2 + padding;
                        const y = pos.y + padding;
                        const nodeWidth = CONFIG.nodeWidth;
                        const nodeHeight = pos.height;
                        const headerHeight = 24;
                        const avatarSize = 36;
                        const avatarBg = node.color || '#757575';
                        const initials = getInitialsForSvg(node.name);
                        const textPadding = 8; // Padding on each side for text
                        const textMaxWidth = nodeWidth - (textPadding * 2);
                        
                        // Node group
                        nodesContent += `<g transform="translate(${x}, ${y})">\n`;
                        
                        // Drop shadow filter reference
                        nodesContent += `  <!-- Node: ${escapeXml(node.name)} -->\n`;
                        
                        // Card background (white body with border)
                        nodesContent += `  <rect x="0" y="0" width="${nodeWidth}" height="${nodeHeight}" rx="${borderRadius}" ry="${borderRadius}" fill="white" stroke="${state.borderLineColor}" stroke-width="2"/>\n`;
                        
                        // Header background (colored)
                        if (borderRadius > 0) {
                            nodesContent += `  <path d="M ${borderRadius} 0 L ${nodeWidth - borderRadius} 0 Q ${nodeWidth} 0 ${nodeWidth} ${borderRadius} L ${nodeWidth} ${headerHeight} L 0 ${headerHeight} L 0 ${borderRadius} Q 0 0 ${borderRadius} 0 Z" fill="${node.color || '#757575'}"/>\n`;
                        } else {
                            nodesContent += `  <rect x="0" y="0" width="${nodeWidth}" height="${headerHeight}" fill="${node.color || '#757575'}"/>\n`;
                        }
                        
                        // Avatar circle
                        const avatarCx = nodeWidth / 2;
                        const avatarCy = headerHeight;
                        nodesContent += `  <circle cx="${avatarCx}" cy="${avatarCy}" r="${avatarSize / 2 + 2}" fill="white"/>\n`;
                        nodesContent += `  <circle cx="${avatarCx}" cy="${avatarCy}" r="${avatarSize / 2}" fill="${avatarBg}"/>\n`;
                        
                        // Avatar initials
                        nodesContent += `  <text x="${avatarCx}" y="${avatarCy + 4}" text-anchor="middle" fill="white" font-family="Inter, -apple-system, BlinkMacSystemFont, sans-serif" font-size="12" font-weight="600">${escapeXml(initials)}</text>\n`;
                        
                        // Name text with wrapping
                        const nameFontSize = 11;
                        const nameLineHeight = 13;
                        const nameLines = wrapText(node.name || '', textMaxWidth, nameFontSize);
                        let nameY = headerHeight + avatarSize / 2 + 16;
                        
                        nodesContent += `  <text x="${nodeWidth / 2}" y="${nameY}" text-anchor="middle" fill="#424242" font-family="Inter, -apple-system, BlinkMacSystemFont, sans-serif" font-size="${nameFontSize}" font-weight="600">\n`;
                        nameLines.forEach((line, i) => {
                            if (i === 0) {
                                nodesContent += `    <tspan x="${nodeWidth / 2}" dy="0">${escapeXml(line)}</tspan>\n`;
                            } else {
                                nodesContent += `    <tspan x="${nodeWidth / 2}" dy="${nameLineHeight}">${escapeXml(line)}</tspan>\n`;
                            }
                        });
                        nodesContent += `  </text>\n`;
                        
                        // Title text with wrapping
                        const titleFontSize = 10;
                        const titleLineHeight = 12;
                        const titleLines = wrapText(node.title || '', textMaxWidth, titleFontSize);
                        const titleY = nameY + (nameLines.length - 1) * nameLineHeight + 14;
                        
                        nodesContent += `  <text x="${nodeWidth / 2}" y="${titleY}" text-anchor="middle" fill="#9e9e9e" font-family="Inter, -apple-system, BlinkMacSystemFont, sans-serif" font-size="${titleFontSize}">\n`;
                        titleLines.forEach((line, i) => {
                            if (i === 0) {
                                nodesContent += `    <tspan x="${nodeWidth / 2}" dy="0">${escapeXml(line)}</tspan>\n`;
                            } else {
                                nodesContent += `    <tspan x="${nodeWidth / 2}" dy="${titleLineHeight}">${escapeXml(line)}</tspan>\n`;
                            }
                        });
                        nodesContent += `  </text>\n`;
                        
                        // WhatsApp number (if shown)
                        if (node.whatsapp && state.showWhatsApp) {
                            const phoneY = titleY + (titleLines.length - 1) * titleLineHeight + 14;
                            const cleanNumber = node.whatsapp.replace(/[^0-9]/g, '');
                            const displayNumber = '+' + cleanNumber;
                            nodesContent += `  <text x="${nodeWidth / 2}" y="${phoneY}" text-anchor="middle" fill="#128C7E" font-family="Inter, -apple-system, BlinkMacSystemFont, sans-serif" font-size="10">${escapeXml(displayNumber)}</text>\n`;
                        }
                        
                        nodesContent += `</g>\n`;
                    });
                    
                    // Create the SVG content
                    const svgContent = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" 
     width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
    <title>Organization Chart</title>
    <desc>Exported from Organization Chart Maker</desc>
    
    <defs>
        <style>
            text { user-select: none; }
        </style>
    </defs>
    
    ${bgColor !== 'transparent' ? `<rect width="100%" height="100%" fill="${bgColor}"/>` : ''}
    
    <!-- Connecting Lines -->
    <g id="lines">
${linesContent}
    </g>
    
    <!-- Nodes -->
    <g id="nodes">
${nodesContent}
    </g>
</svg>`;
                    
                    resolve(svgContent);
                } catch (err) {
                    console.error('SVG Export error:', err);
                    reject(err);
                }
            });
        },
        
        /**
         * Export complete organization data as JSON
         * @returns {Object} Complete organization data with metadata
         */
        async exportAsJSON() {
            const now = new Date();
            
            // Build hierarchy recursively
            const buildHierarchy = (node) => {
                const children = [];
                state.flatNodes.forEach(child => {
                    if (child.manager_id === node.id) {
                        children.push(buildHierarchy(child));
                    }
                });
                
                return {
                    id: node.id,
                    name: node.name,
                    title: node.title,
                    department: node.department,
                    color: node.color || '#1a73e8',
                    avatar_url: node.avatar_url || null,
                    whatsapp: node.whatsapp || null,
                    manager_id: node.manager_id || null,
                    children: children.length > 0 ? children : undefined
                };
            };
            
            // Find root nodes and build hierarchy
            const rootNodes = [];
            const allNodes = [];
            
            state.flatNodes.forEach(node => {
                // Add to flat list
                allNodes.push({
                    id: node.id,
                    name: node.name,
                    title: node.title,
                    department: node.department,
                    color: node.color || '#1a73e8',
                    avatar_url: node.avatar_url || null,
                    whatsapp: node.whatsapp || null,
                    manager_id: node.manager_id || null
                });
                
                // Check if root
                if (!node.manager_id) {
                    rootNodes.push(buildHierarchy(node));
                }
            });
            
            // Get unique departments
            const departments = [...new Set(allNodes.map(n => n.department).filter(Boolean))];
            
            // Calculate statistics
            const maxDepth = this.getChartInfo().maxDepth;
            
            return {
                exportInfo: {
                    version: '1.0',
                    exportDate: now.toISOString(),
                    exportDateLocal: now.toLocaleString(),
                    generator: 'Organization Chart Maker'
                },
                statistics: {
                    totalEmployees: allNodes.length,
                    departments: departments.length,
                    maxDepth: maxDepth
                },
                departments: departments,
                nodes: allNodes,
                hierarchy: rootNodes.length === 1 ? rootNodes[0] : rootNodes
            };
        },
        
        /**
         * Export chart as SVG string (legacy - uses foreignObject)
         * @deprecated Use exportAsSVG instead
         */
        async exportAsSVGLegacy(bgColor = '#ffffff', padding = 50) {
            return new Promise(async (resolve, reject) => {
                try {
                    const orgChart = elements.orgChart;
                    
                    // Store original styles
                    const originalTransform = orgChart.style.transform;
                    
                    // Reset transform for accurate capture
                    orgChart.style.transform = 'none';
                    
                    // Temporarily hide expand/collapse buttons for cleaner export
                    const expandBtns = orgChart.querySelectorAll('.node-expand-btn');
                    expandBtns.forEach(btn => btn.style.display = 'none');
                    
                    // Wait for any transitions to complete
                    await new Promise(r => setTimeout(r, 100));
                    
                    // Get the chart dimensions
                    const chartRect = orgChart.getBoundingClientRect();
                    const width = state.chartWidth + (padding * 2);
                    const height = state.chartHeight + (padding * 2);
                    
                    // Clone the org chart for manipulation
                    const clonedChart = orgChart.cloneNode(true);
                    
                    // Remove expand buttons from clone
                    clonedChart.querySelectorAll('.node-expand-btn').forEach(btn => btn.remove());
                    
                    // Get all computed styles and inline them
                    const inlineStyles = (element, clone) => {
                        const computed = window.getComputedStyle(element);
                        const importantStyles = [
                            'font-family', 'font-size', 'font-weight', 'color', 'background-color',
                            'background', 'border', 'border-radius', 'padding', 'margin',
                            'display', 'flex-direction', 'align-items', 'justify-content', 'gap',
                            'position', 'top', 'left', 'right', 'bottom', 'width', 'height',
                            'min-width', 'min-height', 'max-width', 'max-height',
                            'box-shadow', 'text-align', 'line-height', 'letter-spacing',
                            'overflow', 'white-space', 'text-overflow', 'word-wrap',
                            'transform', 'opacity', 'z-index', 'fill', 'stroke', 'stroke-width'
                        ];
                        
                        importantStyles.forEach(prop => {
                            const value = computed.getPropertyValue(prop);
                            if (value) {
                                clone.style.setProperty(prop, value);
                            }
                        });
                        
                        // Handle CSS variables
                        clone.style.setProperty('--border-line-color', state.borderLineColor);
                        
                        // Recursively inline children
                        const children = element.children;
                        const cloneChildren = clone.children;
                        for (let i = 0; i < children.length; i++) {
                            if (children[i] && cloneChildren[i]) {
                                inlineStyles(children[i], cloneChildren[i]);
                            }
                        }
                    };
                    
                    inlineStyles(orgChart, clonedChart);
                    
                    // Create SVG wrapper
                    const svgNS = 'http://www.w3.org/2000/svg';
                    const svg = document.createElementNS(svgNS, 'svg');
                    svg.setAttribute('xmlns', svgNS);
                    svg.setAttribute('width', width);
                    svg.setAttribute('height', height);
                    svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
                    
                    // Add background rect if not transparent
                    if (bgColor !== 'transparent') {
                        const bgRect = document.createElementNS(svgNS, 'rect');
                        bgRect.setAttribute('width', '100%');
                        bgRect.setAttribute('height', '100%');
                        bgRect.setAttribute('fill', bgColor);
                        svg.appendChild(bgRect);
                    }
                    
                    // Create foreignObject to embed HTML
                    const foreignObject = document.createElementNS(svgNS, 'foreignObject');
                    foreignObject.setAttribute('x', padding);
                    foreignObject.setAttribute('y', padding);
                    foreignObject.setAttribute('width', state.chartWidth);
                    foreignObject.setAttribute('height', state.chartHeight);
                    
                    // Set up the cloned chart
                    clonedChart.style.transform = 'none';
                    clonedChart.style.position = 'relative';
                    clonedChart.style.width = state.chartWidth + 'px';
                    clonedChart.style.height = state.chartHeight + 'px';
                    clonedChart.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
                    
                    foreignObject.appendChild(clonedChart);
                    svg.appendChild(foreignObject);
                    
                    // Serialize to string
                    const serializer = new XMLSerializer();
                    let svgString = serializer.serializeToString(svg);
                    
                    // Clean up and add XML declaration
                    svgString = '<?xml version="1.0" encoding="UTF-8"?>\n' + svgString;
                    
                    // Restore original styles
                    orgChart.style.transform = originalTransform;
                    
                    // Restore expand buttons
                    expandBtns.forEach(btn => btn.style.display = '');
                    
                    resolve(svgString);
                } catch (err) {
                    console.error('SVG Export error:', err);
                    reject(err);
                }
            });
        },
        
        /**
         * Bind all event handlers
         */
        bindEvents() {
            // Click handlers
            elements.chartContainer.addEventListener('click', handleNodeClick);
            elements.chartContainer.addEventListener('dblclick', handleNodeDoubleClick);
            elements.chartContainer.addEventListener('click', handleExpandCollapse);
            
            // Context menu
            elements.chartContainer.addEventListener('contextmenu', (e) => {
                const nodeElement = e.target.closest('.org-node');
                if (nodeElement) {
                    showContextMenu(e, nodeElement.dataset.nodeId);
                }
            });
            
            document.addEventListener('click', (e) => {
                if (!e.target.closest('.context-menu')) {
                    hideContextMenu();
                }
            });
            
            elements.contextMenu.addEventListener('click', (e) => {
                const item = e.target.closest('.context-menu-item');
                if (item) {
                    handleContextMenuAction(item.dataset.action);
                }
            });
            
            // Pan handlers
            elements.chartContainer.addEventListener('mousedown', startPan);
            document.addEventListener('mousemove', doPan);
            document.addEventListener('mouseup', endPan);
            
            // Zoom
            elements.chartContainer.addEventListener('wheel', handleWheel, { passive: false });
            elements.btnZoomIn.addEventListener('click', zoomIn);
            elements.btnZoomOut.addEventListener('click', zoomOut);
            elements.btnZoomFit.addEventListener('click', zoomFit);
            elements.btnZoomReset.addEventListener('click', zoomReset);
            
            // Layout toggle
            elements.btnLayoutVertical.addEventListener('click', () => setLayout('vertical'));
            elements.btnLayoutHorizontal.addEventListener('click', () => setLayout('horizontal'));
            
            // Search
            elements.searchInput.addEventListener('input', handleSearch);
            elements.searchInput.addEventListener('focus', handleSearch);
            elements.searchResults.addEventListener('click', handleSearchResultClick);
            
            document.addEventListener('click', (e) => {
                if (!e.target.closest('.search-container')) {
                    elements.searchResults.classList.remove('show');
                }
            });
            
            // Drag & Drop
            elements.chartContainer.addEventListener('dragstart', handleDragStart);
            document.addEventListener('drag', handleDrag);
            document.addEventListener('dragend', handleDragEnd);
            elements.chartContainer.addEventListener('dragover', handleDragOver);
            elements.chartContainer.addEventListener('dragleave', handleDragLeave);
            elements.chartContainer.addEventListener('drop', handleDrop);
            
            // Keyboard shortcuts
            document.addEventListener('keydown', (e) => {
                if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
                
                if (e.key === 'Escape') {
                    hideContextMenu();
                    state.selectedNodeId = null;
                    document.querySelectorAll('.org-node.selected').forEach(el => {
                        el.classList.remove('selected');
                    });
                }
                
                if (e.key === '+' || e.key === '=') zoomIn();
                if (e.key === '-') zoomOut();
                if (e.key === '0') zoomReset();
                if (e.key === 'f' && e.ctrlKey) {
                    e.preventDefault();
                    elements.searchInput.focus();
                }
            });
            
            // Window resize
            window.addEventListener('resize', () => {
                updateMinimap();
            });
            
            // Settings handlers
            const toggleWhatsApp = document.getElementById('toggleWhatsApp');
            const whatsappToggleText = document.getElementById('whatsappToggleText');
            
            function updateWhatsAppToggleText(isChecked) {
                if (whatsappToggleText) {
                    whatsappToggleText.textContent = isChecked ? 'Hide WhatsApp Numbers' : 'Show WhatsApp Numbers';
                }
            }
            
            if (toggleWhatsApp) {
                // Initialize text based on current state
                updateWhatsAppToggleText(toggleWhatsApp.checked);
                
                toggleWhatsApp.addEventListener('change', (e) => {
                    state.showWhatsApp = e.target.checked;
                    updateWhatsAppToggleText(e.target.checked);
                    if (state.orgData) renderChart();
                });
            }
            
            const closenessSlider = document.getElementById('closenessSlider');
            if (closenessSlider) {
                closenessSlider.addEventListener('input', (e) => {
                    state.closeness = parseFloat(e.target.value);
                    if (state.orgData) {
                        renderChart();
                        zoomFit();
                    }
                });
            }
            
            // Border and Line Color picker
            const borderLineColor = document.getElementById('borderLineColor');
            const borderLineColorHex = document.getElementById('borderLineColorHex');
            
            if (borderLineColor) {
                borderLineColor.addEventListener('input', (e) => {
                    const color = e.target.value;
                    state.borderLineColor = color;
                    if (borderLineColorHex) borderLineColorHex.value = color;
                    applyBorderLineColor(color);
                });
            }
            
            if (borderLineColorHex) {
                borderLineColorHex.addEventListener('input', (e) => {
                    let color = e.target.value;
                    // Add # if missing
                    if (color && !color.startsWith('#')) {
                        color = '#' + color;
                    }
                    // Validate hex color
                    if (/^#[0-9A-Fa-f]{6}$/.test(color)) {
                        state.borderLineColor = color;
                        if (borderLineColor) borderLineColor.value = color;
                        applyBorderLineColor(color);
                    }
                });
                
                borderLineColorHex.addEventListener('blur', (e) => {
                    // Reset to current valid color on blur if invalid
                    e.target.value = state.borderLineColor;
                });
            }
            
            // Line Corner Style toggle
            const btnLineHard = document.getElementById('btnLineHard');
            const btnLineCurved = document.getElementById('btnLineCurved');
            
            if (btnLineHard) {
                btnLineHard.addEventListener('click', () => {
                    state.lineCornerStyle = 'hard';
                    btnLineHard.classList.add('active');
                    if (btnLineCurved) btnLineCurved.classList.remove('active');
                    if (state.orgData) renderChart();
                });
            }
            
            if (btnLineCurved) {
                btnLineCurved.addEventListener('click', () => {
                    state.lineCornerStyle = 'curved';
                    btnLineCurved.classList.add('active');
                    if (btnLineHard) btnLineHard.classList.remove('active');
                    if (state.orgData) renderChart();
                });
            }
            
            // Card Corner Style toggle
            const btnCardHard = document.getElementById('btnCardHard');
            const btnCardCurved = document.getElementById('btnCardCurved');
            
            if (btnCardHard) {
                btnCardHard.addEventListener('click', () => {
                    state.cardCornerStyle = 'hard';
                    btnCardHard.classList.add('active');
                    if (btnCardCurved) btnCardCurved.classList.remove('active');
                    applyCardCornerStyle('hard');
                });
            }
            
            if (btnCardCurved) {
                btnCardCurved.addEventListener('click', () => {
                    state.cardCornerStyle = 'curved';
                    btnCardCurved.classList.add('active');
                    if (btnCardHard) btnCardHard.classList.remove('active');
                    applyCardCornerStyle('curved');
                });
            }
        },
        
        /**
         * Toggle WhatsApp visibility
         */
        setShowWhatsApp(show) {
            state.showWhatsApp = show;
            const toggle = document.getElementById('toggleWhatsApp');
            const whatsappToggleText = document.getElementById('whatsappToggleText');
            if (toggle) toggle.checked = show;
            if (whatsappToggleText) {
                whatsappToggleText.textContent = show ? 'Hide WhatsApp Numbers' : 'Show WhatsApp Numbers';
            }
            if (state.orgData) renderChart();
        },
        
        /**
         * Set chart closeness
         */
        setCloseness(value) {
            state.closeness = Math.max(0.5, Math.min(2, value));
            const slider = document.getElementById('closenessSlider');
            if (slider) slider.value = state.closeness;
            if (state.orgData) {
                renderChart();
                zoomFit();
            }
        },
        
        /**
         * Set border and line color
         */
        setBorderLineColor(color) {
            state.borderLineColor = color;
            const colorPicker = document.getElementById('borderLineColor');
            const colorHex = document.getElementById('borderLineColorHex');
            if (colorPicker) colorPicker.value = color;
            if (colorHex) colorHex.value = color;
            applyBorderLineColor(color);
        },
        
        /**
         * Get current border and line color
         */
        getBorderLineColor() {
            return state.borderLineColor;
        }
    };
    
    /**
     * Apply border and line color to the chart
     */
    function applyBorderLineColor(color) {
        // Update CSS variable
        document.documentElement.style.setProperty('--border-line-color', color);
    }
    
    /**
     * Apply card corner style (hard or curved)
     */
    function applyCardCornerStyle(style) {
        if (style === 'hard') {
            document.documentElement.style.setProperty('--card-radius', '4px');
            document.documentElement.style.setProperty('--card-radius-inner', '2px');
        } else {
            document.documentElement.style.setProperty('--card-radius', '12px');
            document.documentElement.style.setProperty('--card-radius-inner', '10px');
        }
    }
    
    // Canvas helper functions
    function roundRect(ctx, x, y, width, height, radius) {
        ctx.beginPath();
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + width - radius, y);
        ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
        ctx.lineTo(x + width, y + height - radius);
        ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
        ctx.lineTo(x + radius, y + height);
        ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
        ctx.lineTo(x, y + radius);
        ctx.quadraticCurveTo(x, y, x + radius, y);
        ctx.closePath();
    }
    
    function roundRectTop(ctx, x, y, width, height, radius) {
        ctx.beginPath();
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + width - radius, y);
        ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
        ctx.lineTo(x + width, y + height);
        ctx.lineTo(x, y + height);
        ctx.lineTo(x, y + radius);
        ctx.quadraticCurveTo(x, y, x + radius, y);
        ctx.closePath();
    }
    
    function truncateText(ctx, text, maxWidth) {
        if (!text) return '';
        if (ctx.measureText(text).width <= maxWidth) return text;
        
        let truncated = text;
        while (ctx.measureText(truncated + '...').width > maxWidth && truncated.length > 0) {
            truncated = truncated.slice(0, -1);
        }
        return truncated + '...';
    }
    
})();
