        // Global state
        let currentProject = null;
        let projects = safeParseJSON(localStorage.getItem('bytedraft_projects'), []);
        let customFields = safeParseJSON(localStorage.getItem('bytedraft_fields'), []);
        let versionHistory = safeParseJSON(localStorage.getItem('bytedraft_versions'), []);
        let customChangelog = safeParseJSON(localStorage.getItem('bytedraft_custom_changelog'), {});
        // Migration: un-double-encode changelog entries saved by older versions
        Object.keys(customChangelog).forEach(key => {
            if (typeof customChangelog[key] === 'string') {
                try { customChangelog[key] = JSON.parse(customChangelog[key]); }
                catch(e) { customChangelog[key] = []; }
            }
        });

        let _storageWarned = false;  // warn once per session when storage approaches 80%

        function safeParseJSON(value, fallback) {
            try { return value ? JSON.parse(value) : fallback; }
            catch (e) { return fallback; }
        }

        function escapeHtml(str) {
            if (!str) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
        }

        function showToast(message, type = 'info') {
            const colorMap = { success: 'alert-success', warning: 'alert-warning', error: 'alert-danger', info: 'alert-info' };
            const toast = document.createElement('div');
            toast.className = `alert ${colorMap[type] || 'alert-info'} position-fixed shadow`;
            toast.style.cssText = 'top: 20px; left: 50%; transform: translateX(-50%); z-index: 9999; min-width: 300px; text-align: center;';
            toast.textContent = message;
            document.body.appendChild(toast);
            setTimeout(() => toast.remove(), 3000);
        }

        function showConfirm(message, onConfirm, options) {
            const title = (options && options.title) || 'Confirm';
            const okLabel = (options && options.okLabel) || 'Confirm';
            const okClass = (options && options.okClass) || 'btn-danger';
            document.getElementById('confirmModalTitle').textContent = title;
            document.getElementById('confirmModalMessage').textContent = message;
            const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('confirmModal'));
            // Clone the button to atomically remove all prior click listeners
            const oldBtn = document.getElementById('confirmModalOkBtn');
            const okBtn = oldBtn.cloneNode(true);
            oldBtn.replaceWith(okBtn);
            okBtn.textContent = okLabel;
            okBtn.className = 'btn ' + okClass;
            okBtn.addEventListener('click', () => { modal.hide(); onConfirm(); }, { once: true });
            modal.show();
        }

        function getLocalStorageSize() {
            let total = 0;
            for (let i = 0; i < localStorage.length; i++) {
                const key = localStorage.key(i);
                total += (key.length + (localStorage.getItem(key) || '').length) * 2;
            }
            return total;
        }

        function safeSetItem(key, value) {
            if (!_storageWarned && getLocalStorageSize() >= 4194304) {
                _storageWarned = true;
                showToast('Storage is nearly full (>80%). Consider removing logos or exporting old projects.', 'warning');
            }
            try {
                localStorage.setItem(key, value);
            } catch (e) {
                if (e.name === 'QuotaExceededError' || e.name === 'NS_ERROR_DOM_QUOTA_REACHED') {
                    showToast('Storage limit reached. Remove large images or export old projects.', 'error');
                } else {
                    showToast('Storage error. Data may not have been saved.', 'error');
                }
            }
        }



        // Templates will be loaded from templates.json
        let templates = {};

        // Initialize the application
        // On DOMContentLoaded, fetch templates.json and then initialize

        document.addEventListener('DOMContentLoaded', function() {
            // Check if libraries are loaded

            // Use templates from templates.js
            templates = window.templates || {};
            renderProjects();
            renderTemplates();
            renderCustomFields();
            initSidebarCollapse();

            // Dispose Bootstrap modal instances after close to free their event listeners.
            // getOrCreateInstance() re-creates transparently on next open.
            document.addEventListener('hidden.bs.modal', function(e) {
                bootstrap.Modal.getInstance(e.target)?.dispose();
            });
        });

        document.addEventListener('keydown', function(e) {
            if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey) {
                if (e.key === 's') {
                    e.preventDefault();
                    if (currentProject) { updateAllSectionContents(); saveProject(); }
                } else if (e.key === 'f') {
                    e.preventDefault();
                    if (currentProject) showFindReplaceModal();
                }
            }
        });

        function toggleSidebarSection(key) {
            const body = document.getElementById(`sidebar-${key}-body`);
            const chevron = document.getElementById(`chevron-${key}`);
            if (!body) return;
            const isCollapsed = body.style.display === 'none';
            body.style.display = isCollapsed ? '' : 'none';
            if (chevron) chevron.style.transform = isCollapsed ? 'rotate(0deg)' : 'rotate(-90deg)';
            const states = safeParseJSON(localStorage.getItem('bytedraft_sidebar_collapse'), {});
            states[key] = !isCollapsed;
            safeSetItem('bytedraft_sidebar_collapse', JSON.stringify(states));
        }

        function initSidebarCollapse() {
            const states = safeParseJSON(localStorage.getItem('bytedraft_sidebar_collapse'), {});
            ['projects', 'toc', 'templates', 'fields'].forEach(key => {
                if (states[key] === true) {
                    const body = document.getElementById(`sidebar-${key}-body`);
                    const chevron = document.getElementById(`chevron-${key}`);
                    if (body) body.style.display = 'none';
                    if (chevron) chevron.style.transform = 'rotate(-90deg)';
                }
            });
        }

        // Project Management
        function createProject() {
            const name = document.getElementById('newProjectName').value;
            const description = document.getElementById('newProjectDesc').value;
            const templateKey = document.getElementById('newProjectTemplate').value;

            if (!name) {
                showToast('Please enter a project name', 'warning');
                return;
            }

            let templateSections = [];
            if (templateKey) {
                if (templates && templates[templateKey] && Array.isArray(templates[templateKey].sections)) {
                    templateSections = [...templates[templateKey].sections];
                } else {
                    showToast('Template not found or invalid. Please check templates.js and the template key.', 'error');
                    return;
                }
            }

            const project = {
                id: Date.now().toString(),
                name: name,
                description: description,
                status: 'draft',
                createdAt: new Date().toISOString(),
                updatedAt: new Date().toISOString(),
                sections: templateSections.length > 0 ? templateSections : [
                    { id: '1', title: 'Introduction', content: '' }
                ]
            };

            projects.push(project);
            saveProjects();
            renderProjects();

            // Close modal and select new project
            bootstrap.Modal.getInstance(document.getElementById('newProjectModal')).hide();
            selectProject(project.id);
        }

        function selectProject(projectId) {
            currentProject = projects.find(p => p.id === projectId);
            document.getElementById('currentProjectTitle').textContent = currentProject.name;
            renderProjectContent();
            // Update TOC when project is selected
            setTimeout(() => updateTOCPreview(), 200);

        }

        function renderProjects() {
            const container = document.getElementById('projectsList');
            container.innerHTML = '';
            
            const statusFilter = document.getElementById('statusFilter')?.value || 'all';
            const filteredProjects = statusFilter === 'all' ? projects : projects.filter(p => p.status === statusFilter);
            
            if (filteredProjects.length === 0) {
                container.innerHTML = `
                    <div class="text-center py-3 text-muted">
                        <i class="fas fa-folder-open fa-2x mb-2"></i>
                        <p class="mb-0">No projects found</p>
                    </div>
                `;
                return;
            }
            
            filteredProjects.forEach(project => {
                const card = document.createElement('div');
                card.className = `project-card ${currentProject?.id === project.id ? 'active' : ''}`;
                
                card.innerHTML = `
                    <h6 class="mb-1" style="cursor: pointer;" onclick="selectProject('${project.id}')">${escapeHtml(project.name)}</h6>
                    <div class="d-flex align-items-center gap-2 mt-1">
                        <button class="btn btn-sm btn-outline-secondary" onclick="exportProjectAsJSON('${project.id}'); event.stopPropagation();" title="Export as JSON">
                            <i class="fas fa-file-code"></i>
                        </button>
                        <select class="status-selector" onchange="updateProjectStatus('${project.id}', this.value)" onclick="event.stopPropagation()">
                            <option value="draft" ${project.status === 'draft' ? 'selected' : ''}>Draft</option>
                            <option value="working" ${project.status === 'working' ? 'selected' : ''}>Working</option>
                            <option value="publish" ${project.status === 'publish' ? 'selected' : ''}>Publish</option>
                        </select>
                        <button class="btn-delete" onclick="deleteProject('${project.id}'); event.stopPropagation();" title="Delete Project">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                    <small class="text-muted d-block mt-1" style="cursor: pointer;" onclick="selectProject('${project.id}')">${escapeHtml(project.description) || 'No description'}</small>
                    <div class="d-flex justify-content-between align-items-center mt-1">
                        <small class="text-muted">Updated: ${new Date(project.updatedAt).toLocaleDateString()}</small>
                        <span class="status-badge status-${project.status}">${project.status}</span>
                    </div>
                `;
                
                container.appendChild(card);
            });
        }

        function filterProjects() {
            renderProjects();
        }

        function renderProjectContent() {
            const container = document.getElementById('contentArea');
            
            if (!currentProject) {
                // Show default state when no project is selected
                container.innerHTML = `
                    <div class="text-center py-5">
                        <i class="fas fa-file-alt fa-3x text-muted mb-3"></i>
                        <h5 class="text-muted">Select a project to get started</h5>
                        <p class="text-muted">Or create a new project to begin documenting</p>
                    </div>
                `;
                return;
            }
            
            // Show project content when a project is selected
            container.innerHTML = `
                <div class="d-flex justify-content-between align-items-center mb-4">
                    <div>
                        <h5 class="mb-1">${escapeHtml(currentProject.name)}</h5>
                        <p class="text-muted mb-0">${escapeHtml(currentProject.description) || 'No description'}</p>
                        <div class="mt-2">
                            <span class="status-badge status-${currentProject.status}">${currentProject.status}</span>
                            <small class="text-muted ms-2">Last updated: ${new Date(currentProject.updatedAt).toLocaleString()}</small>
                        </div>
                    </div>
                    <div class="d-flex gap-2">
                        <select class="form-select form-select-sm" onchange="updateProjectStatus('${currentProject.id}', this.value)" style="width: auto;">
                            <option value="draft" ${currentProject.status === 'draft' ? 'selected' : ''}>Draft</option>
                            <option value="working" ${currentProject.status === 'working' ? 'selected' : ''}>Working</option>
                            <option value="publish" ${currentProject.status === 'publish' ? 'selected' : ''}>Publish</option>
                        </select>
                        <button class="btn btn-outline-primary" onclick="addSection()">
                            <i class="fas fa-plus me-1"></i>Add Section
                        </button>
                        <button class="btn btn-outline-secondary" onclick="saveProject()">
                            <i class="fas fa-save me-1"></i>Save
                        </button>
                    </div>
                </div>
                <div id="sectionsContainer"></div>
            `;
            
            renderSections();
        }

        function renderSections() {
            updateAllSectionContents(); // Sync editors before destroying them
            tinymce.remove(); // Clean up all TinyMCE editors before re-rendering
            if (!currentProject) return;
            const container = document.getElementById('sectionsContainer');
            container.innerHTML = '';
            
            // Add drop zone to container (remove first to prevent stacking)
            container.removeEventListener('dragover', handleDragOver);
            container.removeEventListener('drop', handleDrop);
            container.addEventListener('dragover', handleDragOver);
            container.addEventListener('drop', handleDrop);
            
            currentProject.sections.forEach((section, index) => {
                renderSubsectionTree(section, [index], container, 0);
            });
        }
        // Recursive rendering for sections and sub-sections
        function buildTinyMCEContentStyle(isDark) {
            return `
                body { 
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
                    font-size: 14px;
                    line-height: 1.6;
                    color: ${isDark ? '#ffffff' : '#212529'};
                    background-color: ${isDark ? '#2b3035' : '#ffffff'};
                }
                h1, h2, h3, h4, h5, h6 { 
                    color: ${isDark ? '#ffffff' : '#212529'};
                    margin-top: 1.5em;
                    margin-bottom: 0.5em;
                }
                p { 
                    margin-bottom: 1em; 
                    color: ${isDark ? '#ffffff' : '#212529'};
                }
                li {
                    color: ${isDark ? '#ffffff' : '#212529'};
                }
                table { border-collapse: collapse; width: 100%; }
                th, td { 
                    border: 1px solid ${isDark ? '#495057' : '#dee2e6'}; 
                    padding: 8px; 
                    color: ${isDark ? '#ffffff' : '#212529'};
                }
                th { 
                    background-color: ${isDark ? '#343a40' : '#f8f9fa'}; 
                    color: ${isDark ? '#ffffff' : '#212529'};
                }
                code { 
                    background-color: ${isDark ? '#343a40' : '#f8f9fa'}; 
                    padding: 2px 4px; 
                    border-radius: 3px; 
                    color: ${isDark ? '#ffffff' : '#212529'};
                }
                pre { 
                    background-color: ${isDark ? '#343a40' : '#f8f9fa'}; 
                    padding: 1em; 
                    border-radius: 4px; 
                    overflow-x: auto; 
                    color: ${isDark ? '#ffffff' : '#212529'};
                }
                a {
                    color: #0d6efd;
                }
                .xref {
                    color: #2563eb;
                    text-decoration: underline;
                    cursor: default;
                    font-style: italic;
                }
            `;
        }

        function getTinyMCEBaseConfig(selector, isDark) {
            return {
                selector,
                height: 300,
                license_key: 'gpl',
                skin: isDark ? 'oxide-dark' : 'oxide',
                content_css: isDark ? 'dark' : 'default',
                menubar: 'file edit view insert format tools table help',
                plugins: [
                    'advlist', 'autolink', 'lists', 'link', 'image', 'charmap', 'preview',
                    'anchor', 'searchreplace', 'visualblocks', 'code', 'fullscreen',
                    'insertdatetime', 'media', 'table', 'help', 'wordcount', 'save',
                    'emoticons'
                ],
                toolbar: [
                    'undo redo | formatselect | bold italic underline strikethrough',
                    'alignleft aligncenter alignright alignjustify | bullist numlist outdent indent | link image media table | code fullscreen help | citations | xref | insertfield'
                ].join(' | '),
                font_size_formats: '8pt 10pt 12pt 14pt 16pt 18pt 24pt 36pt 48pt',
                font_family_formats: 'Arial=arial,helvetica,sans-serif; Courier New=courier new,courier,monospace; Times New Roman=times new roman,times,serif; Verdana=verdana,geneva,sans-serif; Georgia=georgia,palatino,serif; Trebuchet MS=trebuchet ms,geneva,sans-serif; Comic Sans MS=comic sans ms,sans-serif;',
                content_style: buildTinyMCEContentStyle(isDark),
                images_upload_handler: (blobInfo) => new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result);
                    reader.onerror = () => reject('Image upload failed');
                    reader.readAsDataURL(blobInfo.blob());
                }),
                file_picker_types: 'image',
                file_picker_callback: (callback, value, meta) => {
                    if (meta.filetype === 'image') {
                        const input = document.createElement('input');
                        input.setAttribute('type', 'file');
                        input.setAttribute('accept', 'image/*');
                        input.onchange = () => {
                            const file = input.files[0];
                            if (file) {
                                const reader = new FileReader();
                                reader.onload = () => callback(reader.result, { title: file.name });
                                reader.readAsDataURL(file);
                            }
                        };
                        input.click();
                    }
                }
            };
        }

        function renderSubsectionTree(node, path, parentContainer, depth) {
            const numberStr = path.map(i => i + 1).join('.');
            const nodeDiv = document.createElement('div');
            const lockedClass = node.locked ? ' section-locked' : '';
            nodeDiv.className = (depth === 0 ? 'section-item' : 'subsection-item') + lockedClass;
            nodeDiv.style.marginLeft = '';
            nodeDiv.style.marginBottom = '16px';
            nodeDiv.setAttribute('id', `section-${path.join('-')}`);
            nodeDiv.setAttribute('data-path', JSON.stringify(path));
            nodeDiv.setAttribute('draggable', node.locked ? 'false' : 'true');
            
            // Add drag and drop event listeners
            nodeDiv.addEventListener('dragstart', handleDragStart);
            nodeDiv.addEventListener('dragover', handleDragOver);
            nodeDiv.addEventListener('drop', handleDrop);
            nodeDiv.addEventListener('dragenter', handleDragEnter);
            nodeDiv.addEventListener('dragleave', handleDragLeave);
            nodeDiv.addEventListener('dragend', handleDragEnd);
            
            const pk = path.join('-');
            const unresolvedCount = (node.comments || []).filter(c => !c.resolved).length;
            nodeDiv.innerHTML = `
                <div class="d-flex align-items-center mb-1" style="justify-content: space-between;">
                  <div class="d-flex align-items-center" style="gap: 8px;">
                    <div id="draghandle-${pk}" class="drag-handle" draggable="true"
                        style="cursor:${node.locked ? 'not-allowed' : 'grab'}; padding:4px; color:#6c757d; opacity:${node.locked ? '0.3' : '1'}; ${node.locked ? 'pointer-events:none;' : ''}">
                        <i class="fas fa-grip-vertical"></i>
                    </div>
                    <input id="titleinput-${pk}" type="text" class="form-control form-control-sm"
                        value="${escapeHtml(node.title)}" style="width: 220px;"
                        ${node.locked ? 'disabled' : ''}
                        onchange="updateSubsectionTitleByPath(${JSON.stringify(path)}, this.value)">
                    <button id="deletebtn-${pk}" class="btn btn-sm btn-outline-danger btn-icon"
                        ${node.locked ? 'disabled' : ''}
                        onclick="removeSubsectionByPath(${JSON.stringify(path)})">
                        <i class="fas fa-trash"></i>
                    </button>
                    <button id="dupbtn-${pk}" class="btn btn-sm btn-outline-secondary btn-icon"
                        ${node.locked ? 'disabled' : ''}
                        title="Duplicate section"
                        onclick="duplicateSectionByPath(${JSON.stringify(path)})">
                        <i class="fas fa-copy"></i>
                    </button>
                    <button id="addbtn-${pk}" class="btn btn-sm btn-outline-primary btn-icon"
                        ${node.locked ? 'disabled' : ''}
                        onclick="addSubsectionByPath(${JSON.stringify(path)})">
                        <i class="fas fa-plus"></i> Sub-section
                    </button>
                    <button id="lockbtn-${pk}" class="btn btn-sm ${node.locked ? 'btn-warning' : 'btn-outline-secondary'} btn-icon"
                        title="${node.locked ? 'Unlock section' : 'Lock section'}"
                        onclick="toggleSectionLock(${JSON.stringify(path)})">
                        <i class="fas fa-${node.locked ? 'lock' : 'lock-open'}"></i>
                    </button>
                    <button id="commentbtn-${pk}"
                        class="btn btn-sm ${unresolvedCount > 0 ? 'btn-info' : 'btn-outline-secondary'} btn-icon"
                        title="${unresolvedCount > 0 ? unresolvedCount + ' unresolved comment(s)' : 'Comments'}"
                        onclick="showCommentsModal(${JSON.stringify(path)})">
                        <i class="fas fa-comment"></i>${unresolvedCount > 0 ? `<span class="ms-1" style="font-size:0.75em;">${unresolvedCount}</span>` : ''}
                    </button>
                  </div>
                  <div class="d-flex align-items-center gap-2">
                    <span id="wc-${pk}" class="text-muted" style="font-size:0.75em; white-space:nowrap;"></span>
                    <span class="badge bg-secondary" style="font-size: 1em;">${numberStr}</span>
                  </div>
                </div>
                <textarea id="editor-${pk}" class="tinymce-editor"></textarea>
                <div id="subsections-${pk}" class="subsections-container"></div>
            `;
            parentContainer.appendChild(nodeDiv);
            
            // Add mousedown event to drag handle to prevent text selection
            const dragHandle = nodeDiv.querySelector('.drag-handle');
            if (dragHandle) {
                dragHandle.addEventListener('mousedown', (e) => {
                    // Prevent text selection during drag
                    e.preventDefault();
                });
            }
            
            // Initialize TinyMCE
            const isDarkTheme = currentTheme === 'dark';
            tinymce.init({
                ...getTinyMCEBaseConfig(`#editor-${path.join('-')}`, isDarkTheme),
                setup: function(editor) {
                    editor.addShortcut('ctrl+s', 'Save project', () => {
                        updateAllSectionContents();
                        saveProject();
                    });
                    editor.addShortcut('ctrl+f', 'Find & Replace', () => {
                        showFindReplaceModal();
                    });
                    const _icons = window.bytedraftIcons || {};
                    editor.ui.registry.addIcon('bytedraft-cite',  _icons.cite  || '[Cite]');
                    editor.ui.registry.addIcon('bytedraft-xref',  _icons.xref  || '[XRef]');
                    editor.ui.registry.addIcon('bytedraft-field', _icons.field || '{{}}');
                    editor.ui.registry.addMenuButton('insertfield', {
                        icon: 'bytedraft-field',
                        tooltip: 'Insert field placeholder',
                        fetch: function(callback) {
                            if (!customFields.length) {
                                callback([{ type: 'menuitem', text: 'No fields defined', enabled: false, onAction: () => {} }]);
                                return;
                            }
                            callback(customFields.map(f => ({
                                type: 'menuitem',
                                text: f.name,
                                onAction: function() {
                                    editor.insertContent(`{{${f.name}}}`);
                                }
                            })));
                        }
                    });
                    editor.ui.registry.addButton('citations', {
                        icon: 'bytedraft-cite',
                        tooltip: 'Insert citation',
                        onAction: function() {
                            window._activeCitationEditor = editor;
                            showCitationManagerModal();
                        }
                    });
                    editor.ui.registry.addButton('xref', {
                        icon: 'bytedraft-xref',
                        tooltip: 'Insert cross-reference to another section',
                        onAction: function() {
                            window._activeXRefEditor = editor;
                            showCrossRefModal();
                        }
                    });
                    editor.on('BeforeExecCommand', function(e) {
                        if (e.command === 'mceFullScreen' && node.locked) {
                            e.preventDefault();
                            showToast('Unlock the section to use fullscreen', 'warning');
                        }
                    });
                    editor.on('init', function() {
                        editor.setContent(node.content || '');
                        updateSectionWordCount(path.join('-'), node.content || '');
                        updateDocumentWordCount();
                        if (node.locked) enforceLockOnEditor(editor);
                    });
                    editor.on('Change KeyUp', function() {
                        const html = editor.getContent();
                        setNodeContentByPath(path, html);
                        currentProject.updatedAt = new Date().toISOString();
                        updateSectionWordCount(path.join('-'), html);
                        setTimeout(() => { updateTOCPreview(); updateDocumentWordCount(); }, 500);
                    });
                }
            });
            // Render children
            if (node.subsections && node.subsections.length > 0) {
                const subContainer = nodeDiv.querySelector(`#subsections-${path.join('-')}`);
                if (subContainer) {
                    node.subsections.forEach((sub, idx) => {
                        renderSubsectionTree(sub, path.concat(idx), subContainer, depth + 1);
                    });
                }
            }
        }
        // Helper functions for recursive data access
        function getNodeByPath(path) {
            if (!currentProject || !currentProject.sections || !path || path.length === 0) {
                return null;
            }
            
            let node = currentProject.sections[path[0]];
            if (!node) return null;
            
            for (let i = 1; i < path.length; i++) {
                if (!node.subsections || !node.subsections[path[i]]) {
                    return null;
                }
                node = node.subsections[path[i]];
            }
            return node;
        }
        function setNodeContentByPath(path, content) {
            let node = getNodeByPath(path);
            if (node) {
                node.content = content;
            } else {
                console.warn('Node not found for path:', path, 'Content not saved');
            }
        }
        function updateSubsectionTitleByPath(path, value) {
            let node = getNodeByPath(path);
            node.title = value;
            currentProject.updatedAt = new Date().toISOString();
            setTimeout(() => updateTOCPreview(), 100);
        }
        function removeSubsectionByPath(path) {
            if (path.length === 1) {
                currentProject.sections.splice(path[0], 1);
            } else {
                let parent = getNodeByPath(path.slice(0, -1));
                parent.subsections.splice(path[path.length - 1], 1);
            }
            renderSections();
        }
        function addSubsectionByPath(path) {
            let node = getNodeByPath(path);
            if (!node.subsections) node.subsections = [];
            node.subsections.push({
                id: Date.now().toString(),
                title: 'New Sub-section',
                content: '',
                subsections: []
            });
            renderSections();
        }

        function assignNewIds(node) {
            node.id = Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 9);
            if (node.subsections) node.subsections.forEach(sub => assignNewIds(sub));
        }

        function duplicateSectionByPath(path) {
            const source = getNodeByPath(path);
            if (!source) return;
            updateAllSectionContents();
            const clone = JSON.parse(JSON.stringify(source));
            clone.title = clone.title + ' (Copy)';
            clone.locked = false;
            assignNewIds(clone);
            if (path.length === 1) {
                currentProject.sections.splice(path[0] + 1, 0, clone);
            } else {
                const parent = getNodeByPath(path.slice(0, -1));
                parent.subsections.splice(path[path.length - 1] + 1, 0, clone);
            }
            saveProjectData();
            renderSections();
            setTimeout(() => updateTOCPreview(), 100);
        }

        function addSection() {
            if (!currentProject) return;
            currentProject.sections.push({
                id: Date.now().toString(),
                title: 'New Section',
                content: '',
                subsections: []
            });
            renderSections();
        }




        // Persists project data to storage without creating a revision history entry.
        // Called by auto-save — never logs to version history.
        function saveProjectData() {
            if (!currentProject) return;
            const projectIndex = projects.findIndex(p => p.id === currentProject.id);
            if (projectIndex !== -1) {
                projects[projectIndex] = { ...currentProject };
                saveProjects();
            }
        }

        // Manual save: persists data AND logs a revision history entry.
        function saveProject() {
            if (!currentProject) return;

            saveProjectData();

            const projectIndex = projects.findIndex(p => p.id === currentProject.id);
            if (projectIndex !== -1) {
                versionHistory.push({
                    id: Date.now().toString(),
                    projectId: currentProject.id,
                    timestamp: new Date().toISOString(),
                    description: 'Manual save'
                });
                saveVersionHistory();

                const toast = document.createElement('div');
                toast.className = 'alert alert-success position-fixed';
                toast.style.cssText = 'top: 20px; left: 50%; transform: translateX(-50%); z-index: 9999;';
                toast.textContent = 'Project saved successfully!';
                document.body.appendChild(toast);
                setTimeout(() => toast.remove(), 3000);
            }
        }

        function deleteProject(projectId) {
            showConfirm(
                'Are you sure you want to delete this project? This action cannot be undone.',
                function() {
                    // Remove from projects array
                    projects = projects.filter(p => p.id !== projectId);
                    saveProjects();

                    // Remove from version history
                    versionHistory = versionHistory.filter(v => v.projectId !== projectId);
                    saveVersionHistory();

                    // Remove custom changelog if exists
                    if (customChangelog[projectId]) {
                        delete customChangelog[projectId];
                        saveCustomChangelog();
                    }

                    // If this was the current project, clear it and refresh the main editor window
                    if (currentProject && currentProject.id === projectId) {
                        currentProject = null;
                        document.getElementById('currentProjectTitle').textContent = 'Select a Project';
                        renderProjectContent();
                        tinymce.remove();
                    }

                    renderProjects();
                    showToast('Project deleted.', 'success');
                },
                { title: 'Delete Project', okLabel: 'Delete', okClass: 'btn-danger' }
            );
        }

        function updateProjectStatus(projectId, newStatus) {
            if (!['draft', 'working', 'publish'].includes(newStatus)) return;
            const project = projects.find(p => p.id === projectId);
            if (project) {
                project.status = newStatus;
                project.updatedAt = new Date().toISOString();
                saveProjects();
                
                // Add to version history
                versionHistory.push({
                    id: Date.now().toString(),
                    projectId: projectId,
                    timestamp: new Date().toISOString(),
                    description: `Status changed to ${newStatus}`
                });
                saveVersionHistory();
                
                // Update current project if it's the one being updated
                if (currentProject && currentProject.id === projectId) {
                    currentProject.status = newStatus;
                    currentProject.updatedAt = new Date().toISOString();
                }
                
                // Re-render projects list
                renderProjects();
                
                // Show success message
                showToast('Project status updated to ' + newStatus + '!', 'success');
            }
        }

        // Templates
        function renderTemplates() {
            const container = document.getElementById('templatesList');
            container.innerHTML = '';
            
            Object.entries(templates).forEach(([key, template]) => {
                const card = document.createElement('div');
                card.className = 'template-card';
                card.onclick = () => showTemplatePreview(key);
                
                card.innerHTML = `
                    <h6 class="mb-2">${escapeHtml(template.name)}</h6>
                    <small class="text-muted">${template.sections.length} sections</small>
                `;
                
                container.appendChild(card);
            });
        }

        function showTemplatePreview(templateKey) {
            const template = templates[templateKey];
            if (!template) return;
            document.getElementById('templatePreviewTitle').textContent = template.name;
            document.getElementById('templatePreviewDesc').textContent =
                `${template.sections.length} section${template.sections.length !== 1 ? 's' : ''}`;
            const ul = document.getElementById('templatePreviewSections');
            ul.innerHTML = '';
            template.sections.forEach(s => {
                const li = document.createElement('li');
                li.className = 'list-group-item';
                li.textContent = s.title;
                ul.appendChild(li);
            });
            bootstrap.Modal.getOrCreateInstance(document.getElementById('templatePreviewModal')).show();
        }

        // Custom Fields
        function addCustomField() {
            const name = document.getElementById('newFieldName').value;
            const value = document.getElementById('newFieldValue').value;
            const type = document.getElementById('newFieldType').value;
            
            if (!name) {
                showToast('Please enter a field name', 'warning');
                return;
            }

            customFields.push({
                id: Date.now().toString(),
                name: name,
                value: value,
                type: type
            });
            
            saveCustomFields();
            renderCustomFields();
            
            bootstrap.Modal.getInstance(document.getElementById('newFieldModal')).hide();
        }

        function renderCustomFields() {
            const container = document.getElementById('customFieldsList');
            container.innerHTML = '';

            customFields.forEach(field => {
                const fieldDiv = document.createElement('div');
                fieldDiv.className = 'field-item';
                fieldDiv.style.cssText = 'flex-direction: column; align-items: stretch; gap: 4px;';
                fieldDiv.innerHTML = `
                    <div style="display:flex; align-items:center; gap:4px;">
                        <input type="text" class="form-control form-control-sm"
                               value="${escapeHtml(field.name)}"
                               onchange="updateFieldName('${field.id}', this.value)"
                               placeholder="Field name"
                               title="Edit field name (keyword)"
                               style="font-weight:500;">
                        <code id="field-badge-${field.id}"
                              style="font-size:0.78em; white-space:nowrap; padding:2px 5px; background:var(--bs-secondary-bg); border-radius:4px; border:1px solid var(--bs-border-color);"
                              title="Use this placeholder in your document">&#123;&#123;${escapeHtml(field.name)}&#125;&#125;</code>
                        <button class="btn btn-sm btn-outline-danger flex-shrink-0" onclick="removeField('${field.id}')">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                    <input type="text" class="form-control form-control-sm"
                           value="${escapeHtml(field.value)}"
                           onchange="updateFieldValue('${field.id}', this.value)"
                           placeholder="Value"
                           title="Field value substituted at export">
                `;

                // Live-update the badge when the name input changes
                const nameInput = fieldDiv.querySelector('input');
                const badge = fieldDiv.querySelector(`#field-badge-${field.id}`);
                if (nameInput && badge) {
                    nameInput.addEventListener('input', function() {
                        badge.innerHTML = '&#123;&#123;' + escapeHtml(this.value) + '&#125;&#125;';
                    });
                }

                container.appendChild(fieldDiv);
            });
        }

        function updateFieldName(fieldId, name) {
            if (!name.trim()) return;
            const field = customFields.find(f => f.id === fieldId);
            if (field) {
                field.name = name.trim();
                saveCustomFields();
            }
        }

        function updateFieldValue(fieldId, value) {
            const field = customFields.find(f => f.id === fieldId);
            if (field) {
                field.value = value;
                saveCustomFields();
            }
        }

        function removeField(fieldId) {
            customFields = customFields.filter(f => f.id !== fieldId);
            saveCustomFields();
            renderCustomFields();
        }

        // Export Functions
        async function exportDocument(format) {
            if (!currentProject) {
                showToast('Please select a project first', 'warning');
                return;
            }

            switch (format) {
                case 'docx':
                    await exportToDocx();
                    break;

            }
        }

        function resolveFieldsInProject(project) {
            if (!customFields.length) return project;
            const clone = JSON.parse(JSON.stringify(project));
            function walkSections(sections) {
                sections.forEach(section => {
                    if (section.content) section.content = applyFieldSubstitutions(section.content);
                    if (section.subsections && section.subsections.length) walkSections(section.subsections);
                });
            }
            if (clone.sections) walkSections(clone.sections);
            return clone;
        }

        async function exportToDocx() {
            try {
                updateAllSectionContents();
                saveProject();

                const exportProject = resolveFieldsInProject(currentProject);

                // Use the modern DOCX export implementation
                await exportProjectToDocxModern(
                    exportProject,
                    customChangelog[currentProject.id] || '',
                    versionHistory.filter(v => v.projectId === currentProject.id)
                );
                
            } catch (error) {
                console.error('Error in DOCX export:', error);
                showToast('Error creating DOCX file. Please try again.', 'error');
            }
        }



        function downloadFile(blob, filename) {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        // Version History
        function showVersionHistory() {
            if (!currentProject) {
                showToast('Please select a project first', 'warning');
                return;
            }
            
            const projectVersions = versionHistory.filter(v => v.projectId === currentProject.id);
            const container = document.getElementById('versionHistoryList');
            container.innerHTML = '';
            
            projectVersions.forEach(version => {
                const versionDiv = document.createElement('div');
                versionDiv.className = 'version-item';
                versionDiv.innerHTML = `
                    <div class="d-flex justify-content-between">
                        <div>
                            <strong>${escapeHtml(version.description)}</strong>
                            <br>
                            <small class="text-muted">${new Date(version.timestamp).toLocaleString()}</small>
                        </div>
                    </div>
                `;
                container.appendChild(versionDiv);
            });
            
            bootstrap.Modal.getOrCreateInstance(document.getElementById('versionHistoryModal')).show();
        }

        // Utility Functions
        function showNewProjectModal() {
            bootstrap.Modal.getOrCreateInstance(document.getElementById('newProjectModal')).show();
        }

        function showNewFieldModal() {
            bootstrap.Modal.getOrCreateInstance(document.getElementById('newFieldModal')).show();
        }

        function toggleSidebar() {
            document.getElementById('sidebar').classList.toggle('show');
        }

        // Storage Functions
        function saveProjects() {
            safeSetItem('bytedraft_projects', JSON.stringify(projects));
        }

        function saveCustomFields() {
            safeSetItem('bytedraft_fields', JSON.stringify(customFields));
        }

        function applyFieldSubstitutions(html) {
            if (!customFields.length) return html;
            const fieldMap = {};
            customFields.forEach(f => { fieldMap[f.name.toLowerCase()] = f.value; });
            return html.replace(/\{\{([^}]+)\}\}/g, (match, name) => {
                const val = fieldMap[name.trim().toLowerCase()];
                return val !== undefined ? escapeHtml(val) : match;
            });
        }

        function saveVersionHistory() {
            safeSetItem('bytedraft_versions', JSON.stringify(versionHistory));
        }

        function saveCustomChangelog() {
            // This function is kept for backward compatibility but now uses the table data
            if (currentProject) {
                // The data is already saved in updateChangelogCell, just ensure it's persisted
                safeSetItem('bytedraft_custom_changelog', JSON.stringify(customChangelog));
            }
        }

        function showCustomChangelogModal() {
            if (!currentProject) {
                showToast('Please select a project first', 'warning');
                return;
            }
            
            renderChangelogTable();
            bootstrap.Modal.getOrCreateInstance(document.getElementById('customChangelogModal')).show();
        }

        // Auto-save every 30 seconds — data only, no revision history entry
        setInterval(() => {
            if (currentProject) {
                saveProjectData();
            }
        }, 30000);

        // --- TOC PREVIEW STYLE ---
        function updateTOCPreview() {
            if (!currentProject) {
                const container = document.getElementById('tocPreview');
                if (container) {
                    container.innerHTML = '<small class="text-muted">Select a project to see TOC</small>';
                }
                return;
            }
            const toc = generateTOC();
            const pageMap = calculatePageNumbers();
            const container = document.getElementById('tocPreview');
            if (!container) return;
            container.innerHTML = '';
            // Add page summary
            const totalPages = pageMap.length ? pageMap[pageMap.length - 1] : '';
            container.innerHTML += `<div class="mb-2"><strong>Estimated Pages:</strong> <span class="text-muted">${totalPages || ''} total</span></div>`;
            // TOC Title
            container.innerHTML += `<div style="font-size:1.3em;font-weight:bold;color:var(--primary-color);margin-bottom:8px;">Table of Contents</div>`;
            // TOC Items with page numbers
            toc.forEach((item, idx) => {
                container.innerHTML += `<div class="toc-item level-${item.level}" data-path="${item.path}" style="padding-left:${(item.level - 1) * 16}px; cursor:pointer;"><span style="font-weight:bold;">${item.number}.</span> ${escapeHtml(item.title)}<span class="text-muted ms-1 float-end">${pageMap[idx] || ''}</span></div>`;
            });
            // Add click handler for TOC navigation
            container.onclick = function(e) {
                const item = e.target.closest('.toc-item[data-path]');
                if (item) {
                    const section = document.getElementById('section-' + item.dataset.path);
                    if (section) section.scrollIntoView({behavior: 'smooth', block: 'start'});
                }
            };
        }



        // Header/Footer Modal logic
        function showHeaderFooterModal() {
            if (!currentProject) {
                showToast('Please select a project first', 'warning');
                return;
            }
            // Load from localStorage or project
            let headerFooter = safeParseJSON(localStorage.getItem('bytedraft_header_footer'), {});
            let projectHeaderFooter = headerFooter[currentProject.id] || { header: '', footer: '' };
            document.getElementById('headerContent').value = projectHeaderFooter.header || '';
            document.getElementById('footerContent').value = projectHeaderFooter.footer || '';
            const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('headerFooterModal'));
            modal.show();
        }
        function saveHeaderFooter() {
            if (!currentProject) return;
            let headerFooter = safeParseJSON(localStorage.getItem('bytedraft_header_footer'), {});
            headerFooter[currentProject.id] = {
                header: document.getElementById('headerContent').value,
                footer: document.getElementById('footerContent').value
            };
            safeSetItem('bytedraft_header_footer', JSON.stringify(headerFooter));
            bootstrap.Modal.getInstance(document.getElementById('headerFooterModal')).hide();
        }
        // --- Update content capture for export ---
        // Before exporting, walk all sections and subsections recursively and update their .content from TinyMCE:
        function updateAllSectionContents() {
          function updateNodeContent(node, path) {
            const editorId = `editor-${path.join('-')}`;
            const editor = tinymce.get(editorId);
            if (editor) {
              node.content = editor.getContent();
            }
            if (node.subsections && node.subsections.length > 0) {
              node.subsections.forEach((sub, idx) => {
                updateNodeContent(sub, path.concat(idx));
              });
            }
          }
          if (currentProject && currentProject.sections) {
            currentProject.sections.forEach((section, idx) => {
              updateNodeContent(section, [idx]);
            });
          }
        }
        // Call updateAllSectionContents() before any export or saveProject() call.




        

        

        

        
        
        
        
        

        // JSON Export and Import Functions
        function exportProjectAsJSON(projectId) {
            const project = projects.find(p => p.id === projectId);
            if (!project) {
                showToast('Project not found', 'error');
                return;
            }

            // Get additional project data from localStorage
            const allDocInfo = safeParseJSON(localStorage.getItem('bytedraft_docinfo'), {});
            const allCustomChangelog = safeParseJSON(localStorage.getItem('bytedraft_custom_changelog'), {});
            const allHeaderFooter = safeParseJSON(localStorage.getItem('bytedraft_header_footer'), {});
            const allVersionHistory = safeParseJSON(localStorage.getItem('bytedraft_versions'), []);
            const allLogos = safeParseJSON(localStorage.getItem('bytedraft_logos'), {});
            const allPageSettings = safeParseJSON(localStorage.getItem('bytedraft_page_settings'), {});

            // Filter version history for this specific project
            const projectVersionHistory = allVersionHistory.filter(v => v.projectId === projectId);

            // Create a clean copy of the project for export
            const exportData = {
                ...project,
                documentInfo: allDocInfo[projectId] || {},
                customChangelog: allCustomChangelog[projectId] || {},
                headerFooter: allHeaderFooter[projectId] || {},
                versionHistory: projectVersionHistory,
                logo: allLogos[projectId] || null,
                pageSettings: allPageSettings[projectId] || {},
                exportedAt: new Date().toISOString(),
                version: '1.0'
            };

            const jsonString = JSON.stringify(exportData, null, 2);
            const blob = new Blob([jsonString], { type: 'application/json' });
            downloadFile(blob, `${project.name.replace(/[^a-z0-9]/gi, '_')}.json`);
        }

        function importProjectFromJSON() {
            document.getElementById('jsonImportInput').click();
        }

        function handleJSONImport(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const importData = JSON.parse(e.target.result);
                    
                    // Validate the imported data
                    if (!importData.name || !importData.sections) {
                        showToast('Invalid project file: Missing required fields (name, sections)', 'error');
                        return;
                    }

                    // Check if project with same name already exists
                    const existingProject = projects.find(p => p.name === importData.name);
                    if (existingProject) {
                        const newName = prompt(`A project named "${importData.name}" already exists. Please enter a new name:`, `${importData.name}_imported`);
                        if (!newName || newName.trim() === '') {
                            return;
                        }
                        importData.name = newName.trim();
                    }

                    // Generate new ID and timestamps for the imported project
                    const importedProject = {
                        ...importData,
                        id: Date.now().toString(),
                        createdAt: new Date().toISOString(),
                        updatedAt: new Date().toISOString(),
                        importedAt: new Date().toISOString()
                    };

                    // Remove export-specific fields
                    delete importedProject.exportedAt;
                    delete importedProject.version;

                    // Extract additional data before removing it from the project object
                    const documentInfo = importedProject.documentInfo || {};
                    const importedCustomChangelog = importedProject.customChangelog || {};
                    const headerFooter = importedProject.headerFooter || {};
                    const importedVersionHistory = importedProject.versionHistory || [];
                    const importedLogo = importedProject.logo || null;
                    const importedPageSettings = importedProject.pageSettings || {};

                    // Remove additional data from the project object
                    delete importedProject.documentInfo;
                    delete importedProject.customChangelog;
                    delete importedProject.headerFooter;
                    delete importedProject.versionHistory;
                    delete importedProject.logo;
                    delete importedProject.pageSettings;

                    // Add to projects array
                    projects.push(importedProject);
                    
                    // Save additional data to localStorage
                    const allDocInfo = safeParseJSON(localStorage.getItem('bytedraft_docinfo'), {});
                    const allCustomChangelog = safeParseJSON(localStorage.getItem('bytedraft_custom_changelog'), {});
                    const allHeaderFooter = safeParseJSON(localStorage.getItem('bytedraft_header_footer'), {});
                    const allVersionHistory = safeParseJSON(localStorage.getItem('bytedraft_versions'), []);
                    const allLogos = safeParseJSON(localStorage.getItem('bytedraft_logos'), {});
                    const allPageSettings = safeParseJSON(localStorage.getItem('bytedraft_page_settings'), {});

                    allDocInfo[importedProject.id] = documentInfo;
                    // Migrate imported changelog from possible double-encoded format
                    if (typeof importedCustomChangelog === 'string') {
                        try { allCustomChangelog[importedProject.id] = JSON.parse(importedCustomChangelog); }
                        catch(e) { allCustomChangelog[importedProject.id] = []; }
                    } else {
                        allCustomChangelog[importedProject.id] = importedCustomChangelog;
                    }
                    allHeaderFooter[importedProject.id] = headerFooter;
                    if (importedLogo) {
                        allLogos[importedProject.id] = importedLogo;
                    }
                    allPageSettings[importedProject.id] = importedPageSettings;

                    // Update version history entries with new project ID
                    const updatedVersionHistory = importedVersionHistory.map(v => ({
                        ...v,
                        projectId: importedProject.id
                    }));
                    allVersionHistory.push(...updatedVersionHistory);

                    safeSetItem('bytedraft_docinfo', JSON.stringify(allDocInfo));
                    safeSetItem('bytedraft_custom_changelog', JSON.stringify(allCustomChangelog));
                    safeSetItem('bytedraft_header_footer', JSON.stringify(allHeaderFooter));
                    safeSetItem('bytedraft_versions', JSON.stringify(allVersionHistory));
                    safeSetItem('bytedraft_logos', JSON.stringify(allLogos));
                    safeSetItem('bytedraft_page_settings', JSON.stringify(allPageSettings));
                    
                    // Save to localStorage
                    saveProjects();
                    
                    // Refresh the UI
                    renderProjects();
                    
                    // Select the newly imported project
                    selectProject(importedProject.id);
                    
                    // Refresh the global variables to include the imported data
                    customChangelog = safeParseJSON(localStorage.getItem('bytedraft_custom_changelog'), {});
                    versionHistory = safeParseJSON(localStorage.getItem('bytedraft_versions'), []);
                    
                    // Update the document info display if the modal is open
                    const docInfoModal = document.getElementById('documentInfoModal');
                    if (docInfoModal && bootstrap.Modal.getInstance(docInfoModal) && bootstrap.Modal.getInstance(docInfoModal)._isShown) {
                        loadDocumentInfo();
                    }
                    
                    // Update the changelog display if the modal is open
                    const changelogModal = document.getElementById('customChangelogModal');
                    if (changelogModal && bootstrap.Modal.getInstance(changelogModal) && bootstrap.Modal.getInstance(changelogModal)._isShown) {
                        renderChangelogTable();
                    }
                    
                    // Update the version history display if the modal is open
                    const versionHistoryModal = document.getElementById('versionHistoryModal');
                    if (versionHistoryModal && bootstrap.Modal.getInstance(versionHistoryModal) && bootstrap.Modal.getInstance(versionHistoryModal)._isShown) {
                        showVersionHistory();
                    }
                    
                    showToast(`Project "${importedProject.name}" imported successfully!`, 'success');
                    
                } catch (error) {
                    console.error('Error importing project:', error);
                    showToast('Error importing project: Invalid JSON file', 'error');
                }
            };
            reader.onerror = function() {
                showToast('Failed to read file. Please try again.', 'error');
            };

            reader.readAsText(file);
            
            // Reset the file input
            event.target.value = '';
        }

        // Document Info Table Storage Helpers
        function getDocumentInfo(projectId) {
            const allInfo = safeParseJSON(localStorage.getItem('bytedraft_docinfo'), {});
            return allInfo[projectId] || {
                title: '',
                author: '',
                docOwner: '',
                procOwner: '',
                version: '',
                effDate: '',
                lastRev: '',
                nextRev: '',
                link: ''
            };
        }
        function setDocumentInfo(projectId, info) {
            const allInfo = safeParseJSON(localStorage.getItem('bytedraft_docinfo'), {});
            allInfo[projectId] = info;
            safeSetItem('bytedraft_docinfo', JSON.stringify(allInfo));
        }
        function showDocumentInfoModal() {
            if (!currentProject) {
                showToast('Please select a project first', 'warning');
                return;
            }
            const info = getDocumentInfo(currentProject.id);
            document.getElementById('docInfoTitle').value = info.title || '';
            document.getElementById('docInfoAuthor').value = info.author || '';
            document.getElementById('docInfoDocOwner').value = info.docOwner || '';
            document.getElementById('docInfoProcOwner').value = info.procOwner || '';
            document.getElementById('docInfoVersion').value = info.version || '';
            document.getElementById('docInfoEffDate').value = info.effDate || '';
            document.getElementById('docInfoLastRev').value = info.lastRev || '';
            document.getElementById('docInfoNextRev').value = info.nextRev || '';
            document.getElementById('docInfoLink').value = info.link || '';
            bootstrap.Modal.getOrCreateInstance(document.getElementById('documentInfoModal')).show();
        }
        function saveDocumentInfo() {
            if (!currentProject) return;
            const info = {
                title: document.getElementById('docInfoTitle').value,
                author: document.getElementById('docInfoAuthor').value,
                docOwner: document.getElementById('docInfoDocOwner').value,
                procOwner: document.getElementById('docInfoProcOwner').value,
                version: document.getElementById('docInfoVersion').value,
                effDate: document.getElementById('docInfoEffDate').value,
                lastRev: document.getElementById('docInfoLastRev').value,
                nextRev: document.getElementById('docInfoNextRev').value,
                link: document.getElementById('docInfoLink').value
            };
            setDocumentInfo(currentProject.id, info);
            bootstrap.Modal.getInstance(document.getElementById('documentInfoModal')).hide();
        }



        // Restore generateTOC function for TOC preview and exports
        function generateTOC() {
            if (!currentProject || !currentProject.sections) return [];
            const toc = [];
            function walk(node, numberParts, level, path) {
                toc.push({
                    level: level,
                    number: numberParts.join('.'),
                    title: node.title,
                    page: '', // Page calculation can be added later
                    path: path.join('-')
                });
                if (node.subsections && node.subsections.length > 0) {
                    node.subsections.forEach((sub, idx) => {
                        walk(sub, numberParts.concat(idx + 1), level + 1, path.concat(idx));
                    });
                }
            }
            currentProject.sections.forEach((section, idx) => {
                walk(section, [idx + 1], 1, [idx]);
            });
            return toc;
        }

        function calculatePageNumbers() {
            const toc = generateTOC();
            if (!currentProject || !toc.length) return toc.map(() => '');

            const settings = safeParseJSON(localStorage.getItem('bytedraft_page_settings'), {});
            const ps = settings[currentProject.id] || {};
            const charsPerLine = parseInt(ps.charsPerLine) || 80;
            const linesPerPage = parseInt(ps.linesPerPage) || 40;
            const charsPerPage = charsPerLine * linesPerPage;

            let cumulativeChars = 0;
            const startPage = 4; // after title page, changelog, TOC

            return toc.map(item => {
                const pathArray = item.path.split('-').map(Number);
                const node = getNodeByPath(pathArray);
                const pageNum = startPage + Math.floor(cumulativeChars / charsPerPage);
                cumulativeChars += (node?.content || '').replace(/<[^>]*>/g, '').length;
                return pageNum;
            });
        }


        function renderChangelogTable() {
            const projectId = currentProject?.id;
            let data = customChangelog[projectId] || [];
            const tbody = document.getElementById('changelogTableBody');
            tbody.innerHTML = '';
            if (!data.length) data = [{version:'', date:'', author:'', reviewer:'', approver:'', desc:''}];
            data.forEach((row, idx) => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td><input type="text" class="form-control form-control-sm" value="${escapeHtml(row.version||'')}" onchange="updateChangelogCell(${idx},'version',this.value)"></td>
                    <td><input type="date" class="form-control form-control-sm" value="${escapeHtml(row.date||'')}" onchange="updateChangelogCell(${idx},'date',this.value)"></td>
                    <td><input type="text" class="form-control form-control-sm" value="${escapeHtml(row.author||'')}" onchange="updateChangelogCell(${idx},'author',this.value)"></td>
                    <td><input type="text" class="form-control form-control-sm" value="${escapeHtml(row.reviewer||'')}" onchange="updateChangelogCell(${idx},'reviewer',this.value)"></td>
                    <td><input type="text" class="form-control form-control-sm" value="${escapeHtml(row.approver||'')}" onchange="updateChangelogCell(${idx},'approver',this.value)"></td>
                    <td><input type="text" class="form-control form-control-sm" value="${escapeHtml(row.desc||'')}" onchange="updateChangelogCell(${idx},'desc',this.value)"></td>
                    <td><button class="btn btn-sm btn-danger" onclick="removeChangelogRow(${idx})"><i class="fas fa-trash"></i></button></td>
                `;
                tbody.appendChild(tr);
            });
        }
        function addChangelogRow() {
            const projectId = currentProject?.id;
            let data = customChangelog[projectId] || [];
            data.push({version:'', date:'', author:'', reviewer:'', approver:'', desc:''});
            customChangelog[projectId] = data;
            // Save to localStorage immediately
            safeSetItem('bytedraft_custom_changelog', JSON.stringify(customChangelog));
            renderChangelogTable();
        }
        function removeChangelogRow(idx) {
            const projectId = currentProject?.id;
            let data = customChangelog[projectId] || [];
            data.splice(idx, 1);
            customChangelog[projectId] = data;
            // Save to localStorage immediately
            safeSetItem('bytedraft_custom_changelog', JSON.stringify(customChangelog));
            renderChangelogTable();
        }
        function updateChangelogCell(idx, key, value) {
            const projectId = currentProject?.id;
            let data = customChangelog[projectId] || [];
            if (!data[idx]) data[idx] = {version:'', date:'', author:'', reviewer:'', approver:'', desc:''};
            data[idx][key] = value;
            customChangelog[projectId] = data;
            // Save to localStorage immediately
            safeSetItem('bytedraft_custom_changelog', JSON.stringify(customChangelog));
        }
        function saveCustomChangelogTable() {
            // Ensure data is saved to localStorage
            if (currentProject) {
                safeSetItem('bytedraft_custom_changelog', JSON.stringify(customChangelog));
            }
            bootstrap.Modal.getInstance(document.getElementById('customChangelogModal')).hide();
            showToast('Changelog saved successfully.', 'success');
        }
        function showCustomChangelogModal() {
            renderChangelogTable();
            bootstrap.Modal.getOrCreateInstance(document.getElementById('customChangelogModal')).show();
        }


        // Theme Management
        let currentTheme = 'light';

        function initTheme() {
            const savedTheme = localStorage.getItem('bytedraft-theme') || 'light';
            currentTheme = savedTheme;
            document.documentElement.setAttribute('data-theme', currentTheme);
            updateThemeToggleButton();
        }

        function toggleTheme() {
            saveProject();  // ensure current content is persisted
            currentTheme = currentTheme === 'light' ? 'dark' : 'light';
            document.documentElement.setAttribute('data-theme', currentTheme);
            safeSetItem('bytedraft-theme', currentTheme);
            updateThemeToggleButton();
            if (currentProject) {
                renderProjectContent();  // renderSections() inside calls tinymce.remove()
            }
        }

        function updateThemeToggleButton() {
            const themeToggle = document.getElementById('themeToggle');
            const icon = themeToggle.querySelector('i');
            if (currentTheme === 'dark') {
                icon.className = 'fas fa-sun';
            } else {
                icon.className = 'fas fa-moon';
            }
        }

        function forceThemeRefresh() {
            // Force a complete re-render of all sections with current theme
            if (currentProject && window.tinymce && window.tinymce.editors) {
                const editors = window.tinymce.editors;
                
                // Store content from all editors
                const editorContents = {};
                editors.forEach(editor => {
                    const editorId = editor.id;
                    editorContents[editorId] = editor.getContent();
                });
                
                // Destroy all editors
                editors.forEach(editor => {
                    editor.destroy();
                });
                
                // Clear and re-render sections
                const container = document.getElementById('sectionsContainer');
                if (container && currentProject.sections) {
                    container.innerHTML = '';
                    currentProject.sections.forEach((section, index) => {
                        renderSubsectionTree(section, [index], container, 0);
                    });
                }
                
                // Restore content
                setTimeout(() => {
                    Object.keys(editorContents).forEach(editorId => {
                        const editor = window.tinymce.get(editorId);
                        if (editor) {
                            editor.setContent(editorContents[editorId]);
                        }
                    });
                }, 200);
            }
        }

        // Initialize theme when page loads
        document.addEventListener('DOMContentLoaded', function() {
            initTheme();
            
            // Add event listener for theme toggle button
            document.getElementById('themeToggle').addEventListener('click', toggleTheme);
            
            // Ensure theme is applied to any existing editors after a short delay
            setTimeout(() => {
                if (currentProject && window.tinymce) {
                    forceThemeRefresh();
                }
            }, 500);
        });

        // Drag and Drop functionality for reordering sections and subsections
        let draggedElement = null;
        let draggedPath = null;

        function handleDragStart(e) {
            // The dragged element is the section itself
            const sectionElement = e.target;
            if (!sectionElement || !sectionElement.hasAttribute('data-path')) return;
            
            draggedElement = sectionElement;
            draggedPath = JSON.parse(sectionElement.getAttribute('data-path'));
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/html', sectionElement.outerHTML);
            
            // Add visual feedback to the section
            sectionElement.style.opacity = '0.5';
            sectionElement.classList.add('dragging');
            
            // Set the drag image
            e.dataTransfer.setDragImage(sectionElement, 0, 0);
            
        }

        function handleDragOver(e) {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            
            // Add visual feedback for drop zones
            const targetElement = e.target.closest('[data-path]');
            if (targetElement && targetElement !== draggedElement) {
                // Remove any existing drop classes
                targetElement.classList.remove('drag-over', 'drop-as-child', 'drop-as-section');
                
                // Determine drop action based on position
                const rect = targetElement.getBoundingClientRect();
                const dropY = e.clientY;
                const elementCenterY = rect.top + rect.height / 2;
                const targetPath = JSON.parse(targetElement.getAttribute('data-path'));
                
                if (dropY > elementCenterY && targetPath.length === 1) {
                    // Dropping in lower half of a section - will become child
                    targetElement.classList.add('drop-as-child');
                } else if (dropY <= elementCenterY && draggedPath && draggedPath.length > 1 && targetPath.length === 1) {
                    // Dropping in upper half of a section with a subsection - will become section
                    targetElement.classList.add('drop-as-section');
                } else {
                    // Default drop zone
                    targetElement.classList.add('drag-over');
                }
            }
        }

        function handleDragEnter(e) {
            e.preventDefault();
            const targetElement = e.target.closest('[data-path]');
            if (targetElement && targetElement !== draggedElement) {
                targetElement.classList.add('drag-over');
            }
        }

        function handleDragLeave(e) {
            const targetElement = e.target.closest('[data-path]');
            if (targetElement) {
                targetElement.classList.remove('drag-over', 'drop-as-child', 'drop-as-section');
            }
        }

        function handleDragEnd(e) {
            // Clean up any remaining drag state
            if (draggedElement) {
                draggedElement.style.opacity = '';
                draggedElement.classList.remove('dragging');
            }
            
            // Remove all drop-related classes from all elements
            document.querySelectorAll('.drag-over, .drop-as-child, .drop-as-section').forEach(el => {
                el.classList.remove('drag-over', 'drop-as-child', 'drop-as-section');
            });
            
            // Reset drag state
            draggedElement = null;
            draggedPath = null;
        }

        function handleDrop(e) {
            e.preventDefault();
            e.stopPropagation();

            const targetElement = e.target.closest('[data-path]');
            if (!targetElement) {
                return;
            }

            if (!draggedElement || !draggedPath) {
                return;
            }

            const targetPath = JSON.parse(targetElement.getAttribute('data-path'));

            // Remove visual feedback
            draggedElement.style.opacity = '';
            draggedElement.classList.remove('dragging');
            targetElement.classList.remove('drag-over', 'drop-as-child', 'drop-as-section');

            // Don't allow dropping on itself
            if (JSON.stringify(draggedPath) === JSON.stringify(targetPath)) {
                draggedElement = null;
                draggedPath = null;
                return;
            }

            // Determine the intended target based on the drop position
            const rect = targetElement.getBoundingClientRect();
            const dropY = e.clientY;
            const elementCenterY = rect.top + rect.height / 2;

            // If dropping in the upper half of the target, insert before it
            // If dropping in the lower half, insert as a child (for sections) or after it (for subsections)
            let insertAsChild = false;
            let finalTargetPath = targetPath;

            if (dropY > elementCenterY) {
                // Dropping in lower half
                if (targetPath.length === 1) {
                    // Dropping on a section - make it a child
                    insertAsChild = true;
                    finalTargetPath = targetPath; // Keep the section path, insertAsChild will handle it
                } else {
                    // Dropping on a subsection - insert after it
                    finalTargetPath = [...targetPath.slice(0, -1), targetPath[targetPath.length - 1] + 1];
                }
            } else {
                // Dropping in upper half - insert before the target
                finalTargetPath = targetPath;
            }

            // Special case: if dragging a subsection to a section's upper half, make it a top-level section
            if (draggedPath.length > 1 && targetPath.length === 1 && dropY <= elementCenterY) {
                insertAsChild = false;
                finalTargetPath = targetPath; // This will insert before the target section
            }

            const sourcePath = draggedPath;
            try {
                moveNodeByPath(sourcePath, finalTargetPath, insertAsChild);
            } catch (err) {
                showToast('Failed to move section. Please try again.', 'error');
                renderSections();
            } finally {
                draggedElement = null;
                draggedPath = null;
            }
        }

        function moveNodeByPath(sourcePath, targetPath, insertAsChild = false) {
            if (!currentProject) return;
            
            // Get the source node
            const sourceNode = getNodeByPath(sourcePath);
            if (!sourceNode) return;
            
            // Special case: if source is a subsection and target is a section, make it a top-level section
            
            // Clean up TinyMCE editors before moving to prevent path conflicts
            tinymce.remove();
            
            // Remove from source location
            if (sourcePath.length === 1) {
                // Moving a top-level section
                currentProject.sections.splice(sourcePath[0], 1);
            } else {
                // Moving a subsection
                const sourceParent = getNodeByPath(sourcePath.slice(0, -1));
                sourceParent.subsections.splice(sourcePath[sourcePath.length - 1], 1);
            }
            
            // Insert at target location
            if (insertAsChild && targetPath.length === 1) {
                // Inserting as a child of a section
                const targetSection = getNodeByPath(targetPath);
                if (!targetSection) {
                    console.error('Target section not found:', targetPath);
                    return;
                }
                if (!targetSection.subsections) {
                    targetSection.subsections = [];
                }
                targetSection.subsections.unshift(sourceNode); // Add as first child
            } else if (targetPath.length === 1) {
                // Moving to top-level sections
                if (targetPath[0] >= currentProject.sections.length) {
                    // If target index is beyond current sections, append to end
                    currentProject.sections.push(sourceNode);
                } else {
                    currentProject.sections.splice(targetPath[0], 0, sourceNode);
                }
            } else if (targetPath.length > 1) {
                // Moving to subsections
                const targetParent = getNodeByPath(targetPath.slice(0, -1));
                if (!targetParent) {
                    console.error('Target parent not found:', targetPath.slice(0, -1));
                    return;
                }
                if (!targetParent.subsections) {
                    targetParent.subsections = [];
                }
                targetParent.subsections.splice(targetPath[targetPath.length - 1], 0, sourceNode);
            }
            
            // Re-render sections to update the UI with a small delay to ensure TinyMCE cleanup
            setTimeout(() => {
                renderSections();
                // Update TOC to reflect new structure
                setTimeout(() => updateTOCPreview(), 100);
            }, 50);
            
            // Update project timestamp
            currentProject.updatedAt = new Date().toISOString();
            
            showToast('Section reordered successfully!', 'success');
        }


        // Page Settings and Logo Functions
        function showPageSettings() {
            if (!currentProject) {
                showToast('Please select a project first', 'warning');
                return;
            }
            
            // Load current page settings
            const pageSettings = safeParseJSON(localStorage.getItem('bytedraft_page_settings'), {});
            const projectSettings = pageSettings[currentProject.id] || {
                paperSize: 'letter',
                charsPerLine: 80,
                linesPerPage: 40,
                headerHeight: 3,
                paragraphSpacing: 2
            };

            document.getElementById('paperSize').value = projectSettings.paperSize || 'letter';
            document.getElementById('charsPerLine').value = projectSettings.charsPerLine;
            document.getElementById('linesPerPage').value = projectSettings.linesPerPage;
            document.getElementById('headerHeight').value = projectSettings.headerHeight;
            document.getElementById('paragraphSpacing').value = projectSettings.paragraphSpacing;
            
            // Load logo
            loadLogoPreview();
            
            bootstrap.Modal.getOrCreateInstance(document.getElementById('pageSettingsModal')).show();
        }

        function autoSavePageSettings() {
            if (!currentProject) return;
            const pageSettings = safeParseJSON(localStorage.getItem('bytedraft_page_settings'), {});
            pageSettings[currentProject.id] = {
                paperSize: document.getElementById('paperSize').value,
                charsPerLine: parseInt(document.getElementById('charsPerLine').value),
                linesPerPage: parseInt(document.getElementById('linesPerPage').value),
                headerHeight: parseInt(document.getElementById('headerHeight').value),
                paragraphSpacing: parseInt(document.getElementById('paragraphSpacing').value)
            };
            safeSetItem('bytedraft_page_settings', JSON.stringify(pageSettings));
        }

        function savePageSettings() {
            if (!currentProject) return;
            autoSavePageSettings();
            saveLogo();
            bootstrap.Modal.getInstance(document.getElementById('pageSettingsModal')).hide();
            showToast('Page settings saved.', 'success');
        }

        // Logo Functions
        function handleLogoUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            
            if (!file.type.startsWith('image/')) {
                showToast('Please select an image file (PNG, JPG, GIF)', 'warning');
                return;
            }
            
            const reader = new FileReader();
            reader.onload = function(e) {
                const logoData = e.target.result;
                displayLogoPreview(logoData);
                
                // Store logo data
                const logoStorage = safeParseJSON(localStorage.getItem('bytedraft_logos'), {});
                logoStorage[currentProject.id] = logoData;
                safeSetItem('bytedraft_logos', JSON.stringify(logoStorage));
                
            };
            reader.readAsDataURL(file);
        }

        function displayLogoPreview(logoData) {
            const preview = document.getElementById('logoPreview');
            const removeBtn = document.getElementById('removeLogoBtn');
            
            preview.innerHTML = `<img src="${logoData}" style="max-width: 100%; max-height: 100%; object-fit: contain;">`;
            removeBtn.style.display = 'block';
        }

        function loadLogoPreview() {
            if (!currentProject) return;
            
            const logoStorage = safeParseJSON(localStorage.getItem('bytedraft_logos'), {});
            const logoData = logoStorage[currentProject.id];
            
            if (logoData) {
                displayLogoPreview(logoData);
            } else {
                const preview = document.getElementById('logoPreview');
                const removeBtn = document.getElementById('removeLogoBtn');
                
                preview.innerHTML = '<small class="text-muted">No logo</small>';
                removeBtn.style.display = 'none';
            }
        }

        function removeLogo() {
            if (!currentProject) return;
            
            const logoStorage = safeParseJSON(localStorage.getItem('bytedraft_logos'), {});
            delete logoStorage[currentProject.id];
            safeSetItem('bytedraft_logos', JSON.stringify(logoStorage));
            
            const preview = document.getElementById('logoPreview');
            const removeBtn = document.getElementById('removeLogoBtn');
            
            preview.innerHTML = '<small class="text-muted">No logo</small>';
            removeBtn.style.display = 'none';
            
            // Clear file input
            document.getElementById('logoUpload').value = '';
        }

        function saveLogo() {
            // Logo is already saved in handleLogoUpload, this function is for consistency
            // with the savePageSettings flow
        }

        function getProjectLogo(projectId) {
            const logoStorage = safeParseJSON(localStorage.getItem('bytedraft_logos'), {});
            return logoStorage[projectId] || null;
        }

        // ── Citation Manager ──────────────────────────────────────────────────

        function getCitations(projectId) {
            const all = safeParseJSON(localStorage.getItem('bytedraft_references'), {});
            return all[projectId] || [];
        }

        function saveCitations(projectId, refs) {
            const all = safeParseJSON(localStorage.getItem('bytedraft_references'), {});
            all[projectId] = refs;
            safeSetItem('bytedraft_references', JSON.stringify(all));
        }

        function showCitationManagerModal() {
            if (!currentProject) {
                showToast('Please select a project first', 'warning');
                return;
            }
            renderCitationsTable();
            bootstrap.Modal.getOrCreateInstance(document.getElementById('citationManagerModal')).show();
        }

        function renderCitationsTable() {
            if (!currentProject) return;
            const refs = getCitations(currentProject.id);
            const tbody = document.getElementById('citationsTableBody');
            if (refs.length === 0) {
                tbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted">No references yet.</td></tr>';
                return;
            }
            tbody.innerHTML = refs.map((ref, idx) => `
                <tr>
                    <td class="text-center">${idx + 1}</td>
                    <td>${escapeHtml(ref.authors)}</td>
                    <td>${escapeHtml(ref.year)}</td>
                    <td>${escapeHtml(ref.title)}</td>
                    <td>
                        <button class="btn btn-xs btn-outline-primary btn-sm me-1" onclick="insertCitationIntoEditor(${idx})" title="Insert [${idx + 1}]">
                            [${idx + 1}]
                        </button>
                        <button class="btn btn-xs btn-outline-secondary btn-sm me-1" onclick="showAddEditReferenceModal(${idx})" title="Edit">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn btn-xs btn-outline-danger btn-sm" onclick="deleteReference(${idx})" title="Delete">
                            <i class="fas fa-trash"></i>
                        </button>
                    </td>
                </tr>
            `).join('');
        }

        function showAddEditReferenceModal(idx) {
            const modalEl = document.getElementById('addEditReferenceModal');
            document.getElementById('addEditReferenceModalTitle').textContent = idx < 0 ? 'Add Reference' : 'Edit Reference';
            document.getElementById('editingRefIndex').value = idx;

            if (idx >= 0 && currentProject) {
                const refs = getCitations(currentProject.id);
                const ref = refs[idx] || {};
                document.getElementById('refTitle').value = ref.title || '';
                document.getElementById('refAuthors').value = ref.authors || '';
                document.getElementById('refYear').value = ref.year || '';
                document.getElementById('refSource').value = ref.source || '';
                document.getElementById('refUrl').value = ref.url || '';
                document.getElementById('refNotes').value = ref.notes || '';
            } else {
                document.getElementById('refTitle').value = '';
                document.getElementById('refAuthors').value = '';
                document.getElementById('refYear').value = '';
                document.getElementById('refSource').value = '';
                document.getElementById('refUrl').value = '';
                document.getElementById('refNotes').value = '';
            }

            // Show on top of citation manager
            bootstrap.Modal.getOrCreateInstance(modalEl).show();
        }

        function saveReference() {
            if (!currentProject) return;
            const title = document.getElementById('refTitle').value.trim();
            if (!title) {
                showToast('Title is required', 'warning');
                return;
            }
            const ref = {
                id: 'ref-' + Date.now(),
                title,
                authors: document.getElementById('refAuthors').value.trim(),
                year: document.getElementById('refYear').value.trim(),
                source: document.getElementById('refSource').value.trim(),
                url: document.getElementById('refUrl').value.trim(),
                notes: document.getElementById('refNotes').value.trim()
            };

            const refs = getCitations(currentProject.id);
            const idx = parseInt(document.getElementById('editingRefIndex').value, 10);
            if (idx >= 0) {
                ref.id = refs[idx].id; // preserve original id on edit
                refs[idx] = ref;
            } else {
                refs.push(ref);
            }
            saveCitations(currentProject.id, refs);
            renderCitationsTable();

            bootstrap.Modal.getInstance(document.getElementById('addEditReferenceModal')).hide();
            showToast(idx >= 0 ? 'Reference updated' : 'Reference added', 'success');
        }

        function deleteReference(idx) {
            if (!currentProject) return;
            const refs = getCitations(currentProject.id);
            refs.splice(idx, 1);
            saveCitations(currentProject.id, refs);
            renderCitationsTable();
            showToast('Reference deleted', 'info');
        }

        function insertCitationIntoEditor(idx) {
            const editor = window._activeCitationEditor;
            if (!editor) {
                showToast('No active editor — click inside an editor section first', 'warning');
                return;
            }
            const num = idx + 1;
            editor.insertContent(`<sup>[${num}]</sup>`);
            bootstrap.Modal.getInstance(document.getElementById('citationManagerModal')).hide();
        }

        function showCrossRefModal() {
            if (!currentProject) {
                showToast('Please select a project first', 'warning');
                return;
            }
            const toc = generateTOC();
            const tbody = document.getElementById('crossRefTableBody');
            if (toc.length === 0) {
                tbody.innerHTML = '<tr><td colspan="3" class="text-center text-muted">No sections yet.</td></tr>';
            } else {
                tbody.innerHTML = toc.map(entry => {
                    const label = `Section ${escapeHtml(entry.number)} \u2014 ${escapeHtml(entry.title)}`;
                    return `<tr>
                        <td>${escapeHtml(entry.number)}</td>
                        <td>${escapeHtml(entry.title)}</td>
                        <td>
                            <button class="btn btn-sm btn-outline-primary"
                                onclick="insertCrossRef('${escapeHtml(entry.path)}', '${label.replace(/'/g, "\\'")}')">
                                Insert
                            </button>
                        </td>
                    </tr>`;
                }).join('');
            }
            bootstrap.Modal.getOrCreateInstance(document.getElementById('crossRefModal')).show();
        }

        function insertCrossRef(path, label) {
            const editor = window._activeXRefEditor;
            if (!editor) {
                showToast('No active editor — click inside an editor section first', 'warning');
                return;
            }
            editor.insertContent(
                `<span class="xref" data-path="${path}" contenteditable="false">${label}</span>`
            );
            bootstrap.Modal.getInstance(document.getElementById('crossRefModal')).hide();
        }

        // ── Word Count ────────────────────────────────────────────────────────

        function countWords(html) {
            if (!html) return 0;
            const text = html
                .replace(/<[^>]*>/g, ' ')
                .replace(/&[a-z#0-9]+;/gi, ' ')
                .replace(/\s+/g, ' ')
                .trim();
            return text ? text.split(' ').filter(w => w.length > 0).length : 0;
        }

        function updateSectionWordCount(pathKey, html) {
            const el = document.getElementById(`wc-${pathKey}`);
            if (!el) return;
            const n = countWords(html);
            el.textContent = n > 0 ? `${n.toLocaleString()} words` : '';
        }

        function updateDocumentWordCount() {
            if (!currentProject) return;
            let total = 0;
            function walk(node, path) {
                const editor = window.tinymce && tinymce.get(`editor-${path.join('-')}`);
                total += countWords(editor ? editor.getContent() : (node.content || ''));
                if (node.subsections) {
                    node.subsections.forEach((sub, idx) => walk(sub, path.concat(idx)));
                }
            }
            currentProject.sections.forEach((section, idx) => walk(section, [idx]));
            const el = document.getElementById('wordCountSummary');
            if (!el) return;
            if (total === 0) {
                el.textContent = '';
            } else {
                const mins = Math.ceil(total / 200);
                el.textContent = `${total.toLocaleString()} words · ~${mins} min read`;
            }
        }

        // ── Section Locking ───────────────────────────────────────────────────

        function enforceLockOnEditor(editor) {
            const body = editor.getBody();
            if (body) body.setAttribute('contenteditable', 'false');
            // Grey out the fullscreen button with a small delay to let the toolbar render
            setTimeout(function() {
                const container = editor.getContainer();
                if (!container) return;
                const fsBtn = container.querySelector('.tox-tbtn[data-mce-name="fullscreen"]')
                    || Array.from(container.querySelectorAll('.tox-tbtn'))
                        .find(b => (b.title || '').toLowerCase() === 'fullscreen');
                if (fsBtn) {
                    fsBtn.dataset.wasTitle = fsBtn.title;
                    fsBtn.style.opacity = '0.35';
                    fsBtn.style.cursor = 'not-allowed';
                    fsBtn.style.filter = 'grayscale(1)';
                    fsBtn.title = 'Fullscreen disabled (section is locked)';
                }
            }, 200);
        }

        function releaseLockOnEditor(editor) {
            const body = editor.getBody();
            if (body) body.setAttribute('contenteditable', 'true');
            const container = editor.getContainer();
            if (!container) return;
            const fsBtn = container.querySelector('.tox-tbtn[data-mce-name="fullscreen"]')
                || Array.from(container.querySelectorAll('.tox-tbtn'))
                    .find(b => (b.title || '').toLowerCase() === 'fullscreen'
                            || (b.dataset.wasTitle || '').toLowerCase() === 'fullscreen');
            if (fsBtn) {
                fsBtn.style.opacity = '';
                fsBtn.style.cursor = '';
                fsBtn.style.filter = '';
                fsBtn.title = fsBtn.dataset.wasTitle || 'Fullscreen';
                delete fsBtn.dataset.wasTitle;
            }
        }

        function toggleSectionLock(pathArr) {
            const node = getNodeByPath(pathArr);
            if (!node) return;
            node.locked = !node.locked;
            applyLockState(pathArr, node.locked);
            currentProject.updatedAt = new Date().toISOString();
            const projectIndex = projects.findIndex(p => p.id === currentProject.id);
            if (projectIndex !== -1) projects[projectIndex] = { ...currentProject };
            saveProjects();
            showToast(node.locked ? 'Section locked' : 'Section unlocked', 'info');
        }

        function applyLockState(pathArr, locked) {
            const pk = pathArr.join('-');

            // Section container class and drag
            const sectionDiv = document.getElementById(`section-${pk}`);
            if (sectionDiv) {
                sectionDiv.setAttribute('draggable', locked ? 'false' : 'true');
                if (locked) {
                    sectionDiv.classList.add('section-locked');
                } else {
                    sectionDiv.classList.remove('section-locked');
                }
            }

            // Drag handle
            const dragHandle = document.getElementById(`draghandle-${pk}`);
            if (dragHandle) {
                dragHandle.style.cursor = locked ? 'not-allowed' : 'grab';
                dragHandle.style.opacity = locked ? '0.3' : '1';
                dragHandle.style.pointerEvents = locked ? 'none' : '';
            }

            // Title input
            const titleInput = document.getElementById(`titleinput-${pk}`);
            if (titleInput) titleInput.disabled = locked;

            // Delete button
            const deleteBtn = document.getElementById(`deletebtn-${pk}`);
            if (deleteBtn) deleteBtn.disabled = locked;

            // Add subsection button
            const addBtn = document.getElementById(`addbtn-${pk}`);
            if (addBtn) addBtn.disabled = locked;

            // Lock button icon/class
            const lockBtn = document.getElementById(`lockbtn-${pk}`);
            if (lockBtn) {
                lockBtn.className = `btn btn-sm ${locked ? 'btn-warning' : 'btn-outline-secondary'} btn-icon`;
                lockBtn.title = locked ? 'Unlock section' : 'Lock section';
                lockBtn.innerHTML = `<i class="fas fa-${locked ? 'lock' : 'lock-open'}"></i>`;
            }

            // TinyMCE editor body editability
            const editor = window.tinymce && tinymce.get(`editor-${pk}`);
            if (editor) {
                if (locked) {
                    enforceLockOnEditor(editor);
                } else {
                    releaseLockOnEditor(editor);
                }
            }
        }

        // ── Find & Replace ────────────────────────────────────────────────────

        function showFindReplaceModal() {
            if (!currentProject) {
                showToast('Please select a project first', 'warning');
                return;
            }
            document.getElementById('findInput').value = '';
            document.getElementById('replaceInput').value = '';
            document.getElementById('findReplaceResults').innerHTML = '';
            bootstrap.Modal.getOrCreateInstance(document.getElementById('findReplaceModal')).show();
            setTimeout(() => document.getElementById('findInput').focus(), 300);
        }

        function buildFindRegex() {
            const term = document.getElementById('findInput').value;
            if (!term) return null;
            const caseSensitive = document.getElementById('findCaseSensitive').checked;
            const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            return new RegExp(escaped, caseSensitive ? 'g' : 'gi');
        }

        function walkAllSections(fn) {
            function walk(node, path) {
                fn(node, path);
                if (node.subsections) {
                    node.subsections.forEach((sub, idx) => walk(sub, path.concat(idx)));
                }
            }
            currentProject.sections.forEach((section, idx) => walk(section, [idx]));
        }

        function executeFindAll() {
            if (!currentProject) return;
            const regex = buildFindRegex();
            if (!regex) { showToast('Please enter a search term', 'warning'); return; }

            updateAllSectionContents();

            let totalMatches = 0;
            const hits = [];

            walkAllSections((node, path) => {
                const contentMatches = (node.content || '').match(regex) || [];
                const titleMatches = node.title.match(regex) || [];
                const count = contentMatches.length + titleMatches.length;
                if (count > 0) {
                    totalMatches += count;
                    hits.push({ title: node.title, count, inTitle: titleMatches.length > 0 });
                }
            });

            const resultsDiv = document.getElementById('findReplaceResults');
            if (totalMatches === 0) {
                resultsDiv.innerHTML = '<div class="alert alert-secondary py-2 mb-0">No matches found.</div>';
            } else {
                const rows = hits.map(h =>
                    `<li>${escapeHtml(h.title)}${h.inTitle ? ' <span class="badge bg-secondary">title</span>' : ''} — <strong>${h.count}</strong> match${h.count !== 1 ? 'es' : ''}</li>`
                ).join('');
                resultsDiv.innerHTML = `
                    <div class="alert alert-info py-2 mb-2">
                        Found <strong>${totalMatches}</strong> match${totalMatches !== 1 ? 'es' : ''} in <strong>${hits.length}</strong> section${hits.length !== 1 ? 's' : ''}.
                    </div>
                    <ul class="mb-0 small">${rows}</ul>`;
            }
        }

        function executeReplaceAll() {
            if (!currentProject) return;
            const regex = buildFindRegex();
            if (!regex) { showToast('Please enter a search term', 'warning'); return; }

            const replacement = document.getElementById('replaceInput').value;
            updateAllSectionContents();

            let totalReplaced = 0;
            let sectionsModified = 0;

            walkAllSections((node, path) => {
                let changed = false;

                // Replace in content
                const contentMatches = (node.content || '').match(regex) || [];
                if (contentMatches.length > 0) {
                    node.content = node.content.replace(regex, replacement);
                    totalReplaced += contentMatches.length;
                    changed = true;
                    const editor = tinymce.get(`editor-${path.join('-')}`);
                    if (editor) editor.setContent(node.content);
                }

                // Replace in title
                const titleMatches = node.title.match(regex) || [];
                if (titleMatches.length > 0) {
                    node.title = node.title.replace(regex, replacement);
                    totalReplaced += titleMatches.length;
                    changed = true;
                    const titleInput = document.querySelector(`#section-${path.join('-')} input[type="text"]`);
                    if (titleInput) titleInput.value = node.title;
                }

                if (changed) sectionsModified++;
            });

            if (totalReplaced === 0) {
                document.getElementById('findReplaceResults').innerHTML =
                    '<div class="alert alert-secondary py-2 mb-0">No matches found.</div>';
                return;
            }

            currentProject.updatedAt = new Date().toISOString();
            const projectIndex = projects.findIndex(p => p.id === currentProject.id);
            if (projectIndex !== -1) projects[projectIndex] = { ...currentProject };
            saveProjects();
            updateTOCPreview();

            document.getElementById('findReplaceResults').innerHTML =
                `<div class="alert alert-success py-2 mb-0">Replaced <strong>${totalReplaced}</strong> match${totalReplaced !== 1 ? 'es' : ''} across <strong>${sectionsModified}</strong> section${sectionsModified !== 1 ? 's' : ''}.</div>`;
        }

        // ── Section Comments ──────────────────────────────────────────────────

        let _activeCommentPath = null;

        function showCommentsModal(pathArr) {
            _activeCommentPath = pathArr;
            const node = getNodeByPath(pathArr);
            if (!node) return;
            document.getElementById('commentsSectionTitle').textContent = node.title;
            document.getElementById('newCommentText').value = '';
            renderCommentsList(node);
            bootstrap.Modal.getOrCreateInstance(document.getElementById('commentsModal')).show();
            setTimeout(() => document.getElementById('newCommentText').focus(), 300);
        }

        function renderCommentsList(node) {
            const comments = node.comments || [];
            const el = document.getElementById('commentsList');
            if (comments.length === 0) {
                el.innerHTML = '<p class="text-muted text-center mb-0">No comments yet.</p>';
                return;
            }
            el.innerHTML = comments.map(c => `
                <div class="card mb-2${c.resolved ? ' comment-resolved' : ''}">
                  <div class="card-body py-2 px-3">
                    <div class="d-flex justify-content-between align-items-start gap-2">
                      <p class="mb-1 flex-grow-1"${c.resolved ? ' style="text-decoration:line-through"' : ''}>${escapeHtml(c.text)}</p>
                      <div class="d-flex gap-1 flex-shrink-0">
                        <button class="btn btn-sm ${c.resolved ? 'btn-outline-secondary' : 'btn-outline-success'}"
                            onclick="toggleResolveComment('${c.id}')"
                            title="${c.resolved ? 'Unresolve' : 'Resolve'}">
                          <i class="fas fa-${c.resolved ? 'rotate-left' : 'check'}"></i>
                        </button>
                        <button class="btn btn-sm btn-outline-danger"
                            onclick="deleteSectionComment('${c.id}')" title="Delete">
                          <i class="fas fa-trash"></i>
                        </button>
                      </div>
                    </div>
                    <small class="text-muted">${new Date(c.timestamp).toLocaleString()}${c.resolved ? ' · Resolved' : ''}</small>
                  </div>
                </div>`).join('');
        }

        function addSectionComment() {
            const text = document.getElementById('newCommentText').value.trim();
            if (!text) { showToast('Comment cannot be empty', 'warning'); return; }
            const node = getNodeByPath(_activeCommentPath);
            if (!node) return;
            if (!node.comments) node.comments = [];
            node.comments.push({
                id: Date.now().toString(36) + Math.random().toString(36).slice(2),
                text,
                timestamp: new Date().toISOString(),
                resolved: false
            });
            document.getElementById('newCommentText').value = '';
            renderCommentsList(node);
            updateCommentButton(_activeCommentPath, node);
            _saveCurrentProject();
        }

        function toggleResolveComment(commentId) {
            const node = getNodeByPath(_activeCommentPath);
            if (!node?.comments) return;
            const c = node.comments.find(c => c.id === commentId);
            if (c) c.resolved = !c.resolved;
            renderCommentsList(node);
            updateCommentButton(_activeCommentPath, node);
            _saveCurrentProject();
        }

        function deleteSectionComment(commentId) {
            const node = getNodeByPath(_activeCommentPath);
            if (!node?.comments) return;
            node.comments = node.comments.filter(c => c.id !== commentId);
            renderCommentsList(node);
            updateCommentButton(_activeCommentPath, node);
            _saveCurrentProject();
        }

        function updateCommentButton(pathArr, node) {
            const pk = pathArr.join('-');
            const btn = document.getElementById(`commentbtn-${pk}`);
            if (!btn) return;
            const n = (node.comments || []).filter(c => !c.resolved).length;
            btn.className = `btn btn-sm ${n > 0 ? 'btn-info' : 'btn-outline-secondary'} btn-icon`;
            btn.title = n > 0 ? `${n} unresolved comment(s)` : 'Comments';
            btn.innerHTML = `<i class="fas fa-comment"></i>${n > 0 ? `<span class="ms-1" style="font-size:0.75em;">${n}</span>` : ''}`;
        }

        function _saveCurrentProject() {
            currentProject.updatedAt = new Date().toISOString();
            const idx = projects.findIndex(p => p.id === currentProject.id);
            if (idx !== -1) projects[idx] = { ...currentProject };
            saveProjects();
        }
