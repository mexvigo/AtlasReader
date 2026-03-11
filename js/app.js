(function () {
    'use strict';

    const ATLAS_API_URL = 'https://mai-atlas.microsoft.com/api/missions';

    // DOM elements
    const searchInput = document.getElementById('search-input');
    const searchBtn = document.getElementById('search-btn');
    const statusSection = document.getElementById('status-section');
    const statusMessage = document.getElementById('status-message');
    const resultsSection = document.getElementById('results-section');
    const resultsHeading = document.getElementById('results-heading');
    const resultsContainer = document.getElementById('results-container');
    const collapseAllBtn = document.getElementById('collapse-all-btn');
    const expandAllBtn = document.getElementById('expand-all-btn');

    let allMissions = [];
    let missionsLoaded = false;

    // ===== Utility =====

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function highlightText(text, term) {
        if (!term || !text) return escapeHtml(String(text));
        const escaped = escapeHtml(String(text));
        const regex = new RegExp('(' + term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
        return escaped.replace(regex, '<mark>$1</mark>');
    }

    function deepSearch(obj, term) {
        if (obj == null) return false;
        const lowerTerm = term.toLowerCase();
        if (typeof obj === 'string') return obj.toLowerCase().includes(lowerTerm);
        if (typeof obj === 'number' || typeof obj === 'boolean') {
            return String(obj).toLowerCase().includes(lowerTerm);
        }
        if (Array.isArray(obj)) return obj.some(item => deepSearch(item, term));
        if (typeof obj === 'object') {
            return Object.values(obj).some(val => deepSearch(val, term));
        }
        return false;
    }

    // ===== Status =====

    function showStatus(message, type) {
        statusSection.style.display = '';
        statusMessage.className = 'status-message ' + type;
        if (type === 'loading') {
            statusMessage.innerHTML = '<span class="spinner"></span> ' + escapeHtml(message);
        } else {
            statusMessage.textContent = message;
        }
    }

    function hideStatus() {
        statusSection.style.display = 'none';
    }

    // ===== Fetch Missions =====

    async function fetchMissions() {
        if (missionsLoaded) return allMissions;

        showStatus('Loading missions from Atlas...', 'loading');

        const response = await fetch(ATLAS_API_URL, {
            credentials: 'include',
            headers: { 'Accept': 'application/json' }
        });

        if (!response.ok) {
            throw new Error('Failed to fetch missions (HTTP ' + response.status + '). Make sure you are signed into mai-atlas.microsoft.com.');
        }

        const data = await response.json();

        // Handle various response shapes
        if (Array.isArray(data)) {
            allMissions = data;
        } else if (data && Array.isArray(data.value)) {
            allMissions = data.value;
        } else if (data && Array.isArray(data.missions)) {
            allMissions = data.missions;
        } else if (data && Array.isArray(data.data)) {
            allMissions = data.data;
        } else if (data && typeof data === 'object') {
            // Last resort: wrap single object
            allMissions = [data];
        } else {
            allMissions = [];
        }

        missionsLoaded = true;
        return allMissions;
    }

    // ===== Render =====

    function renderField(label, value, term) {
        if (value == null || value === '') return '';
        let html = '<div class="detail-group">';
        html += '<div class="detail-label">' + escapeHtml(label) + '</div>';

        if (Array.isArray(value)) {
            if (value.length === 0) return '';
            // Check if items are simple strings/numbers
            if (value.every(v => typeof v === 'string' || typeof v === 'number')) {
                html += '<div class="tag-list">';
                value.forEach(function (item) {
                    html += '<span class="tag">' + highlightText(String(item), term) + '</span>';
                });
                html += '</div>';
            } else {
                // Complex array - render each as sub-block
                value.forEach(function (item, idx) {
                    html += '<div style="margin-left:1rem; margin-top:0.3rem; padding-left:0.5rem; border-left:2px solid #e0e0e0;">';
                    if (typeof item === 'object' && item !== null) {
                        html += renderObjectFields(item, term);
                    } else {
                        html += '<div class="detail-value">' + highlightText(String(item), term) + '</div>';
                    }
                    html += '</div>';
                });
            }
        } else if (typeof value === 'object') {
            html += '<div style="margin-left:1rem;">' + renderObjectFields(value, term) + '</div>';
        } else {
            const strVal = String(value);
            // Render URLs as links
            if (/^https?:\/\//.test(strVal)) {
                const safeUrl = escapeHtml(strVal);
                html += '<div class="detail-value"><a href="' + safeUrl + '" target="_blank" rel="noopener noreferrer">' + highlightText(strVal, term) + '</a></div>';
            } else {
                html += '<div class="detail-value">' + highlightText(strVal, term) + '</div>';
            }
        }

        html += '</div>';
        return html;
    }

    function renderObjectFields(obj, term) {
        let html = '';
        for (const [key, value] of Object.entries(obj)) {
            const label = key.replace(/([A-Z])/g, ' $1').replace(/[_-]/g, ' ').trim();
            html += renderField(label, value, term);
        }
        return html;
    }

    function getMissionTitle(mission) {
        return mission.title || mission.name || mission.missionName || mission.Name || mission.Title || 'Untitled Mission';
    }

    function renderMission(mission, term, index) {
        const title = getMissionTitle(mission);
        const div = document.createElement('div');
        div.className = 'mission-card';

        const header = document.createElement('div');
        header.className = 'mission-card-header';
        header.innerHTML =
            '<span class="mission-title">' + highlightText(title, term) + '</span>' +
            '<span class="mission-toggle">&#9654;</span>';

        const details = document.createElement('div');
        details.className = 'mission-details';

        // Render all fields
        let detailsHtml = '';
        for (const [key, value] of Object.entries(mission)) {
            const label = key.replace(/([A-Z])/g, ' $1').replace(/[_-]/g, ' ').trim();
            detailsHtml += renderField(label, value, term);
        }
        details.innerHTML = detailsHtml;

        header.addEventListener('click', function () {
            const toggle = header.querySelector('.mission-toggle');
            const isVisible = details.classList.contains('visible');
            if (isVisible) {
                details.classList.remove('visible');
                toggle.classList.remove('expanded');
            } else {
                details.classList.add('visible');
                toggle.classList.add('expanded');
            }
        });

        div.appendChild(header);
        div.appendChild(details);
        return div;
    }

    function renderResults(missions, term) {
        resultsContainer.innerHTML = '';

        if (missions.length === 0) {
            resultsContainer.innerHTML = '<div class="no-results">No missions found matching &ldquo;' + escapeHtml(term) + '&rdquo;</div>';
            resultsHeading.textContent = 'No Results';
        } else {
            resultsHeading.textContent = missions.length + ' Mission' + (missions.length !== 1 ? 's' : '') + ' Found';
            missions.forEach(function (mission, idx) {
                resultsContainer.appendChild(renderMission(mission, term, idx));
            });
        }

        resultsSection.style.display = '';
    }

    // ===== Search =====

    async function performSearch() {
        const term = searchInput.value.trim();
        if (!term) {
            searchInput.focus();
            return;
        }

        searchBtn.disabled = true;
        resultsSection.style.display = 'none';

        try {
            const missions = await fetchMissions();
            hideStatus();

            const filtered = missions.filter(function (mission) {
                return deepSearch(mission, term);
            });

            renderResults(filtered, term);
        } catch (err) {
            showStatus(err.message || 'An error occurred while fetching data.', 'error');
            resultsSection.style.display = 'none';
        } finally {
            searchBtn.disabled = false;
        }
    }

    // ===== Collapse / Expand =====

    function collapseAll() {
        resultsContainer.querySelectorAll('.mission-details.visible').forEach(function (el) {
            el.classList.remove('visible');
            el.previousElementSibling.querySelector('.mission-toggle').classList.remove('expanded');
        });
    }

    function expandAll() {
        resultsContainer.querySelectorAll('.mission-details:not(.visible)').forEach(function (el) {
            el.classList.add('visible');
            el.previousElementSibling.querySelector('.mission-toggle').classList.add('expanded');
        });
    }

    // ===== Event Listeners =====

    searchBtn.addEventListener('click', performSearch);

    searchInput.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
            performSearch();
        }
    });

    collapseAllBtn.addEventListener('click', collapseAll);
    expandAllBtn.addEventListener('click', expandAll);

})();
