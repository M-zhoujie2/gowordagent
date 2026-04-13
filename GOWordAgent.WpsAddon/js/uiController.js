/**
 * UI 控制器
 * 管理界面状态和交互
 */

var UIController = {
    elements: {},

    init: function() {
        // 缓存 DOM 元素
        this.elements = {
            statusBar: document.getElementById('status-bar'),
            connectionStatus: document.getElementById('connection-status'),
            settingsPanel: document.getElementById('settings-panel'),
            resultsPanel: document.getElementById('results-panel'),
            progressInfo: document.getElementById('progress-info'),
            issuesList: document.getElementById('issues-list'),
            providerSelect: document.getElementById('provider-select'),
            apiKeyInput: document.getElementById('api-key'),
            apiUrlInput: document.getElementById('api-url'),
            modelInput: document.getElementById('model'),
            btnConnect: document.getElementById('btn-connect'),
            btnProofread: document.getElementById('btn-proofread'),
            btnClear: document.getElementById('btn-clear')
        };

        this.bindEvents();
    },

    bindEvents: function() {
        var self = this;

        this.elements.btnConnect.addEventListener('click', function() {
            ProofreadWorkflow.connect();
        });

        this.elements.btnProofread.addEventListener('click', function() {
            ProofreadWorkflow.startProofread();
        });

        this.elements.btnClear.addEventListener('click', function() {
            self.clearResults();
        });

        // 模式切换
        var modeRadios = document.querySelectorAll('input[name="mode"]');
        for (var i = 0; i < modeRadios.length; i++) {
            modeRadios[i].addEventListener('change', function() {
                // 模式切换处理
            });
        }
    },

    setConnected: function(connected) {
        var statusEl = this.elements.connectionStatus;
        if (connected) {
            statusEl.textContent = '● 已连接';
            statusEl.className = 'status-connected';
            this.elements.btnProofread.disabled = false;
        } else {
            statusEl.textContent = '● 未连接';
            statusEl.className = 'status-disconnected';
            this.elements.btnProofread.disabled = true;
        }
    },

    showSettings: function() {
        this.elements.settingsPanel.classList.remove('hidden');
        this.elements.resultsPanel.classList.add('hidden');
    },

    showResults: function() {
        this.elements.settingsPanel.classList.add('hidden');
        this.elements.resultsPanel.classList.remove('hidden');
    },

    setProgress: function(message) {
        this.elements.progressInfo.textContent = message;
        this.elements.progressInfo.style.display = 'block';
    },

    hideProgress: function() {
        this.elements.progressInfo.style.display = 'none';
    },

    clearResults: function() {
        this.elements.issuesList.innerHTML = '';
        this.hideProgress();
    },

    addIssue: function(issue) {
        var item = document.createElement('div');
        item.className = 'issue-item';
        
        var severityClass = 'severity-' + (issue.Severity || 'medium');
        var severityText = {
            'high': '高',
            'medium': '中',
            'low': '低'
        }[issue.Severity] || '中';

        item.innerHTML = 
            '<div class="issue-header">' +
                '<span class="issue-type">' + (issue.Type || '问题') + '</span>' +
                '<span class="issue-severity ' + severityClass + '">' + severityText + '</span>' +
            '</div>' +
            '<div class="issue-original">原文：' + this.escapeHtml(issue.Original) + '</div>' +
            '<div class="issue-suggestion">修改：' + this.escapeHtml(issue.Suggestion) + '</div>' +
            '<div class="issue-reason">' + this.escapeHtml(issue.Reason) + '</div>';

        var self = this;
        item.addEventListener('click', function() {
            DocumentService.navigateToOffset(issue.StartOffset, issue.EndOffset);
        });

        this.elements.issuesList.appendChild(item);
    },

    setIssues: function(issues) {
        this.clearResults();
        for (var i = 0; i < issues.length; i++) {
            this.addIssue(issues[i]);
        }
    },

    getConfig: function() {
        var mode = document.querySelector('input[name="mode"]:checked');
        return {
            provider: this.elements.providerSelect.value,
            apiKey: this.elements.apiKeyInput.value,
            apiUrl: this.elements.apiUrlInput.value,
            model: this.elements.modelInput.value,
            mode: mode ? mode.value : 'Precise'
        };
    },

    escapeHtml: function(text) {
        if (!text) return '';
        var div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
};
