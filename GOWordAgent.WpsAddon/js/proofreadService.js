/**
 * 校对工作流
 */

var ProofreadWorkflow = {
    connected: false,

    init: function() {
        this.discoverAndConnect();
    },

    discoverAndConnect: function() {
        var self = this;

        // 应用手动端口配置（如果用户填写了）
        var config = UIController.getConfig();
        if (config.manualPort > 0) {
            ApiClient.defaultPort = config.manualPort;
        }

        // 尝试发现服务
        if (!ApiClient.discoverService()) {
            UIController.setConnected(false);
            setTimeout(function() {
                self.discoverAndConnect();
            }, 2000);
            return;
        }

        // 健康检查
        ApiClient.healthCheck(function(err, data) {
            if (err) {
                console.error('Health check failed:', err);
                UIController.setConnected(false);
                setTimeout(function() {
                    self.discoverAndConnect();
                }, 2000);
                return;
            }

            console.log('Connected to backend');
            self.connected = true;
            UIController.setConnected(true);
            UIController.hideCancelBtn();
        });
    },

    connect: function() {
        var config = UIController.getConfig();
        var self = this;

        // 应用手动端口配置
        if (config.manualPort > 0) {
            ApiClient.defaultPort = config.manualPort;
        }

        UIController.setProgress('正在保存配置...');

        ApiClient.saveConfig({
            provider: config.provider,
            apiKey: config.apiKey,
            apiUrl: config.apiUrl,
            model: config.model,
            autoConnect: true
        }, function(err, data) {
            if (err) {
                alert('保存配置失败: ' + err.message);
                UIController.hideProgress();
                return;
            }

            UIController.setProgress('正在测试连接...');

            // 重新初始化 LLM 服务
            self.discoverAndConnect();

            setTimeout(function() {
                if (self.connected) {
                    alert('连接成功！');
                } else {
                    alert('连接失败，请检查配置');
                }
                UIController.hideProgress();
            }, 1000);
        });
    },

    startProofread: function() {
        if (!this.connected) {
            alert('请先连接到后端服务');
            return;
        }

        var self = this;
        var paragraphs;

        try {
            paragraphs = DocumentService.getDocumentText();
        } catch (e) {
            alert('获取文档内容失败: ' + e.message);
            return;
        }

        if (paragraphs.length === 0) {
            alert('文档内容为空');
            return;
        }

        var config = UIController.getConfig();

        UIController.showResults();
        UIController.setProgress('正在准备校对...');
        document.getElementById('btn-proofread').disabled = true;
        UIController.showCancelBtn();

        var requestData = {
            text: '',
            paragraphs: paragraphs,
            provider: config.provider,
            prompt: '',
            mode: config.mode
        };

        ApiClient.proofread(requestData, function(err, response) {
            document.getElementById('btn-proofread').disabled = false;
            UIController.hideCancelBtn();

            if (err) {
                UIController.setProgress('校对失败: ' + err.message);
                return;
            }

            if (!response || !response.success) {
                UIController.setProgress('校对失败: ' + (response ? '未知错误' : '无响应'));
                return;
            }

            var issues = response.issues || [];
            
            UIController.setIssues(issues);
            UIController.setProgress(
                '校对完成！共发现 ' + issues.length + ' 处问题，' +
                '耗时 ' + response.elapsedSeconds.toFixed(1) + ' 秒'
            );

            // 应用修订到文档
            self.applyRevisions(issues);
        });
    },

    cancelProofread: function() {
        var self = this;
        UIController.setProgress('正在取消校对...');
        ApiClient.cancel(function(err, data) {
            if (err) {
                console.error('Cancel failed:', err);
                UIController.setProgress('取消失败: ' + err.message);
                return;
            }
            UIController.setProgress('已发送取消请求');
            UIController.hideCancelBtn();
        });
    },

    applyRevisions: function(issues) {
        if (!issues || issues.length === 0) return;

        // 按偏移量倒序处理，避免位置偏移
        var sorted = issues.slice().sort(function(a, b) {
            return b.StartOffset - a.StartOffset;
        });

        for (var i = 0; i < sorted.length && i < 50; i++) {
            var issue = sorted[i];
            try {
                var comment = issue.Type + '：' + issue.Reason;
                DocumentService.applyAtOffset(
                    issue.StartOffset,
                    issue.EndOffset,
                    issue.Suggestion,
                    comment
                );
            } catch (e) {
                console.error('Failed to apply revision:', e);
            }
        }
    }
};
