/**
 * 后端 API 通信客户端
 * 使用 XMLHttpRequest（兼容性最广）
 */

var ApiClient = {
    baseUrl: '',
    port: 0,
    apiToken: '',
    defaultPort: 0, // 用户手动指定的兜底端口

    /**
     * 获取当前用户名
     */
    getCurrentUser: function() {
        try {
            if (typeof wps !== 'undefined' && wps.Env && wps.Env.UserName) {
                return wps.Env.UserName;
            }
            return 'unknown';
        } catch (e) {
            return 'unknown';
        }
    },

    /**
     * 获取端口文件路径（兼容新旧路径）
     */
    _getPortFilePaths: function() {
        var userName = this.getCurrentUser();
        var runtimeDir = '';
        try {
            if (typeof wps !== 'undefined' && wps.Env && wps.Env.GetEnvironmentVariable) {
                runtimeDir = wps.Env.GetEnvironmentVariable('XDG_RUNTIME_DIR');
            }
        } catch (e) {}
        var paths = [];
        if (runtimeDir) {
            paths.push(runtimeDir + '/gowordagent-port-' + userName + '.json');
        }
        paths.push('/tmp/gowordagent-port-' + userName + '.json');
        paths.push('/tmp/gowordagent-port.json');
        // 备用：配置文件目录
        paths.push('/home/' + userName + '/.config/gowordagent/service-port.json');
        return paths;
    },

    /**
     * 从文件系统发现服务
     */
    _discoverFromFileSystem: function() {
        try {
            var fs = wps.FileSystem;
            if (!fs) {
                console.log('FileSystem not available');
                return null;
            }

            var paths = this._getPortFilePaths();
            for (var i = 0; i < paths.length; i++) {
                try {
                    var portData = fs.ReadFile(paths[i]);
                    if (portData) {
                        console.log('Found port file: ' + paths[i]);
                        return JSON.parse(portData);
                    }
                } catch (e) {}
            }
        } catch (e) {
            console.error('FileSystem discovery failed:', e);
        }
        return null;
    },

    /**
     * 尝试通过默认/手动端口直接连接
     */
    _tryDirectPort: function(port) {
        if (!port || port <= 0) return false;
        this.port = port;
        this.baseUrl = 'http://127.0.0.1:' + this.port;
        console.log('Trying direct port ' + this.baseUrl);
        return true;
    },

    /**
     * 发现后端服务端口
     */
    discoverService: function() {
        // 1. 优先通过文件系统发现
        var info = this._discoverFromFileSystem();
        if (info) {
            this.port = info.port || info.Port;
            this.apiToken = info.apiToken || info.apitoken || '';
            this.baseUrl = 'http://127.0.0.1:' + this.port;
            console.log('Discovered service via FileSystem at ' + this.baseUrl);
            return true;
        }

        // 2. 尝试手动端口
        if (this.defaultPort > 0 && this._tryDirectPort(this.defaultPort)) {
            return true;
        }

        console.log('Port file not found in any known location and no manual port set');
        return false;
    },

    /**
     * 重新发现服务（用于 Token 失效或后端重启后）
     */
    rediscover: function() {
        console.log('Rediscovering service...');
        return this.discoverService();
    },

    /**
     * GET 请求（支持 401 后自动重发现一次）
     */
    get: function(path, callback, _retried) {
        var self = this;
        var xhr = new XMLHttpRequest();
        xhr.open('GET', this.baseUrl + path, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        if (this.apiToken) {
            xhr.setRequestHeader('X-Api-Token', this.apiToken);
        }
        xhr.timeout = 30000; // 30秒超时
        
        xhr.ontimeout = function() {
            callback(new Error('Request timeout'), null);
        };
        
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var data = JSON.parse(xhr.responseText);
                        callback(null, data);
                    } catch (e) {
                        callback(e, null);
                    }
                } else if (xhr.status === 401 && !_retried) {
                    // Token 可能已失效，尝试重新发现服务并重试一次
                    if (self.rediscover()) {
                        self.get(path, callback, true);
                    } else {
                        callback(new Error('Unauthorized and rediscovery failed'), null);
                    }
                } else {
                    callback(new Error('HTTP ' + xhr.status), null);
                }
            }
        };
        
        xhr.onerror = function() {
            callback(new Error('Network error'), null);
        };
        
        xhr.send();
    },

    /**
     * POST 请求（支持 401 后自动重发现一次）
     */
    post: function(path, data, callback, _retried) {
        var self = this;
        var xhr = new XMLHttpRequest();
        xhr.open('POST', this.baseUrl + path, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        if (this.apiToken) {
            xhr.setRequestHeader('X-Api-Token', this.apiToken);
        }
        xhr.timeout = 600000; // 10分钟超时（校对请求可能较长）
        
        xhr.ontimeout = function() {
            callback(new Error('Request timeout'), null);
        };
        
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var response = JSON.parse(xhr.responseText);
                        callback(null, response);
                    } catch (e) {
                        callback(e, null);
                    }
                } else if (xhr.status === 401 && !_retried) {
                    // Token 可能已失效，尝试重新发现服务并重试一次
                    if (self.rediscover()) {
                        self.post(path, data, callback, true);
                    } else {
                        callback(new Error('Unauthorized and rediscovery failed'), null);
                    }
                } else {
                    try {
                        var error = JSON.parse(xhr.responseText);
                        callback(new Error(error.error || 'HTTP ' + xhr.status), null);
                    } catch (e) {
                        callback(new Error('HTTP ' + xhr.status), null);
                    }
                }
            }
        };
        
        xhr.onerror = function() {
            callback(new Error('Network error'), null);
        };
        
        xhr.send(JSON.stringify(data));
    },

    /**
     * 健康检查
     */
    healthCheck: function(callback) {
        this.get('/api/proofread/health', callback);
    },

    /**
     * 执行校对
     */
    proofread: function(data, callback) {
        this.post('/api/proofread', data, callback);
    },

    /**
     * 取消校对
     */
    cancel: function(callback) {
        this.post('/api/proofread/cancel', {}, callback);
    },

    /**
     * 获取配置
     */
    getConfig: function(callback) {
        this.get('/api/proofread/config', callback);
    },

    /**
     * 保存配置
     */
    saveConfig: function(data, callback) {
        this.post('/api/proofread/config', data, callback);
    },

    /**
     * 清空缓存
     */
    clearCache: function(callback) {
        this.post('/api/proofread/clear-cache', {}, callback);
    },

    /**
     * 获取统计信息
     */
    getStats: function(callback) {
        this.get('/api/proofread/stats', callback);
    }
};
