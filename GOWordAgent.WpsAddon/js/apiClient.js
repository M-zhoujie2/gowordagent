/**
 * 后端 API 通信客户端
 * 使用 XMLHttpRequest（兼容性最广）
 */

var ApiClient = {
    baseUrl: '',
    port: 0,

    /**
     * 发现后端服务端口
     */
    discoverService: function() {
        try {
            var fs = wps.FileSystem;
            if (!fs) {
                console.error('FileSystem not available');
                return false;
            }

            var portData = fs.ReadFile('/tmp/gowordagent-port.json');
            if (!portData) {
                console.log('Port file not found');
                return false;
            }

            var info = JSON.parse(portData);
            this.port = info.port;
            this.baseUrl = 'http://127.0.0.1:' + info.port;
            console.log('Discovered service at ' + this.baseUrl);
            return true;
        } catch (e) {
            console.error('Service discovery failed:', e);
            return false;
        }
    },

    /**
     * GET 请求
     */
    get: function(path, callback) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', this.baseUrl + path, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var data = JSON.parse(xhr.responseText);
                        callback(null, data);
                    } catch (e) {
                        callback(e, null);
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
     * POST 请求
     */
    post: function(path, data, callback) {
        var xhr = new XMLHttpRequest();
        xhr.open('POST', this.baseUrl + path, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var response = JSON.parse(xhr.responseText);
                        callback(null, response);
                    } catch (e) {
                        callback(e, null);
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
    }
};
