# GOWordAgent 银河麒麟 V10 测试指南

> **版本**：v1.0  
> **日期**：2026-04-13

---

## 目录

1. [测试环境准备](#测试环境准备)
2. [开发环境测试（Windows）](#开发环境测试windows)
3. [PoC 测试（银河麒麟 V10）](#poc-测试银河麒麟-v10)
4. [功能测试](#功能测试)
5. [问题排查](#问题排查)

---

## 测试环境准备

### 环境清单

| 环境 | 操作系统 | 用途 |
|------|----------|------|
| 开发环境 | Windows 10/11 | 代码构建、初步调试 |
| 目标环境 | 银河麒麟 V10 SP1 (x86_64) | 实际部署测试 |
| 备选环境 | Ubuntu 20.04/22.04 (x86_64) | 无麒麟时的替代测试 |

### 必需软件

**开发环境（Windows）**：
- .NET 8 SDK (https://dotnet.microsoft.com/download/dotnet/8.0)
- WPS Office Windows 版（用于初步 JS API 测试）
- Postman 或 curl（API 测试）

**目标环境（银河麒麟 V10）**：
- WPS Office for Linux 12.1.2.25838+
- 浏览器（用于测试 HTTP 访问）
- curl（命令行 HTTP 测试）

---

## 开发环境测试（Windows）

### 1. 构建测试

```powershell
# 1. 进入项目目录
cd D:\Project\word\GOWordAgentAddIn-master

# 2. 构建 Core 类库
dotnet build GOWordAgent.Core\GOWordAgent.Core.csproj -c Release

# 3. 构建后端服务（linux-x64）
dotnet publish GOWordAgent.WpsService\GOWordAgent.WpsService.csproj `
    -c Release -r linux-x64 --self-contained true `
    -p:PublishSingleFile=true -o .\publish\backend

# 4. 检查输出文件
ls .\publish\backend\
```

**期望输出**：
```
gowordagent-server          # 主程序（约 50-80MB）
*.dll                       # 依赖库
```

### 2. Windows 本地运行测试（可选）

修改 `Program.cs` 临时支持 Windows 本地测试：

```csharp
// 修改端口文件路径为 Windows 临时目录
var portFile = Path.Combine(Path.GetTempPath(), "gowordagent-port.json");
```

然后运行：

```powershell
cd GOWordAgent.WpsService
dotnet run

# 检查输出
# 应该显示：GOWordAgent Service started on port XXXX
```

### 3. API 接口测试（使用 Postman/curl）

**健康检查**：
```bash
# 获取端口
PORT=$(cat /tmp/gowordagent-port.json | grep -o '"port":[0-9]*' | cut -d: -f2)

# 测试健康检查
curl http://127.0.0.1:$PORT/api/proofread/health

# 期望响应：
# {"status":"ok","timestamp":"2026-04-13T..."}
```

**校对接口测试**：
```bash
curl -X POST http://127.0.0.1:$PORT/api/proofread \
  -H "Content-Type: application/json" \
  -d '{
    "text": "这是一个测试文档。",
    "paragraphs": [
      {"index":0,"startOffset":0,"endOffset":11,"text":"这是一个测试文档。"}
    ],
    "provider": "DeepSeek",
    "apiKey": "your-api-key",
    "mode": "Precise"
  }'
```

---

## PoC 测试（银河麒麟 V10）

### Day 0：环境侦察

**目标**：确认 WPS 版本和基础环境

```bash
# SSH 连接到麒麟机器
ssh user@kylin-host

# 1. 检查 WPS 版本
wps --version
# 或
cat /opt/kingsoft/wps-office/office6/version.txt

# 期望输出：12.1.2.25838 或更高

# 2. 检查架构
uname -m
# 期望输出：x86_64

# 3. 检查 .NET 8 能否运行（测试依赖库）
ldd --version
# 期望：glibc 2.31+（Ubuntu 20.04 级别）
```

**决策检查表**：
- [ ] WPS 版本 >= 11.1.0.14309 （继续）
- [ ] WPS 版本 < 11.1 （切换到 LibreOffice 方案）

### Day 1：HTTP 通信 PoC

**目标**：验证 WPS 加载项能否与本地 HTTP 服务通信

#### 步骤 1：启动测试后端

在麒麟机器上创建简单测试服务器：

```bash
# 创建测试目录
mkdir -p /tmp/gowordagent-test
cd /tmp/gowordagent-test

# 创建测试服务器 (test_server.py)
cat > test_server.py << 'EOF'
#!/usr/bin/env python3
import json
import http.server
import socketserver
import threading

PORT = 8765

class Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/api/test':
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            response = {'status': 'ok', 'message': 'Hello from Python'}
            self.wfile.write(json.dumps(response).encode())
        else:
            self.send_response(404)
            self.end_headers()
    
    def do_POST(self):
        if self.path == '/api/test':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            response = {
                'status': 'ok',
                'received': json.loads(post_data.decode())
            }
            self.wfile.write(json.dumps(response).encode())
    
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

# 写入端口文件
with open('/tmp/gowordagent-port.json', 'w') as f:
    json.dump({'port': PORT, 'pid': 0, 'timestamp': 0}, f)

print(f'Server running on port {PORT}')
with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
    httpd.serve_forever()
EOF

# 启动测试服务器
chmod +x test_server.py
python3 test_server.py &

# 验证服务器运行
curl http://127.0.0.1:8765/api/test
```

#### 步骤 2：创建 WPS 测试加载项

```bash
# 创建测试加载项目录
mkdir -p ~/.local/share/Kingsoft/wps/jsaddons/com.test.http

cat > ~/.local/share/Kingsoft/wps/jsaddons/com.test.http/package.json << 'EOF'
{
  "name": "http-test",
  "wps": {
    "id": "com.test.http",
    "name": "HTTP测试",
    "version": "1.0.0"
  }
}
EOF

cat > ~/.local/share/Kingsoft/wps/jsaddons/com.test.http/index.html << 'EOF'
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>HTTP测试</title>
</head>
<body>
    <h3>HTTP通信测试</h3>
    <button onclick="testGet()">测试 GET</button>
    <button onclick="testPost()">测试 POST</button>
    <div id="result" style="margin-top: 20px; padding: 10px; background: #f0f0f0;"></div>
    
    <script>
        function log(msg) {
            document.getElementById('result').innerHTML += msg + '<br>';
            console.log(msg);
        }
        
        function testGet() {
            log('=== 测试 GET ===');
            try {
                var xhr = new XMLHttpRequest();
                xhr.open('GET', 'http://127.0.0.1:8765/api/test', false);  // 同步请求
                xhr.send();
                log('Status: ' + xhr.status);
                log('Response: ' + xhr.responseText);
            } catch (e) {
                log('Error: ' + e.message);
            }
        }
        
        function testPost() {
            log('=== 测试 POST ===');
            try {
                var xhr = new XMLHttpRequest();
                xhr.open('POST', 'http://127.0.0.1:8765/api/test', false);
                xhr.setRequestHeader('Content-Type', 'application/json');
                xhr.send(JSON.stringify({test: 'data'}));
                log('Status: ' + xhr.status);
                log('Response: ' + xhr.responseText);
            } catch (e) {
                log('Error: ' + e.message);
            }
        }
        
        // 页面加载时自动测试
        window.onload = function() {
            log('页面加载完成');
            log('User Agent: ' + navigator.userAgent);
        };
    </script>
</body>
</html>
EOF
```

#### 步骤 3：验证测试

1. 重启 WPS 文字
2. 查看右侧边栏是否出现 "HTTP测试" 面板
3. 点击 "测试 GET" 和 "测试 POST" 按钮
4. 观察结果区域输出

**期望结果**：
```
=== 测试 GET ===
Status: 200
Response: {"status": "ok", "message": "Hello from Python"}

=== 测试 POST ===
Status: 200
Response: {"status": "ok", "received": {"test": "data"}}
```

**Go/No-Go 决策**：
- ✅ **通过**：GET/POST 都返回 200，继续后续开发
- ❌ **失败**：任何错误（如 "Network error"、"Blocked"），切换到 LibreOffice 方案

---

## 功能测试

### 测试 1：后端服务启动

```bash
# 1. 复制构建好的后端到测试目录
cp -r /path/to/publish/backend /opt/gowordagent

# 2. 手动启动服务
/opt/gowordagent/gowordagent-server

# 3. 检查端口文件
cat /tmp/gowordagent-port.json
# 期望：{"port":xxxxx,"pid":xxxxx,"timestamp":xxxxx}

# 4. 测试健康检查
PORT=$(cat /tmp/gowordagent-port.json | python3 -c "import json,sys; print(json.load(sys.stdin)['port'])")
curl http://127.0.0.1:$PORT/api/proofread/health
```

### 测试 2：WPS 加载项加载

```bash
# 1. 安装加载项
mkdir -p ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin
cp -r /path/to/addon/* ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin/

# 2. 检查文件
ls ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin/
# 期望：index.html, main.js, css/, js/

# 3. 重启 WPS，检查侧边栏
```

### 测试 3：完整校对流程

**前置条件**：
- 后端服务已启动
- WPS 加载项已加载
- 有有效的 AI API Key

**测试步骤**：

1. **连接后端**
   - 打开 WPS 文字
   - 在智能校对面板中输入 API Key
   - 点击 "保存并连接"
   - 期望：状态显示 "已连接"

2. **提取文档文本**
   - 打开一个包含文字的文档
   - 点击 "开始校对"
   - 期望：显示 "正在准备校对..."

3. **执行校对**
   - 等待校对完成
   - 期望：显示 "校对完成！共发现 X 处问题"

4. **查看结果**
   - 检查问题列表是否显示
   - 每个问题应显示：类型、严重程度、原文、修改建议、理由

5. **点击定位**
   - 点击某个问题
   - 期望：WPS 跳转到文档对应位置

6. **检查修订**
   - 在 WPS 中打开 "审阅" 面板
   - 期望：看到修订标记和批注

### 测试 4：异常情况

**后端未启动**：
```bash
# 停止后端
systemctl --user stop gowordagent
rm -f /tmp/gowordagent-port.json

# 重启 WPS，检查加载项行为
# 期望：显示 "未连接"，提示用户启动服务
```

**网络异常**：
```bash
# 使用无效 API Key 测试错误处理
# 期望：友好的错误提示，不崩溃
```

---

## 问题排查

### 问题 1：后端无法启动

**现象**：执行 `gowordagent-server` 无输出或报错

**排查步骤**：
```bash
# 1. 检查文件完整性
ls -la /opt/gowordagent/

# 2. 检查依赖库
ldd /opt/gowordagent/gowordagent-server

# 3. 手动运行查看详细错误
/opt/gowordagent/gowordagent-server 2>&1

# 4. 检查端口占用
netstat -tlnp | grep gowordagent
fuser /tmp/gowordagent-port.json
```

**常见原因**：
- glibc 版本过低（需要 2.31+）
- 端口被占用
- 权限不足

### 问题 2：WPS 加载项不显示

**现象**：侧边栏没有 "智能校对" 面板

**排查步骤**：
```bash
# 1. 检查加载项目录
ls -la ~/.local/share/Kingsoft/wps/jsaddons/

# 2. 检查文件权限
ls -la ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin/

# 3. 检查 package.json 格式
cat ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin/package.json

# 4. 查看 WPS 日志
tail -f ~/.local/share/Kingsoft/wps/log/*/jsaddons.log
```

**常见原因**：
- package.json 格式错误
- 文件权限问题（应为 644 或 755）
- WPS 版本过低

### 问题 3：HTTP 通信失败

**现象**：加载项显示 "未连接" 或请求超时

**排查步骤**：
```bash
# 1. 确认后端运行
systemctl --user status gowordagent

# 2. 检查端口文件
cat /tmp/gowordagent-port.json

# 3. 测试本地访问
PORT=$(cat /tmp/gowordagent-port.json | grep -o '"port":[0-9]*' | cut -d: -f2)
curl -v http://127.0.0.1:$PORT/api/proofread/health

# 4. 检查防火墙
sudo iptables -L | grep $PORT
```

**常见原因**：
- 后端未启动
- 防火墙阻止
- WPS WebView 安全策略限制

### 问题 4：文本定位失败

**现象**：点击问题后没有跳转到正确位置

**排查步骤**：
```bash
# 1. 检查返回的偏移量是否正确
curl -X POST http://127.0.0.1:$PORT/api/proofread \
  -H "Content-Type: application/json" \
  -d '{...}' | python3 -m json.tool

# 2. 检查文档长度是否匹配
echo "文档字符数: $(cat document.txt | wc -c)"
```

**常见原因**：
- 文档在校对后被修改
- 偏移量计算错误
- 特殊字符编码问题

---

## 自动化测试脚本

创建 `run_tests.sh`：

```bash
#!/bin/bash

set -e

BASE_URL="http://127.0.0.1:$(cat /tmp/gowordagent-port.json | python3 -c "import json,sys; print(json.load(sys.stdin)['port'])")"

echo "=== GOWordAgent 自动化测试 ==="
echo "API Base URL: $BASE_URL"
echo ""

# 测试 1：健康检查
echo "[TEST 1] 健康检查..."
curl -s "$BASE_URL/api/proofread/health" | grep -q '"status":"ok"' && echo "✅ 通过" || echo "❌ 失败"

# 测试 2：配置获取
echo "[TEST 2] 获取配置..."
curl -s "$BASE_URL/api/proofread/config" | grep -q 'provider' && echo "✅ 通过" || echo "❌ 失败"

# 测试 3：校对接口（需要 API Key）
echo "[TEST 3] 校对接口..."
RESPONSE=$(curl -s -X POST "$BASE_URL/api/proofread" \
  -H "Content-Type: application/json" \
  -d '{
    "text": "测试文本",
    "paragraphs": [{"index":0,"startOffset":0,"endOffset":4,"text":"测试文本"}],
    "provider": "DeepSeek",
    "apiKey": "test-key",
    "mode": "Precise"
  }' 2>&1)

if echo "$RESPONSE" | grep -q '"success":true\|"error":'; then
    echo "✅ 通过 (有响应)"
else
    echo "❌ 失败: $RESPONSE"
fi

echo ""
echo "=== 测试完成 ==="
```

使用方法：
```bash
chmod +x run_tests.sh
./run_tests.sh
```

---

## 测试报告模板

```markdown
# 测试报告

## 基本信息
- 日期：2026-04-XX
- 测试人员：XXX
- 环境：银河麒麟 V10 SP1 (x86_64)
- WPS 版本：12.1.2.25838

## 测试结果

| 测试项 | 状态 | 备注 |
|--------|------|------|
| Day 0 - 环境侦察 | ✅/❌ | |
| Day 1 - HTTP PoC | ✅/❌ | |
| 后端启动 | ✅/❌ | |
| 加载项加载 | ✅/❌ | |
| 连接后端 | ✅/❌ | |
| 文本提取 | ✅/❌ | |
| 校对执行 | ✅/❌ | |
| 结果写回 | ✅/❌ | |
| 点击定位 | ✅/❌ | |

## 问题记录

1. 问题描述：...
   - 复现步骤：...
   - 期望结果：...
   - 实际结果：...
   - 解决方案：...

## 结论

- [ ] 测试通过，可以进入下一阶段
- [ ] 存在问题，需要修复后重新测试
- [ ] 需要切换到 B 计划 (LibreOffice)
```

---

*文档结束*
