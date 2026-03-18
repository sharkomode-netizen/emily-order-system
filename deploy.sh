#!/bin/bash
# EMILY 订单管理系统 — 腾讯云部署脚本
# 用法: ssh user@server 'bash -s' < deploy.sh
# 或:   scp deploy.sh user@server:~ && ssh user@server 'bash deploy.sh'

set -e

APP_DIR="/opt/emily-order-system"
APP_USER="emily"
PORT="${PORT:-5566}"

echo "=========================================="
echo "  EMILY 订单管理系统 — 部署到腾讯云"
echo "=========================================="

# 1. 系统依赖
echo "[1/6] 安装系统依赖..."
apt-get update -qq
apt-get install -y -qq python3 python3-pip python3-venv nginx supervisor

# 2. 创建应用用户和目录
echo "[2/6] 创建应用目录..."
id -u $APP_USER &>/dev/null || useradd -r -s /bin/false $APP_USER
mkdir -p $APP_DIR/{uploads,output,templates}
cp app.py $APP_DIR/ 2>/dev/null || echo "请先上传 app.py 到服务器"
cp -r templates/* $APP_DIR/templates/ 2>/dev/null || true

# 3. Python 虚拟环境
echo "[3/6] 配置Python环境..."
cd $APP_DIR
python3 -m venv .venv
.venv/bin/pip install -q flask openpyxl pdfplumber anthropic gunicorn pillow

# 4. 环境变量配置
echo "[4/6] 配置环境变量..."
cat > $APP_DIR/.env << 'ENVEOF'
# 必填：Anthropic API Key（云端无法用 Claude CLI）
ANTHROPIC_API_KEY=your-api-key-here
PORT=5566
FLASK_ENV=production
SECRET_KEY=$(python3 -c "import secrets; print(secrets.token_hex(32))")
ENVEOF

# 5. Supervisor 进程管理
echo "[5/6] 配置进程守护..."
cat > /etc/supervisor/conf.d/emily.conf << SUPEOF
[program:emily]
directory=$APP_DIR
command=$APP_DIR/.venv/bin/gunicorn -w 2 -b 0.0.0.0:$PORT --timeout 600 app:app
user=$APP_USER
autostart=true
autorestart=true
environment=ANTHROPIC_API_KEY="%(ENV_ANTHROPIC_API_KEY)s",FLASK_ENV="production",SECRET_KEY="emily-prod-key"
stdout_logfile=/var/log/emily.log
stderr_logfile=/var/log/emily-error.log
SUPEOF

# 6. Nginx 反代（支持多端口 + 域名）
echo "[6/6] 配置Nginx反代..."
cat > /etc/nginx/sites-available/emily << 'NGXEOF'
server {
    listen 80;
    listen 8080;
    listen 8888;
    server_name _;

    client_max_body_size 50M;

    location / {
        proxy_pass http://127.0.0.1:5566;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_read_timeout 600;
    }

    location /static {
        alias /opt/emily-order-system/static;
        expires 7d;
    }
}
NGXEOF

ln -sf /etc/nginx/sites-available/emily /etc/nginx/sites-enabled/
rm -f /etc/nginx/sites-enabled/default

# 设置权限
chown -R $APP_USER:$APP_USER $APP_DIR

# 重启服务
supervisorctl reread
supervisorctl update
supervisorctl restart emily
nginx -t && systemctl reload nginx

echo ""
echo "=========================================="
echo "  部署完成！"
echo "  访问: http://服务器IP:80"
echo "  或:   http://服务器IP:8080"
echo "  或:   http://服务器IP:8888"
echo ""
echo "  重要：请编辑 $APP_DIR/.env"
echo "  填入 ANTHROPIC_API_KEY"
echo "=========================================="
