#!/bin/bash
# ============================================
# 国际电商工具箱 - 一键部署脚本
# 在服务器上执行: bash deploy.sh
# ============================================

set -e

echo "=========================================="
echo "  国际电商工具箱 - 一键部署"
echo "=========================================="

# 0. 设置访问密码
read -p "请设置访问用户名 (默认 admin): " AUTH_USER
AUTH_USER=${AUTH_USER:-admin}
read -s -p "请设置访问密码: " AUTH_PASS
echo ""
if [ -z "$AUTH_PASS" ]; then
  echo "❌ 密码不能为空！"
  exit 1
fi

# 1. 安装 nginx、git 和密码工具
echo "[1/5] 安装依赖..."
apt update -y
apt install -y nginx git apache2-utils

# 2. 创建密码文件
echo "[2/5] 配置访问密码..."
htpasswd -bc /etc/nginx/.htpasswd "$AUTH_USER" "$AUTH_PASS"
echo "✅ 用户 $AUTH_USER 密码已设置"

# 3. 克隆代码
echo "[3/5] 拉取代码..."
DEPLOY_DIR="/var/www/toolbox"
if [ -d "$DEPLOY_DIR/.git" ]; then
  echo "目录已存在，拉取最新代码..."
  cd "$DEPLOY_DIR"
  git pull origin main
else
  rm -rf "$DEPLOY_DIR"
  git clone https://github.com/Damon-mrlong/GUOJIDIANSHANGGONGJUXIANG-WEB.git "$DEPLOY_DIR"
fi

# 4. 配置 nginx
echo "[4/5] 配置 nginx..."
cat > /etc/nginx/sites-available/toolbox << 'EOF'
server {
    listen 80;
    server_name _;

    root /var/www/toolbox;
    index web/index.html;

    # 密码认证
    auth_basic "国际电商工具箱 - 请输入访问凭证";
    auth_basic_user_file /etc/nginx/.htpasswd;

    # 中文路径支持
    charset utf-8;

    # 访问根路径自动跳转到 web/
    location = / {
        return 301 /web/;
    }

    # 静态资源
    location / {
        try_files $uri $uri/ =404;
        add_header Cache-Control "public, max-age=3600";
    }

    # 关闭目录列表
    autoindex off;

    # 大文件限制
    client_max_body_size 50M;

    # Gzip 压缩
    gzip on;
    gzip_types text/plain text/css application/javascript application/json text/xml;
    gzip_min_length 1000;
}
EOF

# 启用站点配置
ln -sf /etc/nginx/sites-available/toolbox /etc/nginx/sites-enabled/toolbox
rm -f /etc/nginx/sites-enabled/default

# 5. 配置防火墙 + 启动 nginx
echo "[5/5] 启动服务..."
nginx -t
systemctl restart nginx
systemctl enable nginx

# 防火墙（如果 ufw 可用）
if command -v ufw &> /dev/null; then
  ufw allow 22/tcp
  ufw allow 80/tcp
  echo "y" | ufw enable 2>/dev/null || true
  echo "✅ 防火墙已配置（仅开放 22、80 端口）"
fi

echo ""
echo "=========================================="
echo "  ✅ 部署完成！"
echo "  访问地址: http://163.7.12.28/web/"
echo "  用户名: $AUTH_USER"
echo "  密码: (你刚才设置的密码)"
echo "=========================================="
echo ""
echo "后续操作:"
echo "  更新代码: cd /var/www/toolbox && git pull origin main"
echo "  修改密码: htpasswd /etc/nginx/.htpasswd $AUTH_USER"
echo "  添加用户: htpasswd /etc/nginx/.htpasswd 新用户名"
