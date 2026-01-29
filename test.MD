# 文件防篡改系统 v2.5

**File-Anti-Tampering-System v2.5 Strict Log Edition**

基于 Linux 底层属性（`chattr`）与 Python `watchdog` 实时监控开发的轻量级文件防篡改与审计系统，专为 **CentOS 7/8/9** 及 **宝塔面板（BT Panel）** 环境设计，提供可视化的 Web 管理界面。

> **版本特性（v2.5）**：采用「精简审计」策略——系统**仅记录**受保护且未解锁文件的异常变动（新建、修改、删除、重命名）。手动解锁（白名单）的文件操作及非受保护后缀的文件变动将自动过滤，不再产生冗余日志。

---

## 一、核心功能

### 1.1 底层强力锁定

利用 Linux 文件系统 `i` 属性（Immutable）锁定文件，即使是 Root 用户也无法直接修改或删除，必须先通过系统解锁。

### 1.2 可视化文件管理

提供 Web 界面浏览目录，支持一键**锁定/解锁**单个文件或整个目录。

### 1.3 智能日志审计（v2.5 核心）

- **自动过滤**：后缀不在保护列表的文件（如 `.log`、`.tmp`）不记录。
- **白名单过滤**：手动解锁的文件，在编辑期间不记录日志。
- **违规记录**：仅记录处于锁定状态下的受保护文件发生的违规变动。

### 1.4 安全认证

后台登录集成 **Google 2FA（双因素认证）**，首次登录强制绑定，保障管理后台安全。

### 1.5 自动化运维

- **配置热更新**：修改配置后自动应用。
- **日志自动清理**：可设置保留天数，系统每日自动清理过期日志。

---

## 二、快速安装

### 2.1 运行安装脚本

将安装脚本保存为 `install.sh`，赋予权限并运行（需 Root 权限）：

```bash
chmod +x install.sh
./install.sh

2.2 登录后台

• 访问地址：http://<服务器IP>:18888
• 默认密码：admin123
2.3 初始化配置

---
三、使用指南
3.1 如何修改受保护的文件
为了避免触发报警日志，请遵循以下流程：

3.2 配置说明
在系统配置页面可以调整以下参数：

| 配置项 | 说明 |
| --- | --- |
| 受保护后缀 | 定义哪些文件需要被锁定（如 .php、.js、.html）。只有在此列表中的文件后缀发生变动时，系统才会记录日志。非受保护后缀（如 .log）的变动会被自动忽略。 |
| 日志保留天数 | 定义审计日志在数据库中保存的时间（默认 7 天）。系统每 24 小时会自动检查并清理超过该天数的旧日志。 |

---
四、系统维护
4.1 目录结构

| 项目 | 路径 |
| --- | --- |
| 程序目录 | /www/server/tamper_protect |
| 日志文件 | /www/server/tamper_protect/logs/tamper_protect.log（程序运行调试日志） |
| 数据库 | /www/server/tamper_protect/data/tamper_protect.db（SQLite 存储配置与审计记录） |
| 配置文件 | /www/server/tamper_protect/data/config.json |

4.2 常用管理命令

bash# 启动服务
systemctl start tamper-protect

# 停止服务
systemctl stop tamper-protect

# 重启服务
systemctl restart tamper-protect

# 查看运行状态
systemctl status tamper-protect

# 查看实时运行日志（用于排查故障）
tail -f /www/server/tamper_protect/logs/tamper_protect.log

---
五、卸载说明
如果您决定不再使用该系统，请严格按照以下步骤卸载，以免文件仍处于锁定状态：

bash# 1. 停止并禁用服务
systemctl stop tamper-protect
systemctl disable tamper-protect

# 2. 删除服务配置
rm -f /etc/systemd/system/tamper-protect.service
systemctl daemon-reload

# 3. 删除程序文件
rm -rf /www/server/tamper_protect

# 4.（重要）手动解锁文件
# 卸载程序不会自动解锁已保护的文件。您必须手动执行以下命令解锁：
# 请将 /www/wwwroot/您的网站目录 替换为实际路径
chattr -R -i /www/wwwroot/您的网站目录

---
六、免责声明

```
