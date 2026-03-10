/**
 * 日志控制台管理器
 */

class Logger {
    constructor() {
        this.entriesEl = document.getElementById('log-entries');
        this.bodyEl = document.getElementById('log-body');
        this.consoleEl = document.getElementById('log-console');
        this.clearBtn = document.getElementById('log-clear-btn');
        this.toggleBtn = document.getElementById('log-toggle-btn');

        this.clearBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.clear();
        });

        this.toggleBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.toggle();
        });

        // 点击 header 也可以 toggle
        document.querySelector('.log-header').addEventListener('click', () => {
            this.toggle();
        });
    }

    /**
     * 添加日志
     * @param {string} msg - 日志消息
     * @param {'info'|'success'|'warning'|'error'} level - 日志级别
     */
    log(msg, level = 'info') {
        const entry = document.createElement('div');
        entry.className = 'log-entry';

        const time = new Date().toLocaleTimeString('zh-CN', { hour12: false });

        entry.innerHTML = `
      <span class="log-time">${time}</span>
      <span class="log-msg ${level}">${this.escapeHtml(msg)}</span>
    `;

        this.entriesEl.appendChild(entry);
        // 自动滚动到底部
        this.bodyEl.scrollTop = this.bodyEl.scrollHeight;

        // 如果是折叠状态，展开
        if (this.consoleEl.classList.contains('collapsed')) {
            this.consoleEl.classList.remove('collapsed');
            this.toggleBtn.textContent = '▼';
        }
    }

    info(msg) { this.log(msg, 'info'); }
    success(msg) { this.log(msg, 'success'); }
    warn(msg) { this.log(msg, 'warning'); }
    error(msg) { this.log(msg, 'error'); }

    clear() {
        this.entriesEl.innerHTML = '';
    }

    toggle() {
        const collapsed = this.consoleEl.classList.toggle('collapsed');
        this.toggleBtn.textContent = collapsed ? '▲' : '▼';
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
}

// 单例
let instance = null;

export function getLogger() {
    if (!instance) {
        instance = new Logger();
    }
    return instance;
}
