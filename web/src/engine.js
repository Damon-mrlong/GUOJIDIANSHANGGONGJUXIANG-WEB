/**
 * Pyodide 引擎管理器
 * 负责加载 Pyodide、安装依赖包、执行 Python 代码
 */

const PYODIDE_CDN = 'https://cdn.jsdelivr.net/pyodide/v0.27.0/full/';

let pyodide = null;
let isReady = false;

/**
 * 初始化 Pyodide 运行时
 * @param {Function} onProgress - 进度回调 (percent, message)
 * @returns {Promise<void>}
 */
export async function initPyodide(onProgress = () => { }) {
    if (isReady) return;

    onProgress(10, '加载 Pyodide 运行时...');

    // 动态加载 Pyodide 脚本
    await loadScript(`${PYODIDE_CDN}pyodide.js`);

    onProgress(30, '初始化 Python 环境...');

    pyodide = await globalThis.loadPyodide({
        indexURL: PYODIDE_CDN,
    });

    onProgress(50, '安装 Pandas...');
    await pyodide.loadPackage('pandas');

    onProgress(70, '安装 Openpyxl...');
    await pyodide.loadPackage('openpyxl');

    onProgress(85, '加载处理引擎...');
    // 注册 Python 引擎代码
    await registerEngines();

    onProgress(100, '就绪');
    isReady = true;
}

/**
 * 动态加载外部脚本
 */
function loadScript(src) {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = resolve;
        script.onerror = () => reject(new Error(`加载脚本失败: ${src}`));
        document.head.appendChild(script);
    });
}

/**
 * 注册所有 Python 引擎代码到 Pyodide
 */
async function registerEngines() {
    // 注入桥接辅助代码
    await pyodide.runPythonAsync(`
import pandas as pd
import io
import json
import os
from datetime import datetime

# 全局日志回调（由 JS 设置）
_log_callback = None

def set_log_callback(cb):
    global _log_callback
    _log_callback = cb

def _log(msg):
    if _log_callback:
        _log_callback(str(msg))
    else:
        print(msg)
  `);
}

/**
 * 将 JS File 对象转为 Pyodide 可用的 bytes
 * @param {File} file
 * @returns {Promise<Uint8Array>}
 */
export async function fileToBytes(file) {
    const buffer = await file.arrayBuffer();
    return new Uint8Array(buffer);
}

/**
 * 在 Pyodide 中执行 Python 代码
 * @param {string} code
 * @returns {Promise<any>}
 */
export async function runPython(code) {
    if (!isReady) throw new Error('Pyodide 尚未初始化');
    return await pyodide.runPythonAsync(code);
}

/**
 * 获取 pyodide 实例
 */
export function getPyodide() {
    return pyodide;
}

/**
 * 设置 Python 端日志回调
 * @param {Function} callback
 */
export function setLogCallback(callback) {
    if (!pyodide) return;
    pyodide.globals.set('_log_callback', callback);
}

/**
 * 将文件加载到 Pyodide 虚拟文件系统
 * @param {File} file - JS File 对象
 * @param {string} path - 虚拟文件系统中的路径
 */
export async function loadFileToFS(file, path) {
    const bytes = await fileToBytes(file);
    pyodide.FS.writeFile(path, bytes);
}

/**
 * 从 Pyodide 虚拟文件系统读取文件
 * @param {string} path
 * @returns {Uint8Array}
 */
export function readFileFromFS(path) {
    return pyodide.FS.readFile(path);
}

/**
 * 触发浏览器下载
 * @param {Uint8Array} data
 * @param {string} filename
 */
export function downloadFile(data, filename) {
    const blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

/**
 * 检查引擎是否就绪
 */
export function isEngineReady() {
    return isReady;
}
