// ==UserScript==
// @name XHS_Comment_Collector
// @version 1.2
// @description 小红书评论辅助收集脚本，在用户浏览时，通过捕获评论接口，进行评论内容获取，支持二级评论获取
// @author TOBEFREE
// @match https://www.xiaohongshu.com/*
// @run-at document-start
// @grant none
// @downloadURL https://github.com/TOTOBEFREE/XHS_Comment_Collector/releases/latest/download/xhs_comment_collector.js
// @updateURL https://github.com/TOTOBEFREE/XHS_Comment_Collector/releases/latest/download/xhs_comment_collector.js
// @require https://npm.elemecdn.com/xlsx/dist/xlsx.full.min.js
// ==/UserScript==


(() => {  
  "use strict";

  // ---------- 配置 ----------
  const MAIN_COMMENT_RE = /https:\/\/edith\.xiaohongshu\.com\/api\/sns\/web\/v2\/comment\/page\b/i;
  const SUB_COMMENT_RE = /https:\/\/edith\.xiaohongshu\.com\/api\/sns\/web\/v2\/comment\/sub\/page\b/i;
  const TITLE_SELECTOR = "#detail-title";
  const DESC_TITLE_SELECTOR = "#detail-desc .note-text span";
  const DESC_SELECTOR = "#detail-desc";
  const COVER_SELECTOR = ".xhs-slider-container .note-slider-img img";
  const STORAGE_KEY_REPLY_CAPTURE = "__xhs_reply_capture_enabled__";
  const STORAGE_KEY_AUTO_EXPAND = "__xhs_auto_expand_enabled__";
  const STORAGE_KEY_SELECTED_TITLE = "__xhs_selected_title__";
  const DB_NAME = "xhs-comment-collector";
  const DB_VERSION = 1;
  const COMMENTS_STORE = "comments";
  const META_STORE = "meta";
  const META_LAST_RESTORE_AT = "lastRestoreAt";
  const PERSIST_BATCH_SIZE = 20;
  const PERSIST_DELAY = 2000;
  const DEFAULT_TIP_HTML = "⚡ 自动捕获主评；可选抓取追评（自动展开回复）<br>";
  const XLSX_CDN_LIST = [
    "https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js",
    "https://unpkg.com/xlsx/dist/xlsx.full.min.js"
  ];

  // ---------- 数据存储 ----------
  const records = [];                 // 所有评论（含主评和子评）
  const seenKey = new Set();           // 去重集合
  let mainCommentCount = 0;            // 主评论数量
  let expandTimer = null;              // 展开按钮轮询计时器
  let expandIdleTicks = 0;             // 连续空轮询次数
  let pendingPersistQueue = [];
  let persistTimer = null;
  let isPersisting = false;
  let restoreInProgress = false;
  let persistDisabled = false;
  let lastRestoreCount = 0;
  let noteMetaCache = null;

  function isVisibleAndClickableElement(el) {
    return el instanceof HTMLElement
      && el.offsetParent !== null
      && el.getAttribute("aria-disabled") !== "true"
      && !el.hasAttribute("disabled");
  }

  function clearPersistTimer() {
    if (!persistTimer) return;
    clearTimeout(persistTimer);
    persistTimer = null;
  }

  function removePersistedBatch(batch) {
    const savedKeys = new Set(batch.map((item) => item.record_key));
    pendingPersistQueue = pendingPersistQueue.filter((item) => !savedKeys.has(item.record_key));
  }

  function buildSelectOptionsHtml(titles, selectedTitle, titleCounts) {
    const defaultOption = '<option value="">📌 选择文章导出</option>';
    const titleOptions = titles
      .map((title) => {
        const selected = title === selectedTitle ? " selected" : "";
        const count = titleCounts.get(title) || 0;
        return `<option value="${escapeHtml(title)}"${selected}>${escapeHtml(title)} (${count})</option>`;
      })
      .join("");

    return defaultOption + titleOptions;
  }

  function collectTitleCounts() {
    const titleCounts = new Map();
    for (const record of records) {
      titleCounts.set(record.title, (titleCounts.get(record.title) || 0) + 1);
    }
    return titleCounts;
  }

  function resetRuntimeRecords() {
    records.length = 0;
    seenKey.clear();
    pendingPersistQueue = [];
    clearPersistTimer();
    mainCommentCount = 0;
    lastRestoreCount = 0;
    noteMetaCache = null;
  }

  const storage = {
    dbPromise: null,

    init() {
      if (persistDisabled) return Promise.resolve(null);
      if (this.dbPromise) return this.dbPromise;

      this.dbPromise = new Promise((resolve) => {
        try {
          const request = indexedDB.open(DB_NAME, DB_VERSION);
          request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(COMMENTS_STORE)) {
              const commentsStore = db.createObjectStore(COMMENTS_STORE, { keyPath: "record_key" });
              commentsStore.createIndex("title", "title", { unique: false });
              commentsStore.createIndex("captured_at", "captured_at", { unique: false });
            }
            if (!db.objectStoreNames.contains(META_STORE)) {
              db.createObjectStore(META_STORE, { keyPath: "key" });
            }
          };
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => {
            persistDisabled = true;
            console.warn("[xhs] indexedDB 初始化失败，将降级为纯内存模式", request.error);
            resolve(null);
          };
        } catch (e) {
          persistDisabled = true;
          console.warn("[xhs] indexedDB 不可用，将降级为纯内存模式", e);
          resolve(null);
        }
      });

      return this.dbPromise;
    },

    async withStore(storeName, mode, handler) {
      const db = await this.init();
      if (!db) return null;

      return new Promise((resolve, reject) => {
        try {
          const tx = db.transaction(storeName, mode);
          const store = tx.objectStore(storeName);
          const result = handler(store, tx);
          tx.oncomplete = () => resolve(result ?? null);
          tx.onerror = () => reject(tx.error || new Error("indexedDB transaction failed"));
          tx.onabort = () => reject(tx.error || new Error("indexedDB transaction aborted"));
        } catch (e) {
          reject(e);
        }
      });
    },

    async saveCommentsBatch(batch) {
      if (!Array.isArray(batch) || batch.length === 0) return;
      await this.withStore(COMMENTS_STORE, "readwrite", (store) => {
        for (const item of batch) {
          store.put(item);
        }
      });
    },

    async loadAllComments() {
      const db = await this.init();
      if (!db) return [];

      return new Promise((resolve, reject) => {
        try {
          const tx = db.transaction(COMMENTS_STORE, "readonly");
          const store = tx.objectStore(COMMENTS_STORE);
          const request = store.getAll();
          request.onsuccess = () => resolve(Array.isArray(request.result) ? request.result : []);
          request.onerror = () => reject(request.error || new Error("indexedDB getAll failed"));
        } catch (e) {
          reject(e);
        }
      });
    },

    async clearAllComments() {
      await this.withStore(COMMENTS_STORE, "readwrite", (store) => store.clear());
    },

    async setMeta(key, value) {
      await this.withStore(META_STORE, "readwrite", (store) => {
        store.put({ key, value });
      });
    }
  };

  // ---------- 工具函数 ----------
  const pad2 = (n) => String(n).padStart(2, "0");
  const formatTs = (ms) => {
    const n = typeof ms === "string" ? Number(ms) : ms;
    if (!Number.isFinite(n)) return "";
    const d = new Date(n);
    return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
  };

  const formatDateForFilename = (ms) => {
    const d = new Date(ms);
    return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
  };

  const safeSheetName = (name, fallback = "Sheet") => {
    const raw = String(name || "").trim() || fallback;
    const cleaned = raw.replace(/[\\/*?:\[\]]/g, " ").replace(/\s+/g, " ").trim() || fallback;
    return cleaned.slice(0, 31);
  };

  const normalizeUrl = (u) => {
    try {
      if (typeof u === "string") return new URL(u, location.href).href;
      if (u && typeof u.url === "string") return new URL(u.url, location.href).href;
    } catch {}
    return "";
  };

  const getApiType = (url) => {
    if (SUB_COMMENT_RE.test(url || "")) return "sub";
    if (MAIN_COMMENT_RE.test(url || "")) return "main";
    return null;
  };

  const shouldHandleApiUrl = (url) => {
    const apiType = getApiType(url);
    if (apiType === "main") return true;
    if (apiType === "sub") return ui.isReplyCaptureEnabled();
    return false;
  };

  const shouldAutoExpand = () => ui.isAutoExpandEnabled();

  const parseParentIdFromSubUrl = (url) => {
    try {
      const u = new URL(url, location.href);
      const keys = ["root_comment_id", "comment_id", "target_comment_id", "top_comment_id", "parent_comment_id"];
      for (const k of keys) {
        const v = u.searchParams.get(k);
        if (v) return v;
      }
    } catch {}
    return null;
  };

  const pickIpLocation = (c) => c?.ip_location ?? c?.ipLocation ?? c?.ip_location_str ?? c?.ipLocationStr ?? "";

  const getCurrentTitle = () => {
    const primary = readText(TITLE_SELECTOR);
    if (primary) return primary;
    const fallback = readText(DESC_TITLE_SELECTOR);
    return fallback || "未知标题";
  };

  const getCurrentDesc = () => readText(DESC_SELECTOR);

  const getCurrentCoverImage = () => {
    const primary = document.querySelector(COVER_SELECTOR)?.src || "";
    if (primary) return primary;

    try {
      const bgImage = document.querySelector(".xgplayer-poster")?.style?.backgroundImage || "";
      const match = bgImage.match(/url\(["']?(.*?)["']?\)/);
      return match?.[1] || "";
    } catch {}

    return "";
  };

  const getCurrentNoteMeta = () => {
    const pageKey = getPageCacheKey();
    if (noteMetaCache && noteMetaCache.pageKey === pageKey) {
      return noteMetaCache.value;
    }

    const value = {
      title: getCurrentTitle(),
      desc: getCurrentDesc(),
      coverImage: getCurrentCoverImage(),
    };
    noteMetaCache = { pageKey, value };
    return value;
  };

  const escapeHtml = (s) => String(s).replace(/[&<>"']/g, (m) => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#039;" })[m]);

  const getPageCacheKey = () => location.href.split("#")[0];

  const readText = (selector) => {
    const el = document.querySelector(selector);
    if (!el) return "";
    const text = typeof el.textContent === "string" ? el.textContent : el.innerText;
    return String(text || "").trim();
  };

  function buildRecordKey({ title, text, createTimeRaw, ipLocation, commentId, parentCommentId, targetCommentId, dedupeSuffix = "" }) {
    return commentId
      ? `id:${commentId}`
      : `${title}__${text}__${createTimeRaw || ""}__${ipLocation || ""}__parent:${parentCommentId || ""}__target:${targetCommentId || ""}${dedupeSuffix}`;
  }

  function buildRecordPayload({ title, noteDesc, coverImage, content, createTimeRaw, ipLocation, commentId, parentCommentId, targetCommentId, dedupeSuffix = "", capturedAt }) {
    const text = (content ?? "").trim();
    if (!text) return null;

    const recordKey = buildRecordKey({
      title,
      text,
      createTimeRaw,
      ipLocation,
      commentId,
      parentCommentId,
      targetCommentId,
      dedupeSuffix,
    });

    return {
      record_key: recordKey,
      title,
      note_desc: String(noteDesc ?? ""),
      cover_image: String(coverImage ?? ""),
      id: commentId || null,
      content: text,
      create_time: formatTs(createTimeRaw),
      create_time_raw: createTimeRaw,
      captured_at: capturedAt ?? Date.now(),
      ip_location: String(ipLocation ?? ""),
      commentId: commentId || null,
      parentCommentId: parentCommentId || null,
      targetCommentId: targetCommentId || null,
    };
  }

  function buildMainCommentPayload(comment, noteMeta) {
    return {
      title: noteMeta.title,
      noteDesc: noteMeta.desc,
      coverImage: noteMeta.coverImage,
      content: comment?.content ?? "",
      createTimeRaw: comment?.create_time ?? comment?.createTime ?? "",
      ipLocation: String(pickIpLocation(comment) ?? ""),
      commentId: comment?.id ?? comment?.comment_id ?? null,
      parentCommentId: null,
      targetCommentId: null,
      dedupeSuffix: "__main",
    };
  }

  function buildSubCommentPayload(comment, noteMeta, fallbackParentId, dedupeSuffix) {
    const targetCommentId = comment?.target_comment?.id ?? fallbackParentId ?? null;
    return {
      title: noteMeta.title,
      noteDesc: noteMeta.desc,
      coverImage: noteMeta.coverImage,
      content: comment?.content ?? "",
      createTimeRaw: comment?.create_time ?? comment?.createTime ?? "",
      ipLocation: String(pickIpLocation(comment) ?? ""),
      commentId: comment?.id ?? comment?.comment_id ?? null,
      parentCommentId: targetCommentId ?? fallbackParentId ?? null,
      targetCommentId,
      dedupeSuffix,
    };
  }

  function hydrateSeenKeyFromRecords() {
    seenKey.clear();
    for (const record of records) {
      const key = record?.record_key || buildRecordKey({
        title: record?.title || "",
        text: record?.content || "",
        createTimeRaw: record?.create_time_raw,
        ipLocation: record?.ip_location,
        commentId: record?.commentId || record?.id,
        parentCommentId: record?.parentCommentId,
        targetCommentId: record?.targetCommentId,
      });
      if (key) {
        record.record_key = key;
        seenKey.add(key);
      }
    }
  }

  function schedulePersist(force = false) {
    if (persistDisabled || restoreInProgress || pendingPersistQueue.length === 0) return;
    if (force || pendingPersistQueue.length >= PERSIST_BATCH_SIZE) {
      flushPendingRecords(force);
      return;
    }
    if (persistTimer) return;
    persistTimer = setTimeout(() => {
      persistTimer = null;
      flushPendingRecords();
    }, PERSIST_DELAY);
  }

  async function flushPendingRecords(force = false) {
    if (persistDisabled || restoreInProgress) return;
    clearPersistTimer();
    if (isPersisting || pendingPersistQueue.length === 0) return;
    isPersisting = true;

    const batch = pendingPersistQueue.slice();
    try {
      await storage.saveCommentsBatch(batch);
      removePersistedBatch(batch);
      if (force) {
        await storage.setMeta(META_LAST_RESTORE_AT, Date.now());
      }
    } catch (e) {
      console.warn("[xhs] 评论备份失败，将稍后重试", e);
    } finally {
      isPersisting = false;
      if (pendingPersistQueue.length > 0) {
        schedulePersist();
      }
    }
  }

  const truncateTitle = (title, maxLength = 18) => {
    if (!title) return "";
    return title.length > maxLength ? title.slice(0, maxLength) + "..." : title;
  };

  function findNextShowMoreButton() {
    const btns = document.querySelectorAll(".show-more");
    for (const btn of btns) {
      if (!isVisibleAndClickableElement(btn)) continue;
      return btn;
    }
    return null;
  }

  function clickNextShowMoreButton() {
    const btn = findNextShowMoreButton();
    if (!btn) return false;
    btn.click();
    return true;
  }

  function getRandomExpandDelay() {
    return 2000 + Math.floor(Math.random() * 3000);
  }

  function scheduleNextExpand() {
    if (expandTimer || !shouldAutoExpand()) return;
    expandTimer = setTimeout(() => {
      expandTimer = null;

      if (!shouldAutoExpand()) {
        stopExpandLoop();
        return;
      }

      const clicked = clickNextShowMoreButton();
      if (clicked) {
        expandIdleTicks = 0;
        scheduleNextExpand();
        return;
      }

      expandIdleTicks++;
      if (expandIdleTicks < 5) {
        scheduleNextExpand();
      }
    }, getRandomExpandDelay());
  }

  function startExpandLoop() {
    if (!shouldAutoExpand()) return;
    expandIdleTicks = 0;
    const clicked = clickNextShowMoreButton();
    if (clicked) {
      scheduleNextExpand();
      return;
    }
    scheduleNextExpand();
  }

  function stopExpandLoop() {
    if (expandTimer) {
      clearTimeout(expandTimer);
      expandTimer = null;
    }
    expandIdleTicks = 0;
  }

  function addCommentRecord({ title, noteDesc, coverImage, content, createTimeRaw, ipLocation, commentId, parentCommentId, targetCommentId, dedupeSuffix = "" }) {
    const record = buildRecordPayload({
      title,
      noteDesc,
      coverImage,
      content,
      createTimeRaw,
      ipLocation,
      commentId,
      parentCommentId,
      targetCommentId,
      dedupeSuffix,
    });
    if (!record) return false;
    if (seenKey.has(record.record_key)) return false;

    seenKey.add(record.record_key);
    records.push(record);

    if (!restoreInProgress && !persistDisabled) {
      pendingPersistQueue.push(record);
      schedulePersist();
    }
    return true;
  }

  // ---------- 处理接口响应（支持追评）----------
  function handleResultJson(url, result) {
    const apiType = getApiType(url);
    if (!apiType) return;
    if (apiType === "sub" && !ui.isReplyCaptureEnabled()) return;

    const comments = result?.data?.comments;
    if (!Array.isArray(comments) || comments.length === 0) return;

    const noteMeta = getCurrentNoteMeta();
    const title = noteMeta.title;
    let added = 0;
    let newMainAdded = 0;
    const subUrlParentId = apiType === "sub" ? parseParentIdFromSubUrl(url) : null;

    for (const c of comments) {
      if (apiType === "main") {
        const mainPayload = buildMainCommentPayload(c, noteMeta);
        const mainId = mainPayload.commentId;

        if (addCommentRecord(mainPayload)) {
          mainCommentCount++;
          newMainAdded++;
          added++;
        }

        const subComments = c?.sub_comments;
        if (ui.isReplyCaptureEnabled() && Array.isArray(subComments) && subComments.length > 0) {
          for (const sub of subComments) {
            const subPayload = buildSubCommentPayload(sub, noteMeta, mainId, "__sub_main_api");
            if (addCommentRecord(subPayload)) {
              added++;
            }
          }
        }
      } else {
        const subPayload = buildSubCommentPayload(c, noteMeta, subUrlParentId, "__sub_api");
        if (addCommentRecord(subPayload)) {
          added++;
        }
      }
    }

    if (newMainAdded > 0) {
      startExpandLoop();
    }

    if (added > 0) {
      ui.scheduleUpdate();
    }
  }

  // ---------- UI 面板 ----------
  const ui = {
    root: null,
    shadow: null,
    els: {},
    xlsxLoading: false,
    selectedTitle: "",
    selectionInitialized: false,
    prevTitleCounts: new Map(),
    previewPageSize: 20,
    previewRenderCount: 20,
    previewLastTotal: 0,
    updateScheduled: false,
    statusMessage: "",

    setStatus(message) {
      this.statusMessage = message || "";
      if (this.els.tip) {
        this.els.tip.innerHTML = this.statusMessage || DEFAULT_TIP_HTML;
      }
    },

    scheduleUpdate() {
      if (this.updateScheduled) return;
      this.updateScheduled = true;
      const runner = () => {
        this.updateScheduled = false;
        this.update();
      };
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(runner);
      } else {
        setTimeout(runner, 16);
      }
    },

    async restorePersistedRecords() {
      if (persistDisabled) return;
      restoreInProgress = true;
      try {
        const persisted = await storage.loadAllComments();
        if (!Array.isArray(persisted) || persisted.length === 0) return;

        records.length = 0;
        for (const item of persisted) {
          records.push(item);
        }
        lastRestoreCount = persisted.length;
        hydrateSeenKeyFromRecords();
        this.setStatus(`⚡ 已自动恢复 ${persisted.length} 条本地备份评论`);
      } catch (e) {
        console.warn("[xhs] 本地评论恢复失败", e);
        this.setStatus("⚡ 本地备份恢复失败，已继续使用内存模式");
      } finally {
        restoreInProgress = false;
        this.setStatus(lastRestoreCount > 0 ? `⚡ 已自动恢复 ${lastRestoreCount} 条本地备份评论` : "");
        this.scheduleUpdate();
      }
    },

    init() {
      const mount = () => {
        if (this.root) return;

        const host = document.createElement("div");
        host.id = "__xhs_comment_spy_panel__";
        host.style.cssText = `
          position: fixed;
          right: 0.875rem;
          top: 0.875rem;
          z-index: 2147483647;
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'PingFang SC', 'Microsoft YaHei', sans-serif;
        `;

        const shadow = host.attachShadow({ mode: "open" });
        shadow.innerHTML = `
          <style>
            * { box-sizing: border-box; }
            .panel {
              width: 23.75rem;
              max-width: calc(100vw - 1.75rem);
              background: #fff7f7;
              color: #3f2a2a;
              border: 0.0625rem solid #f3cfd1;
              border-radius: 1rem;
              overflow: hidden;
              box-shadow: 0 0.75rem 1.75rem rgba(156, 23, 49, 0.14);
            }
            .header {
              display: flex;
              align-items: center;
              justify-content: space-between;
              padding: 0.75rem 1rem;
              border-bottom: 0.0625rem solid #f2d8db;
              background: linear-gradient(180deg, #ffeef0 0%, #fff7f7 100%);
            }
            .title-area {
              display: flex;
              flex-direction: column;
            }
            .title-main {
              font-size: 1rem;
              font-weight: 600;
              letter-spacing: 0.01875rem;
            }
            .title-count {
              font-size: 0.8125rem;
              color: #a16b75;
              margin-top: 0.125rem;
            }
            .btn-group {
              display: flex;
              gap: 0.5rem;
            }
            button, select {
              background: #fff;
              border: 0.0625rem solid #e9c6cb;
              color: #4a2c31;
              border-radius: 0.625rem;
              padding: 0.5rem 0.75rem;
              font-size: 0.8125rem;
              cursor: pointer;
              transition: all 0.2s ease;
              outline: none;
            }
            button {
              font-weight: 600;
            }
            button:hover:not(:disabled) {
              border-color: #d897a1;
              background: #fff2f4;
            }
            button:disabled {
              opacity: 0.45;
              cursor: not-allowed;
            }
            select {
              min-width: 8.75rem;
              width: 100%;
              background: #fff;
            }
            select option {
              background: #fff;
              color: #4a2c31;
            }
            .content {
              padding: 0.75rem 1rem 1rem;
              background: linear-gradient(180deg, #fff9f9 0%, #fff3f4 100%);
            }
            .select-row {
              margin-bottom: 0.625rem;
            }
            .option-row {
              margin-bottom: 0.625rem;
            }
            .check {
              display: inline-flex;
              align-items: center;
              gap: 0.375rem;
              font-size: 0.8125rem;
              color: #5b3840;
              user-select: none;
            }
            .check input {
              accent-color: #d95470;
              cursor: pointer;
            }
            .action-row {
              display: flex;
              gap: 0.5rem;
              margin-top: 0.75rem;
            }
            .action-row button {
              flex: 1;
            }
            .tip {
              font-size: 0.75rem;
              color: #94626c;
              margin: 0.375rem 0 0.625rem;
              line-height: 1.5;
              background: #fff1f3;
              border: 0.0625rem solid #f3d5d9;
              padding: 0.5rem 0.625rem;
              border-radius: 0.625rem;
            }
            .list {
              max-height: 23rem;
              overflow-y: auto;
              border: 0.0625rem solid #f1d9dd;
              border-radius: 0.625rem;
              padding: 0.5rem;
              background: #fff;
            }
            .item {
              padding: 0.5rem 0.625rem;
              border-bottom: 0.0625rem dashed #f0d5da;
            }
            .item.is-sub {
              margin-left: 0.875rem;
              padding-left: 0.75rem;
              border-left: 0.125rem solid #f1d5da;
              background: #fff9fa;
            }
            .item:last-child { border-bottom: none; }
            .meta {
              font-size: 0.75rem;
              color: #91666e;
              margin-bottom: 0.25rem;
              display: flex;
              gap: 0.5rem;
              flex-wrap: wrap;
            }
            .title-tag {
              background: #ffe6ea;
              padding: 0.125rem 0.5rem;
              border-radius: 999px;
              color: #7a3b47;
            }
            .sub-indicator {
              color: #b2868f;
              margin-left: 0.25rem;
            }
            .content-text {
              font-size: 0.8125rem;
              white-space: pre-wrap;
              word-break: break-word;
              color: #3f2a2a;
            }
            .btn-primary {
              background: #d95470;
              color: #fff;
              border-color: #d95470;
            }
            .btn-primary:hover:not(:disabled) {
              background: #c74762;
              border-color: #c74762;
            }
            .mini-expand {
              display: none;
              width: 100%;
              height: 100%;
              align-items: center;
              justify-content: center;
              background: transparent;
              border: none;
              color: #b8405b;
              font-size: 0;
              line-height: 1;
              cursor: pointer;
            }
            .mini-expand svg {
              width: 1rem;
              height: 1rem;
              stroke: currentColor;
            }
            .mini-expand:hover {
              background: #ffe8ec;
            }
            .panel.is-minimized .content {
              display: none;
            }
            .panel.is-minimized .header {
              display: none;
            }
            .panel.is-minimized {
              width: 2.75rem;
              height: 2.25rem;
              border-radius: 0.75rem;
            }
            .panel.is-minimized .mini-expand {
              display: flex;
            }
          </style>

          <div class="panel">
            <button id="miniExpand" class="mini-expand" aria-label="展开面板">
              <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M8 3H3v5"></path>
                <path d="M16 3h5v5"></path>
                <path d="M8 21H3v-5"></path>
                <path d="M16 21h5v-5"></path>
              </svg>
            </button>
            <div class="header">
              <div class="title-area">
                <div class="title-main">XHS 评论抓取</div>
                <div class="title-count" id="count">0 条评论</div>
              </div>
              <div class="btn-group">
                <button id="clear">清空</button>
                <button id="hide">最小化</button>
              </div>
            </div>

            <div class="content">
              <div class="tip" id="tip">
                ⚡ 自动捕获主评；可选抓取追评（自动展开回复）<br>
              </div>

              <div class="option-row" style="display: flex; gap: 1rem;">
                <label class="check">
                  <input id="replyCaptureToggle" type="checkbox" checked>
                  抓取追评
                </label>
                <label class="check">
                  <input id="autoExpandToggle" type="checkbox">
                  自动展开追评
                </label>
              </div>

              <div class="select-row">
                <select id="articleSelect">
                  <option value="">📌 选择文章导出</option>
                </select>
              </div>

              <div class="list" id="list"></div>

              <div class="action-row">
                <button id="exportSelected" disabled>导出所选</button>
                <button id="exportAll" class="btn-primary" disabled>全部导出</button>
              </div>
            </div>
          </div>
        `;

        this.root = host;
        this.shadow = shadow;
        this.els.count = shadow.getElementById("count");
        this.els.list = shadow.getElementById("list");
        this.els.articleSelect = shadow.getElementById("articleSelect");
        this.els.exportSelected = shadow.getElementById("exportSelected");
        this.els.exportAll = shadow.getElementById("exportAll");
        this.els.clear = shadow.getElementById("clear");
        this.els.hide = shadow.getElementById("hide");
        this.els.miniExpand = shadow.getElementById("miniExpand");
        this.els.replyCaptureToggle = shadow.getElementById("replyCaptureToggle");
        this.els.autoExpandToggle = shadow.getElementById("autoExpandToggle");
        this.els.tip = shadow.getElementById("tip");

        this.loadSettings();

        // 事件绑定
        this.els.exportAll.addEventListener("click", () => this.exportAll());
        this.els.exportSelected.addEventListener("click", () => this.exportSelected());
        this.els.clear.addEventListener("click", () => this.clear());
        this.els.hide.addEventListener("click", () => this.toggle());
        this.els.miniExpand.addEventListener("click", () => this.toggle());
        this.els.replyCaptureToggle.addEventListener("change", () => {
          this.saveSettings();
          if (!this.isReplyCaptureEnabled()) {
            stopExpandLoop();
          } else if (this.isAutoExpandEnabled()) {
            startExpandLoop();
          }
        });
        this.els.autoExpandToggle.addEventListener("change", () => {
          if (this.isAutoExpandEnabled()) {
            this.els.replyCaptureToggle.checked = true;
          }
          this.saveSettings();
          if (this.isAutoExpandEnabled()) {
            startExpandLoop();
          } else {
            stopExpandLoop();
          }
        });
        this.els.articleSelect.addEventListener("change", () => {
          this.selectedTitle = this.els.articleSelect.value;
          this.saveSettings();
          this.previewRenderCount = this.previewPageSize;
          this.scheduleUpdate();
        });
        this.els.list.addEventListener("scroll", () => this.onPreviewScroll());

        document.documentElement.appendChild(host);
        this.update();
        storage.init().finally(() => this.restorePersistedRecords());
      };

      if (document.documentElement) mount();
      else new MutationObserver(() => {
        if (document.documentElement) mount();
      }).observe(document, { childList: true, subtree: true });
    },

    loadSettings() {
      try {
        const savedReplyCapture = localStorage.getItem(STORAGE_KEY_REPLY_CAPTURE);
        const savedAutoExpand = localStorage.getItem(STORAGE_KEY_AUTO_EXPAND);
        const savedSelectedTitle = localStorage.getItem(STORAGE_KEY_SELECTED_TITLE);
        if (savedReplyCapture !== null) {
          this.els.replyCaptureToggle.checked = savedReplyCapture === "true";
        }
        if (savedAutoExpand !== null) {
          this.els.autoExpandToggle.checked = savedAutoExpand === "true";
        }
        if (savedSelectedTitle) {
          this.selectedTitle = savedSelectedTitle;
          this.selectionInitialized = true;
        }
      } catch {}
    },

    saveSettings() {
      try {
        localStorage.setItem(STORAGE_KEY_REPLY_CAPTURE, this.els.replyCaptureToggle.checked);
        localStorage.setItem(STORAGE_KEY_AUTO_EXPAND, this.els.autoExpandToggle.checked);
        localStorage.setItem(STORAGE_KEY_SELECTED_TITLE, this.selectedTitle || "");
      } catch {}
    },

    update() {
      if (!this.shadow) return;

      this.els.count.textContent = `${records.length} 条评论`;

      const titleCounts = collectTitleCounts();
      const titles = [...titleCounts.keys()].sort();
      const select = this.els.articleSelect;
      const currentPageTitle = getCurrentNoteMeta().title;
      const hasCurrentPageTitle = !!currentPageTitle && currentPageTitle !== "未知标题";
      const currentVal = hasCurrentPageTitle && titleCounts.has(currentPageTitle)
        ? currentPageTitle
        : "";

      const prevCurrentCount = currentVal ? (this.prevTitleCounts.get(currentVal) || 0) : 0;
      const currCurrentCount = currentVal ? (titleCounts.get(currentVal) || 0) : 0;
      const currentTitleHasNewCapture = !!currentVal && currCurrentCount > prevCurrentCount;

      const prevSelectedTitle = this.selectedTitle;
      if (!this.selectionInitialized) {
        this.selectedTitle = currentVal;
        this.selectionInitialized = true;
      } else if (currentTitleHasNewCapture && currentVal) {
        this.selectedTitle = currentVal;
      }

      if (this.selectedTitle && !titleCounts.has(this.selectedTitle)) {
        this.selectedTitle = currentVal;
      }

      const selectionChangedByAutoSync = this.selectedTitle !== prevSelectedTitle;
      if (selectionChangedByAutoSync) {
        this.previewRenderCount = this.previewPageSize;
      }

      select.innerHTML = buildSelectOptionsHtml(titles, this.selectedTitle, titleCounts);
      select.value = this.selectedTitle;
      this.prevTitleCounts = titleCounts;
      this.saveSettings();

      this.els.exportAll.disabled = records.length === 0;
      this.els.exportSelected.disabled = !select.value;

      const orderedPreview = this.buildOrderedPreviewRecords(this.selectedTitle);
      const totalPreview = orderedPreview.length;

      if (this.previewRenderCount < this.previewPageSize && totalPreview >= this.previewPageSize) {
        this.previewRenderCount = this.previewPageSize;
      }
      if (this.previewRenderCount > totalPreview) {
        this.previewRenderCount = totalPreview;
      }

      this.renderPreviewList(orderedPreview, !this.selectedTitle);
      this.previewLastTotal = totalPreview;

      if (selectionChangedByAutoSync) {
        this.els.list.scrollTop = 0;
      }
    },

    buildOrderedPreviewRecords(titleFilter = "") {
      const source = titleFilter
        ? records.filter(r => r.title === titleFilter)
        : records;

      const sorted = source.slice().sort((a, b) => {
        if (a.captured_at !== b.captured_at) return (a.captured_at || 0) - (b.captured_at || 0);
        return (a.create_time_raw || 0) - (b.create_time_raw || 0);
      });

      const mains = sorted.filter(r => r.parentCommentId === null);
      const subsByMain = new Map();
      const orphanSubs = [];

      for (const r of sorted) {
        if (r.parentCommentId === null) continue;
        const linkedMainId = r.targetCommentId || r.parentCommentId;
        if (!linkedMainId) {
          orphanSubs.push(r);
          continue;
        }
        const key = `${r.title}__${linkedMainId}`;
        if (!subsByMain.has(key)) subsByMain.set(key, []);
        subsByMain.get(key).push(r);
      }

      const ordered = [];
      for (const main of mains) {
        ordered.push(main);
        const key = `${main.title}__${main.id}`;
        const subs = subsByMain.get(key) || [];
        for (const sub of subs) {
          ordered.push(sub);
        }
      }

      for (const sub of orphanSubs) {
        ordered.push(sub);
      }

      return ordered;
    },

    renderPreviewList(orderedPreview, showTitleTag) {
      const shown = orderedPreview.slice(0, Math.min(this.previewRenderCount, orderedPreview.length));
      this.els.list.innerHTML = shown.map(r => {
        const meta = `${r.create_time || ""}${r.ip_location ? " | " + r.ip_location : ""}`;
        const isSub = r.parentCommentId !== null;
        return `<div class="item ${isSub ? "is-sub" : ""}">
          <div class="meta">
            ${showTitleTag ? `<span class="title-tag">📄 ${escapeHtml(truncateTitle(r.title))}</span>` : ""}
            <span>${escapeHtml(meta)}</span>
            ${isSub ? '<span class="sub-indicator">↳ 追评</span>' : ''}
          </div>
          <div class="content-text">${escapeHtml(r.content)}</div>
        </div>`;
      }).join("");
    },

    onPreviewScroll() {
      const list = this.els.list;
      if (!list) return;
      if (list.scrollTop + list.clientHeight < list.scrollHeight - 8) return;

      const ordered = this.buildOrderedPreviewRecords(this.selectedTitle);
      const total = ordered.length;
      if (this.previewRenderCount >= total) return;

      this.previewRenderCount = Math.min(this.previewRenderCount + this.previewPageSize, total);
      this.renderPreviewList(ordered, !this.selectedTitle);
    },

    async clear() {
      await flushPendingRecords(true).catch(() => {});
      resetRuntimeRecords();
      this.selectedTitle = "";
      this.selectionInitialized = false;
      this.prevTitleCounts = new Map();
      this.previewRenderCount = this.previewPageSize;
      this.previewLastTotal = 0;
      this.setStatus("");
      stopExpandLoop();
      await storage.clearAllComments().catch((e) => {
        console.warn("[xhs] 清理本地备份失败", e);
      });
      this.scheduleUpdate();
    },

    toggle() {
      const panel = this.shadow.querySelector(".panel");
      panel.classList.toggle("is-minimized");
      this.els.hide.textContent = "最小化";
    },

    isReplyCaptureEnabled() {
      return !!this.els.replyCaptureToggle?.checked;
    },

    isAutoExpandEnabled() {
      return !!this.els.autoExpandToggle?.checked;
    },

    // ---------- 导出功能（支持层级序号）----------
    async ensureXlsx() {
      if (typeof XLSX !== "undefined") return;
      if (this.xlsxLoading) throw new Error("XLSX 正在加载，请稍后");
      this.xlsxLoading = true;

      for (const src of XLSX_CDN_LIST) {
        try {
          await new Promise((resolve, reject) => {
            const script = document.createElement("script");
            script.src = src;
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
          });
          this.xlsxLoading = false;
          return;
        } catch {}
      }
      this.xlsxLoading = false;
      throw new Error("所有 CDN 均加载失败");
    },

    // 构建带层级序号的导出数据（按文章分组）
    buildDataForTitle(titleFilter) {
      const filtered = records.filter(r => r.title === titleFilter);
      const mains = filtered.filter(r => r.parentCommentId === null);
      const rows = [
        ["序号", "评论内容", "评论时间", "IP归属地", "", "", ""],
      ];

      for (let i = 0; i < mains.length; i++) {
        const main = mains[i];
        const mainIdx = i + 1;

        rows.push([
          String(mainIdx),
          main.content,
          main.create_time,
          main.ip_location,
          "",
          "",
          "",
        ]);

        const subs = filtered.filter(r => {
          if (r.parentCommentId === null) return false;
          return r.targetCommentId === main.id || r.parentCommentId === main.id;
        });
        subs.sort((a, b) => (a.create_time_raw || 0) - (b.create_time_raw || 0));

        for (let j = 0; j < subs.length; j++) {
          const sub = subs[j];
          rows.push([
            `  ${mainIdx}_${j+1}`,
            sub.content,
            sub.create_time,
            sub.ip_location,
            "",
            "",
            "",
          ]);
        }
      }

      // 添加仪表盘数据（在右侧，与评论表格间隔一列）
      // 仪表盘从第1行开始，标题一列，内容一列
      if (rows.length > 0) {
        // 添加仪表盘数据行
        const dashboardData = [
          ["标题", titleFilter || ""],
          ["正文", filtered[0]?.note_desc || ""],
          ["主图链接", filtered[0]?.cover_image || ""],
        ];
        
        // 确保有足够的行来放置仪表盘数据
        for (let i = 0; i < dashboardData.length; i++) {
          const rowIndex = i + 1;
          if (rowIndex < rows.length) {
            // 如果行已存在，在该行的第6列（索引5）和第7列（索引6）添加仪表盘数据
            rows[rowIndex][5] = dashboardData[i][0];
            rows[rowIndex][6] = dashboardData[i][1];
          } else {
            // 如果行不存在，创建新行
            rows.push(["", "", "", "", "", dashboardData[i][0], dashboardData[i][1]]);
          }
        }
      }

      return rows;
    },

    async exportAll() {
      if (records.length === 0) return alert("暂无数据");
      try {
        await flushPendingRecords(true);
        await this.ensureXlsx();
        const wb = XLSX.utils.book_new();
        const titles = [...new Set(records.map(r => r.title))].sort();
        const usedSheetNames = new Set();

        for (const title of titles) {
          const data = this.buildDataForTitle(title);
          if (data.length === 0) continue;
          const baseName = safeSheetName(title, "评论");
          let sheetName = baseName;
          let idx = 1;
          while (usedSheetNames.has(sheetName)) {
            const suffix = `_${idx++}`;
            sheetName = `${baseName.slice(0, 31 - suffix.length)}${suffix}`;
          }
          usedSheetNames.add(sheetName);

          const ws = XLSX.utils.aoa_to_sheet(data);
          ws["!cols"] = [
            { wch: 12 },
            { wch: 80 },
            { wch: 24 },
            { wch: 18 },
            { wch: 2 },
            { wch: 12 },
            { wch: 40 },
          ];
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        }

        const filename = `【${formatDateForFilename(Date.now())}】小红书评论合并导出.xlsx`;
        this.download(wb, filename);
      } catch (e) {
        alert("导出失败: " + e.message);
      }
    },

    async exportSelected() {
      const title = this.els.articleSelect.value;
      if (!title) return alert("请选择文章");
      const filtered = records.filter(r => r.title === title);
      if (filtered.length === 0) return alert("该文章无评论");

      try {
        await flushPendingRecords(true);
        await this.ensureXlsx();
        const data = this.buildDataForTitle(title);
        const firstTime = filtered[0].create_time_raw;
        const dateStr = firstTime ? formatDateForFilename(firstTime) : formatDateForFilename(Date.now());
        const safeTitle = title.replace(/[\\/:*?"<>|]/g, '_');
        const filename = `【${dateStr}】${safeTitle}.xlsx`;

        const wb = this.buildWorkbook(data, safeSheetName(title, "评论"));
        this.download(wb, filename);
      } catch (e) {
        alert("导出失败: " + e.message);
      }
    },

    buildWorkbook(data, sheetName) {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(data);
      ws["!cols"] = [
        { wch: 12 },
        { wch: 80 },
        { wch: 24 },
        { wch: 18 },
        { wch: 2 },
        { wch: 12 },
        { wch: 40 },
      ];
      XLSX.utils.book_append_sheet(wb, ws, safeSheetName(sheetName, "评论"));
      return wb;
    },

    download(wb, filename) {
      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
    }
  };

  // 启动 UI
  ui.init();

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") {
      flushPendingRecords(true).catch(() => {});
    }
  });

  window.addEventListener("pagehide", () => {
    flushPendingRecords(true).catch(() => {});
  });

  // ---------- 拦截响应 ----------
  const handledResponses = new WeakSet();
  const origJson = Response.prototype.json;
  const origText = Response.prototype.text;

  Response.prototype.json = async function () {
    const url = this.url || "";
    const data = await origJson.apply(this, arguments);
    if (shouldHandleApiUrl(url) && !handledResponses.has(this)) {
      handledResponses.add(this);
      try {
        handleResultJson(url, data);
      } catch (e) {
        console.warn("[xhs] JSON处理错误:", e);
      }
    }
    return data;
  };

  Response.prototype.text = async function () {
    const url = this.url || "";
    const txt = await origText.apply(this, arguments);
    if (shouldHandleApiUrl(url) && !handledResponses.has(this)) {
      handledResponses.add(this);
      try {
        const data = JSON.parse(txt);
        handleResultJson(url, data);
      } catch (e) {
        // 非JSON，忽略
      }
    }
    return txt;
  };

  // fetch 简单包装
  const hookFetch = (fetchFn) => function (...args) {
    const url = normalizeUrl(args[0]);
    return fetchFn.apply(this, args).then(res => {
      if (shouldHandleApiUrl(url)) {
      }
      return res;
    });
  };

  let _fetch = window.fetch;
  Object.defineProperty(window, "fetch", {
    configurable: true,
    get() { return _fetch; },
    set(v) { _fetch = hookFetch(v); }
  });
  window.fetch = _fetch;

  // XHR 拦截
  const origOpen = XMLHttpRequest.prototype.open;
  const origSend = XMLHttpRequest.prototype.send;

  XMLHttpRequest.prototype.open = function (method, url) {
    this.__spyUrl = normalizeUrl(url);
    return origOpen.apply(this, arguments);
  };

  XMLHttpRequest.prototype.send = function (body) {
    if (shouldHandleApiUrl(this.__spyUrl || "")) {
      this.addEventListener("load", () => {
        try {
          let payload = (this.responseType && this.responseType !== "text" && this.responseType !== "")
            ? this.response
            : this.responseText;
          if (typeof payload === "string") {
            try { payload = JSON.parse(payload); } catch { payload = null; }
          }
          if (payload) handleResultJson(this.__spyUrl, payload);
        } catch (e) {
          console.warn("[xhs] XHR处理错误:", e);
        }
      });
    }
    return origSend.apply(this, arguments);
  };

  // 调试接口
  window.__xhsCommentSpy = {
    getRecords: () => records.slice(),
    clear: () => ui.clear(),
    exportAll: () => ui.exportAll(),
    flush: () => flushPendingRecords(true),
    exportSelected: (title) => {
      if (title) {
        ui.els.articleSelect.value = title;
        ui.exportSelected();
      }
    }
  };
})();
