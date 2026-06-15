/**
 * 今週の良かったこと3つだけ日記 - script.js
 * =============================================
 * GASのWebアプリURLをここに設定してください。
 * デプロイ後に発行されるURLを下記の GAS_URL に貼り付けます。
 */
const GAS_URL = "https://script.google.com/macros/s/AKfycbwvYfwcOJLXMHeiI803_rWUsyVDgvJJWJViNn0_t7bEBxjZ-vE_YNz3UXlt_t8pa52H/exec";
// ↑ここを自分のGAS WebアプリURLに書き換えてください

// =============================================
// 状態管理
// =============================================
let pendingDeleteId = null;
let isSubmitting = false;

// =============================================
// DOM参照
// =============================================
const form = document.getElementById("diary-form");
const editIdInput = document.getElementById("edit-id");
const good1 = document.getElementById("good1");
const good2 = document.getElementById("good2");
const good3 = document.getElementById("good3");
const submitBtn = document.getElementById("submit-btn");
const submitLabel = document.getElementById("submit-label");
const cancelBtn = document.getElementById("cancel-btn");
const formTitle = document.getElementById("form-title");
const formSubtitle = document.getElementById("form-subtitle");
const formMessage = document.getElementById("form-message");
const inputPanel = document.getElementById("input-section");

const loadingState = document.getElementById("loading-state");
const emptyState = document.getElementById("empty-state");
const errorState = document.getElementById("error-state");
const errorMessage = document.getElementById("error-message");
const entriesList = document.getElementById("entries-list");

const deleteModal = document.getElementById("delete-modal");
const modalCancel = document.getElementById("modal-cancel");
const modalConfirm = document.getElementById("modal-confirm");

// =============================================
// 初期化
// =============================================
document.addEventListener("DOMContentLoaded", () => {
  loadEntries();
  setupCharCounters();
  setupModal();
  setupForm();
  setupCancelBtn();
});

// =============================================
// 文字数カウンター
// =============================================
function setupCharCounters() {
  [[good1, "count1"], [good2, "count2"], [good3, "count3"]].forEach(([ta, countId]) => {
    const counter = document.getElementById(countId);
    ta.addEventListener("input", () => {
      const len = ta.value.length;
      counter.textContent = `${len} / 200`;
      counter.classList.toggle("warn", len > 160);
      if (len > 200) ta.value = ta.value.slice(0, 200);
    });
  });
}

// =============================================
// フォーム送信
// =============================================
function setupForm() {
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (isSubmitting) return;

    const g1 = good1.value.trim();
    const g2 = good2.value.trim();
    const g3 = good3.value.trim();

    if (!g1 || !g2 || !g3) {
      showMessage("3つすべて入力してください。", "error");
      return;
    }

    const id = editIdInput.value;
    const action = id ? "update" : "create";

    setSubmitting(true);

    try {
      const params = new URLSearchParams({ action, good1: g1, good2: g2, good3: g3 });
      if (id) params.append("id", id);

      const res = await fetch(GAS_URL, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params.toString(),
      });

      const data = await res.json();

      if (data.status === "ok") {
        showMessage(action === "create" ? "保存しました ✦" : "更新しました ✦", "success");
        resetForm();
        await loadEntries();
      } else {
        showMessage(`エラー: ${data.message || "不明なエラー"}`, "error");
      }
    } catch (err) {
      console.error(err);
      showMessage("通信エラーが発生しました。GASのURLを確認してください。", "error");
    } finally {
      setSubmitting(false);
    }
  });
}

function setSubmitting(state) {
  isSubmitting = state;
  submitBtn.disabled = state;
  submitLabel.textContent = state
    ? (editIdInput.value ? "更新中..." : "保存中...")
    : (editIdInput.value ? "更新する" : "保存する");
}

// =============================================
// フォームリセット
// =============================================
function resetForm() {
  form.reset();
  editIdInput.value = "";
  formTitle.textContent = "今週の良かったことを書く";
  formSubtitle.textContent = "3つだけ、ゆっくり思い出してみてください。";
  submitLabel.textContent = "保存する";
  cancelBtn.style.display = "none";
  inputPanel.classList.remove("editing");
  ["count1", "count2", "count3"].forEach(id => {
    document.getElementById(id).textContent = "0 / 200";
  });
}

// =============================================
// キャンセルボタン
// =============================================
function setupCancelBtn() {
  cancelBtn.addEventListener("click", () => {
    resetForm();
    hideMessage();
  });
}

// =============================================
// 編集モードへ切り替え
// =============================================
function startEditing(entry) {
  editIdInput.value = entry.id;
  good1.value = entry.good1;
  good2.value = entry.good2;
  good3.value = entry.good3;

  // 文字数カウンターを更新
  ["good1", "good2", "good3"].forEach((name, i) => {
    const ta = document.getElementById(name);
    const counter = document.getElementById(`count${i + 1}`);
    counter.textContent = `${ta.value.length} / 200`;
  });

  formTitle.textContent = "記録を編集する";
  formSubtitle.textContent = "内容を変更して「更新する」を押してください。";
  submitLabel.textContent = "更新する";
  cancelBtn.style.display = "inline-flex";
  inputPanel.classList.add("editing");
  hideMessage();

  // スマホではフォームまでスクロール
  inputPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  setTimeout(() => good1.focus(), 400);
}

// =============================================
// データ取得
// =============================================
async function loadEntries() {
  showState("loading");

  try {
    const res = await fetch(`${GAS_URL}?action=read`, { cache: "no-cache" });
    const data = await res.json();

    if (data.status === "ok") {
      renderEntries(data.entries);
    } else {
      showState("error");
      errorMessage.textContent = data.message || "データ取得に失敗しました。";
    }
  } catch (err) {
    console.error(err);
    showState("error");
    errorMessage.textContent = "GASのURLが正しく設定されていないか、デプロイが完了していない可能性があります。";
  }
}

// =============================================
// エントリーレンダリング
// =============================================
function renderEntries(entries) {
  if (!entries || entries.length === 0) {
    showState("empty");
    return;
  }

  showState("list");
  entriesList.innerHTML = "";

  // 新しい順に並べる
  const sorted = [...entries].sort((a, b) => b.id - a.id);

  sorted.forEach((entry, index) => {
    const card = createEntryCard(entry, index);
    entriesList.appendChild(card);
  });
}

function createEntryCard(entry, index) {
  const card = document.createElement("article");
  card.className = "entry-card";
  card.dataset.id = entry.id;
  card.style.animationDelay = `${index * 0.06}s`;

  const date = formatDate(entry.id);

  card.innerHTML = `
    <div class="entry-date">${date}</div>
    <ul class="entry-goods">
      <li class="entry-good">
        <span class="entry-good-num">1</span>
        <span>${escapeHtml(entry.good1)}</span>
      </li>
      <li class="entry-good">
        <span class="entry-good-num">2</span>
        <span>${escapeHtml(entry.good2)}</span>
      </li>
      <li class="entry-good">
        <span class="entry-good-num">3</span>
        <span>${escapeHtml(entry.good3)}</span>
      </li>
    </ul>
    <div class="entry-actions">
      <button class="btn btn-ghost btn-sm edit-btn" aria-label="この記録を編集する">
        ✏️ 編集
      </button>
      <button class="btn btn-ghost btn-sm delete-btn" style="color: var(--danger);" aria-label="この記録を削除する">
        🗑️ 削除
      </button>
    </div>
  `;

  card.querySelector(".edit-btn").addEventListener("click", () => startEditing(entry));
  card.querySelector(".delete-btn").addEventListener("click", () => openDeleteModal(entry.id));

  return card;
}

// =============================================
// 削除モーダル
// =============================================
function setupModal() {
  modalCancel.addEventListener("click", closeDeleteModal);
  modalConfirm.addEventListener("click", confirmDelete);
  deleteModal.addEventListener("click", (e) => {
    if (e.target === deleteModal) closeDeleteModal();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && deleteModal.classList.contains("open")) closeDeleteModal();
  });
}

function openDeleteModal(id) {
  pendingDeleteId = id;
  deleteModal.classList.add("open");
  modalConfirm.focus();
}

function closeDeleteModal() {
  deleteModal.classList.remove("open");
  pendingDeleteId = null;
}

async function confirmDelete() {
  if (!pendingDeleteId) return;
  const id = pendingDeleteId;
  closeDeleteModal();

  try {
    const params = new URLSearchParams({ action: "delete", id });
    const res = await fetch(GAS_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    });
    const data = await res.json();

    if (data.status === "ok") {
      // 楽観的UI更新
      const card = document.querySelector(`.entry-card[data-id="${id}"]`);
      if (card) {
        card.style.transition = "opacity 0.3s ease, transform 0.3s ease";
        card.style.opacity = "0";
        card.style.transform = "scale(0.96)";
        setTimeout(async () => {
          await loadEntries();
        }, 300);
      } else {
        await loadEntries();
      }
      // 編集中の投稿を削除した場合はフォームをリセット
      if (editIdInput.value === String(id)) resetForm();
    } else {
      alert(`削除できませんでした: ${data.message || "不明なエラー"}`);
    }
  } catch (err) {
    console.error(err);
    alert("通信エラーが発生しました。");
  }
}

// =============================================
// ユーティリティ
// =============================================
function showState(state) {
  loadingState.style.display = state === "loading" ? "block" : "none";
  emptyState.style.display = state === "empty" ? "block" : "none";
  errorState.style.display = state === "error" ? "block" : "none";
  entriesList.style.display = state === "list" ? "flex" : "none";
}

function showMessage(text, type) {
  formMessage.textContent = text;
  formMessage.className = `form-message show ${type}`;
}

function hideMessage() {
  formMessage.className = "form-message";
}

function formatDate(id) {
  const ts = Number(id);
  if (isNaN(ts)) return id;
  const d = new Date(ts);
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const h = String(d.getHours()).padStart(2, "0");
  const min = String(d.getMinutes()).padStart(2, "0");
  const week = ["日", "月", "火", "水", "木", "金", "土"][d.getDay()];
  return `${y}年${m}月${day}日（${week}） ${h}:${min}`;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
