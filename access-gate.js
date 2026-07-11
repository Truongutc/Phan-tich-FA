/**
 * access-gate.js — Khoá truy cập file PDF/Excel gốc trên Google Drive bằng mật khẩu, dùng chung cho
 * mọi trang (index.html + các dashboard ngành). Web vẫn XEM ĐƯỢC báo cáo tự do — chỉ chặn khi bấm
 * tải/mở file gốc trên Drive (nút "Báo Cáo PDF"/"Excel Model" ở trang chi tiết cổ phiếu, và nút
 * "Google Drive" ở trang chủ).
 *
 * Đổi mật khẩu qua GitHub Action "Cập nhật mật khẩu truy cập Drive" (.github/workflows/
 * update_drive_password.yml) — KHÔNG sửa tay dòng dưới, Action tìm đúng dòng có nhãn SYS_PWD_DRIVE.
 *
 * Đây là khoá "mềm" (chặn ở phía trình duyệt) để hạn chế người xem thường bấm nhầm/tải tràn lan,
 * KHÔNG phải bảo mật thật sự — link Drive gốc vẫn xem được nếu ai đó cố tình lấy trực tiếp.
 */
'use strict';

const DRIVE_ACCESS_PASSWORD = '123123'; // SYS_PWD_DRIVE

const _DRIVE_GATE_SESSION_KEY = 'aic_drive_unlocked';

function _driveGateUnlocked() {
    return sessionStorage.getItem(_DRIVE_GATE_SESSION_KEY) === '1';
}

function _driveGateEnsureModal() {
    let modal = document.getElementById('drive-gate-modal');
    if (modal) return modal;

    modal = document.createElement('div');
    modal.id = 'drive-gate-modal';
    modal.innerHTML = `
        <div class="drive-gate-backdrop">
          <div class="drive-gate-box">
            <h3>🔒 Nhập mật khẩu để truy cập file</h3>
            <p>File PDF/Excel gốc trên Google Drive được giới hạn truy cập.</p>
            <input type="password" id="drive-gate-input" placeholder="Mật khẩu" autocomplete="off" />
            <div class="drive-gate-err" id="drive-gate-err">Mật khẩu không đúng!</div>
            <div class="drive-gate-actions">
              <button type="button" id="drive-gate-cancel">Huỷ</button>
              <button type="button" id="drive-gate-submit">Xác nhận</button>
            </div>
          </div>
        </div>`;
    document.body.appendChild(modal);

    const style = document.createElement('style');
    style.textContent = `
        #drive-gate-modal { display: none; }
        #drive-gate-modal .drive-gate-backdrop {
            position: fixed; inset: 0; background: rgba(0,0,0,.65);
            display: flex; align-items: center; justify-content: center; z-index: 9999;
        }
        #drive-gate-modal .drive-gate-box {
            background: #0f172a; border: 1px solid #334155; border-radius: 12px;
            padding: 24px; width: 320px; max-width: 90vw; color: #e2e8f0;
            font-family: 'Inter', sans-serif; box-shadow: 0 20px 50px rgba(0,0,0,.5);
        }
        #drive-gate-modal h3 { margin: 0 0 8px; font-size: 16px; }
        #drive-gate-modal p { margin: 0 0 14px; font-size: 13px; color: #94a3b8; }
        #drive-gate-modal input {
            width: 100%; box-sizing: border-box; padding: 9px 10px; border-radius: 6px;
            border: 1px solid #334155; background: #1e293b; color: #fff; margin-bottom: 8px;
            font-size: 14px;
        }
        #drive-gate-modal input:focus { outline: none; border-color: #2563eb; }
        #drive-gate-modal .drive-gate-err {
            display: none; color: #f87171; font-size: 12px; margin-bottom: 8px;
        }
        #drive-gate-modal .drive-gate-actions { display: flex; justify-content: flex-end; gap: 8px; }
        #drive-gate-modal button {
            padding: 7px 16px; border-radius: 6px; border: none; cursor: pointer; font-size: 13px;
            font-weight: 600;
        }
        #drive-gate-modal #drive-gate-cancel { background: #334155; color: #e2e8f0; }
        #drive-gate-modal #drive-gate-cancel:hover { background: #475569; }
        #drive-gate-modal #drive-gate-submit { background: #2563eb; color: #fff; }
        #drive-gate-modal #drive-gate-submit:hover { background: #1d4ed8; }
    `;
    document.head.appendChild(style);
    return modal;
}

function _driveGateShow(onSuccess) {
    const modal = _driveGateEnsureModal();
    modal.style.display = 'block';
    const input = document.getElementById('drive-gate-input');
    const err = document.getElementById('drive-gate-err');
    const submitBtn = document.getElementById('drive-gate-submit');
    const cancelBtn = document.getElementById('drive-gate-cancel');
    input.value = '';
    err.style.display = 'none';
    setTimeout(() => input.focus(), 0);

    function cleanup() {
        modal.style.display = 'none';
        submitBtn.onclick = null;
        cancelBtn.onclick = null;
        input.onkeydown = null;
    }
    function trySubmit() {
        if (input.value === DRIVE_ACCESS_PASSWORD) {
            sessionStorage.setItem(_DRIVE_GATE_SESSION_KEY, '1');
            cleanup();
            onSuccess();
        } else {
            err.style.display = 'block';
            input.value = '';
            input.focus();
        }
    }
    submitBtn.onclick = trySubmit;
    cancelBtn.onclick = cleanup;
    input.onkeydown = (e) => { if (e.key === 'Enter') trySubmit(); if (e.key === 'Escape') cleanup(); };
}

// Chặn click ở PHA CAPTURE trên bất kỳ link nào đánh dấu cần khoá — bắt được cả link mà app_*.js
// gắn href/hiện ra SAU khi trang đã tải xong (data/<TICKER>.json fetch xong mới set href), vì đây là
// 1 listener gắn 1 lần trên document, không cần biết phần tử tồn tại từ trước.
document.addEventListener('click', function (e) {
    const el = e.target.closest('#download-pdf, #download-excel, [data-drive-gate]');
    if (!el || !el.href) return;
    if (_driveGateUnlocked()) return; // đã mở khoá trong phiên trình duyệt này — cho qua bình thường
    e.preventDefault();
    e.stopPropagation();
    const targetUrl = el.href;
    const openInNewTab = el.target === '_blank';
    _driveGateShow(() => {
        if (openInNewTab) window.open(targetUrl, '_blank');
        else window.location.href = targetUrl;
    });
}, true);
