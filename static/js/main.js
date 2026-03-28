// 倒计时函数
function updateCountdown() {
    const targetDate = new Date(2026, 3, 3, 0, 0, 0);
    const now = new Date();
    const diff = targetDate - now;
    
    if (diff <= 0) {
        document.getElementById('days').textContent = '00';
        document.getElementById('hours').textContent = '00';
        document.getElementById('minutes').textContent = '00';
        document.getElementById('seconds').textContent = '00';
        return;
    }
    
    const days = Math.floor(diff / (1000 * 60 * 60 * 24));
    const hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
    const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
    const seconds = Math.floor((diff % (1000 * 60)) / 1000);
    
    document.getElementById('days').textContent = String(days).padStart(2, '0');
    document.getElementById('hours').textContent = String(hours).padStart(2, '0');
    document.getElementById('minutes').textContent = String(minutes).padStart(2, '0');
    document.getElementById('seconds').textContent = String(seconds).padStart(2, '0');
}

// 显示消息
function showMessage(element, msg, type) {
    element.textContent = msg;
    element.className = `message ${type}`;
    element.style.display = 'block';
    setTimeout(() => {
        element.style.display = 'none';
    }, 3000);
}

// 显示弹窗
function showModal(message, callback) {
    const modal = document.createElement('div');
    modal.className = 'modal';
    modal.innerHTML = `
        <div class="modal-content">
            <p>${message}</p>
        </div>
    `;
    document.body.appendChild(modal);
    setTimeout(() => {
        modal.remove();
        if (callback) callback();
    }, 3000);
}