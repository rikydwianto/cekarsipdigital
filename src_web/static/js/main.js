/* Arsip Owncloud Web Server - Main JavaScript */

// Update waktu real-time
function updateTime() {
  const timeElement = document.getElementById("server-time");
  const footerTimeElement = document.getElementById("footer-time");

  if (timeElement || footerTimeElement) {
    const now = new Date();
    const options = {
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    };
    const timeString = now.toLocaleDateString("id-ID", options);

    if (timeElement) {
      timeElement.textContent = timeString;
    }
    if (footerTimeElement) {
      footerTimeElement.textContent = timeString;
    }
  }
}

// Copy server URL to clipboard
function copyServerUrl() {
  // Get URL from the alert div
  const urlElement = document.querySelector(".alert-info .font-monospace a");
  if (urlElement) {
    const url = urlElement.href;

    navigator.clipboard
      .writeText(url)
      .then(() => {
        // Show success feedback
        const button = event.target.closest("button");
        const originalHTML = button.innerHTML;
        button.innerHTML = '<i class="bi bi-check"></i> Copied!';
        button.classList.remove("btn-outline-primary");
        button.classList.add("btn-success");

        setTimeout(() => {
          button.innerHTML = originalHTML;
          button.classList.remove("btn-success");
          button.classList.add("btn-outline-primary");
        }, 2000);
      })
      .catch((err) => {
        alert("Gagal copy URL: " + err);
      });
  }
}

// Refresh QR Code
function refreshQRCode() {
  const qrImage = document.querySelector("#qr-container img");
  if (qrImage) {
    const button = event.target.closest("button");
    const originalHTML = button.innerHTML;

    // Show loading
    button.innerHTML = '<i class="bi bi-arrow-clockwise"></i> Loading...';
    button.disabled = true;

    // Reload QR code with cache buster
    const currentSrc = qrImage.src.split("?")[0];
    qrImage.src = currentSrc + "?t=" + new Date().getTime();

    // Reset button after 1 second
    setTimeout(() => {
      button.innerHTML = originalHTML;
      button.disabled = false;
    }, 1000);
  }
}

// Update setiap detik
setInterval(updateTime, 1000);

// Initial update
document.addEventListener("DOMContentLoaded", function () {
  updateTime();

  // Add fade-in animation to cards
  const cards = document.querySelectorAll(".card");
  cards.forEach((card, index) => {
    card.style.animation = `fadeIn 0.5s ease-in ${index * 0.1}s`;
    card.style.animationFillMode = "both";
  });

  // Add click to copy functionality for code blocks
  const codeBlocks = document.querySelectorAll("pre code");
  codeBlocks.forEach((block) => {
    block.style.cursor = "pointer";
    block.title = "Click to copy";

    block.addEventListener("click", function () {
      const text = this.textContent;
      navigator.clipboard.writeText(text).then(() => {
        // Show temporary tooltip
        const originalTitle = this.title;
        this.title = "Copied!";
        setTimeout(() => {
          this.title = originalTitle;
        }, 2000);
      });
    });
  });

  console.log("🌐 Arsip Owncloud Web Server - Ready!");
  console.log("📚 Bootstrap 5 + Custom Template System");
  console.log("🚀 API Routes: /api/hello, /api/status, /api/server-info");
});
