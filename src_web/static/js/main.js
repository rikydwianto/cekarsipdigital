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


$(document).ready(function() {
  $('.btn-submit').hide();
  $('.select2').select2();
  $("#nomor_center").empty();
  $("#nomor_center").append('<option value="">Memuat...</option>');
  function loadNoCenter(){
    $.ajax({
    url: '/api/data_center',
    method: 'GET',
    dataType: 'json',
    success: function(respon) {
      data = respon.data;
      $("#nomor_center").empty();
      $("#nomor_center").append('<option value="">Pilih Nomor</option>');
      data.forEach(function(item) {
        $("#nomor_center").append('<option value="' + item + '">' + item + '</option>');
      });
    },
    error: function() {
      $("#nomor_center").empty();
      $("#nomor_center").append('<option value="">Gagal memuat data</option>');
    }
  });
  }
  loadNoCenter();
  $("#nomor_center").on('click', function() {
    loadNoCenter();
  });
  $("#nomor_center").on('change', function() {
    var selectedValue = $(this).val();
    console.log("Nomor Center dipilih: " + selectedValue);
  });
  function getAnggotaByCenter(nomor_center){
     $.ajax({
      url: '/api/data_center/' + nomor_center,
      method: 'GET',
      dataType: 'json',
      success: function(respon) {
        data = respon.data;
        id_nama.empty();
        id_nama.append('<option value="">Pilih Anggota</option>');
        data.forEach(function(item) {
          id_nama.append('<option value="' + item.ID_NAMA_ANGGOTA + '">' + item.NOMOR_CENTER  + ' - ' + item.ID_NAMA_ANGGOTA + '</option>');
        });
      },
      error: function() {
        id_nama.empty();
        id_nama.append('<option value="">Gagal memuat data anggota</option>');
      }
    });
  }
  var id_nama = $("#id_anggota");
  id_nama.empty();
  id_nama.append('<option value="">Pilih Center dulu...</option>');
  $("#nomor_center").on('change', function() {
    var selectedNoCenter = $(this).val();
    id_nama.empty();
    id_nama.append('<option value="">Memuat anggota center ' + selectedNoCenter + ' ...</option>');
    getAnggotaByCenter(selectedNoCenter);
  });
  $("#id_anggota").on('change', function() {
    var selectedValue = $(this).val();
    $.ajax({
      url: '/api/anggota/' + selectedValue,
      method: 'GET',
      success: function(respon) {
        // Lakukan sesuatu dengan data anggota yang diterima
        data = respon.data[0];
        $("#folder").val(data.PATH);
        $('.btn-submit').show();
      },
      error: function() {
        console.error("Gagal memuat data anggota");
      }
    });
  });
  $(".btn-submit").on('click', function(e) {
  e.preventDefault(); // cegah reload

  var form = $("#form-anggota-keluar")[0];
  var formData = new FormData(form); // ambil semua input termasuk file

  var tgl_keluar = $("#tanggal_keluar").val();
  if (!tgl_keluar) {
    Swal.fire({
      title: 'Peringatan',
      text: "Tanggal keluar harus diisi.",
      icon: 'warning',
      confirmButtonText: 'OK'
    });
    return;
  }

  Swal.fire({
    title: 'Konfirmasi',
    text: "Apakah Anda yakin ingin menyimpan data ini?",
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#3085d6',
    cancelButtonColor: '#d33',
    confirmButtonText: 'Ya, simpan',
    cancelButtonText: 'Batal',
  }).then((result) => {
    if (result.isConfirmed) {
      $.ajax({
        url: '/api/anggota_keluar', // ganti dengan endpoint backend kamu
        type: 'POST',
        data: formData,
        processData: false, // penting agar FormData tidak diubah ke query string
        contentType: false, // biarkan browser set otomatis
        success: function(response) {
          Swal.fire({
            title: 'Berhasil!',
            text: 'Data berhasil disimpan.',
            icon: 'success',
            confirmButtonText: 'OK'
          });
          console.log("Response:", response);
        },
        error: function(xhr, status, error) {
          Swal.fire({
            title: 'Gagal!',
            text: 'Terjadi kesalahan saat menyimpan data.',
            icon: 'error',
            confirmButtonText: 'OK'
          });
          console.error("Error:", error);
        }
      });
    }
  });
});

});