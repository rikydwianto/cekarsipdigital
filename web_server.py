"""
Web Server Module untuk Arsip Owncloud
Menyediakan HTTP server sederhana untuk akses file arsip melalui browser
"""

import os
import socket
import threading
import time
from http.server import HTTPServer, BaseHTTPRequestHandler
import json
from datetime import datetime
import platform
import io
import gzip
import hashlib
from functools import lru_cache

# QR Code library
import qrcode

# Pandas for Excel reading
import pandas as pd

# Import template loader
from src_web.template_loader import TemplateLoader, get_default_context, create_api_response, json_response
from datetime import datetime

def get_tahun(tanggal_str: str) -> str:
    """
    Mengambil tahun dari tanggal (format YYYY-MM-DD)
    Contoh: '2025-10-19' → '2025'
    """
    try:
        tanggal = datetime.strptime(tanggal_str, "%Y-%m-%d")
        return str(tanggal.year)
    except Exception as e:
        raise ValueError(f"Format tanggal tidak valid: {tanggal_str} ({e})")


def get_bulan(tanggal_str: str) -> str:
    """
    Mengambil bulan dari tanggal dalam format '01.JANUARI'
    Contoh: '2025-10-19' → '10.OKTOBER'
    """
    try:
        tanggal = datetime.strptime(tanggal_str, "%Y-%m-%d")
        bulan_index = tanggal.month

        bulan_mapping = {
            1: "JANUARI",
            2: "FEBRUARI",
            3: "MARET",
            4: "APRIL",
            5: "MEI",
            6: "JUNI",
            7: "JULI",
            8: "AGUSTUS",
            9: "SEPTEMBER",
            10: "OKTOBER",
            11: "NOVEMBER",
            12: "DESEMBER",
        }

        bulan_nama = bulan_mapping.get(bulan_index, "UNKNOWN")
        return f"{bulan_index:02d}.{bulan_nama}"

    except Exception as e:
        raise ValueError(f"Format tanggal tidak valid: {tanggal_str} ({e})")


class DataCache:
    """Simple cache manager for database operations"""
    
    def __init__(self, max_size=100, ttl=300):  # 5 minutes TTL
        self.cache = {}
        self.timestamps = {}
        self.max_size = max_size
        self.ttl = ttl
    
    def get(self, key):
        """Get cached data if still valid"""
        if key in self.cache:
            if time.time() - self.timestamps[key] < self.ttl:
                return self.cache[key]
            else:
                # Expired, remove from cache
                del self.cache[key]
                del self.timestamps[key]
        return None
    
    def set(self, key, value):
        """Set cache with automatic cleanup"""
        # Clean old entries if cache is full
        if len(self.cache) >= self.max_size:
            oldest = min(self.timestamps, key=self.timestamps.get)
            del self.cache[oldest]
            del self.timestamps[oldest]
        
        self.cache[key] = value
        self.timestamps[key] = time.time()
    
    def invalidate(self, pattern=None):
        """Invalidate cache entries"""
        if pattern is None:
            self.cache.clear()
            self.timestamps.clear()
        else:
            keys_to_remove = [k for k in self.cache.keys() if pattern in k]
            for key in keys_to_remove:
                del self.cache[key]
                del self.timestamps[key]

# Global cache instance
data_cache = DataCache()

class HelloWorldHandler(BaseHTTPRequestHandler):
    """Custom HTTP Request Handler dengan template system dan API support"""
    
    def __init__(self, *args, **kwargs):
        self.template_loader = TemplateLoader()
        super().__init__(*args, **kwargs)
    
    def log_message(self, format, *args):
        """Override log_message untuk suppress console output"""
        pass
    
    def do_OPTIONS(self):
        """Handle OPTIONS request for CORS preflight"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Accept, Accept-Encoding')
        self.send_header('Access-Control-Max-Age', '86400')  # 24 hours
        self.end_headers()
    
    def do_POST(self):
        """Handle POST request"""
        try:
            # Handle API POST routes
            if self.path.startswith('/api/'):
                self.serve_api_post()
                return
            
            # 404 Not Found untuk non-API POST
            self.send_error(404, "POST endpoint not found")
            
        except Exception as e:
            response = create_api_response(
                success=False,
                message='Internal server error',
                error=str(e)
            )
            self.send_json_response(response, status_code=500)
    
    def do_GET(self):
        """Handle GET request"""
        try:
            # Handle static files
            if self.path.startswith('/static/'):
                self.serve_static_file()
                return
            
            # Handle QR Code generation
            if self.path == '/qrcode':
                self.serve_qrcode()
                return
            
            # Handle API routes
            if self.path.startswith('/api/'):
                self.serve_api()
                return
            
            # Handle home page
            if self.path == '/' or self.path == '/index.html':
                self.serve_home_page()
                return
            
            # Handle about page
            if self.path == '/about':
                self.serve_about_page()
                return
            
            if self.path == '/form_anggota_keluar':
                self.formpageAnggotaKeluar()
                return
            
            # Handle API docs page
            if self.path == '/api/docs':
                self.serve_api_docs_page()
                return
            
            # 404 Not Found
            self.send_error(404, "Page not found")
            
        except Exception as e:
            self.serve_error_page(str(e))

    
    def serve_static_file(self):
        """Serve static files (CSS, JS, images)"""
        try:
            # Remove /static/ prefix
            file_path = self.path[8:]  # len('/static/') = 8
            
            content, content_type = self.template_loader.get_static_file(file_path)
            
            if content is None:
                self.send_error(404, "Static file not found")
                return
            
            self.send_response(200)
            self.send_header('Content-type', content_type)
            
            if isinstance(content, str):
                content_bytes = content.encode('utf-8')
            else:
                content_bytes = content
            
            self.send_header('Content-Length', str(len(content_bytes)))
            self.end_headers()
            self.wfile.write(content_bytes)
            
        except Exception as e:
            self.send_error(500, f"Error serving static file: {str(e)}")
    
    def serve_home_page(self):
        """Serve halaman utama"""
        try:
            # Get server info
            server_info = self.server.server_info if hasattr(self.server, 'server_info') else {}
            server_info['page_title'] = 'Home'
            
            # Get context
            context = get_default_context(server_info)
            
            # Render template dengan partial
            html_content = self.template_loader.render('index.html', context, active_page='home')
            
            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(html_content.encode('utf-8'))))
            self.end_headers()
            self.wfile.write(html_content.encode('utf-8'))
            
        except Exception as e:
            self.serve_error_page(str(e))
    
    def serve_about_page(self):
        """Serve halaman about"""
        try:
            # Get server info
            server_info = self.server.server_info if hasattr(self.server, 'server_info') else {}
            server_info['page_title'] = 'About'
            
            # Get context
            context = get_default_context(server_info)
            
            # Render template
            html_content = self.template_loader.render('about.html', context, active_page='about')
            
            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(html_content.encode('utf-8'))))
            self.end_headers()
            self.wfile.write(html_content.encode('utf-8'))
            
        except Exception as e:
            self.serve_error_page(str(e))
    
    def formpageAnggotaKeluar(self):
        """Serve halaman form anggota keluar"""
        try:
            # Get server info
            server_info = self.server.server_info if hasattr(self.server, 'server_info') else {}
            server_info['page_title'] = 'Form Anggota Keluar'

            # Get context
            context = get_default_context(server_info)

            # Render template
            html_content = self.template_loader.render('form_anggota_keluar.html', context, active_page='form_anggota_keluar')

            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(html_content.encode('utf-8'))))
            self.end_headers()
            self.wfile.write(html_content.encode('utf-8'))

        except Exception as e:
            self.serve_error_page(str(e))
            
            # Get context
            context = get_default_context(server_info)
            
            # Render template
            html_content = self.template_loader.render('about.html', context, active_page='about')
            
            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(html_content.encode('utf-8'))))
            self.end_headers()
            self.wfile.write(html_content.encode('utf-8'))
            
        except Exception as e:
            self.serve_error_page(str(e))
    
    def serve_api_docs_page(self):
        """Serve halaman API documentation"""
        try:
            # Get server info
            server_info = self.server.server_info if hasattr(self.server, 'server_info') else {}
            server_info['page_title'] = 'API Documentation'
            
            # Get context
            context = get_default_context(server_info)
            
            # Render template
            html_content = self.template_loader.render('api_docs.html', context, active_page='api')
            
            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(html_content.encode('utf-8'))))
            self.end_headers()
            self.wfile.write(html_content.encode('utf-8'))
            
        except Exception as e:
            self.serve_error_page(str(e))
    
    def serve_qrcode(self):
        """Generate dan serve QR code untuk URL server"""
        try:
            # Get server info
            server_info = self.server.server_info if hasattr(self.server, 'server_info') else {}
            local_ip = server_info.get('local_ip', '127.0.0.1')
            port = server_info.get('port', 8080)
            
            # Create URL
            url = f"http://{local_ip}:{port}"
            
            # Generate QR code
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(url)
            qr.make(fit=True)
            
            # Create image
            img = qr.make_image(fill_color="black", back_color="white")
            
            # Save to bytes
            img_buffer = io.BytesIO()
            img.save(img_buffer, format='PNG')
            img_bytes = img_buffer.getvalue()
            
            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'image/png')
            self.send_header('Content-Length', str(len(img_bytes)))
            self.send_header('Cache-Control', 'no-cache')
            self.end_headers()
            self.wfile.write(img_bytes)
            
        except Exception as e:
            self.send_error(500, f"Error generating QR code: {str(e)}")
    
    def serve_api(self):
        """Serve API endpoints dengan JSON response"""
        try:
            # Get server info
            server_info = self.server.server_info if hasattr(self.server, 'server_info') else {}
            
            # API: /api/hello
            if self.path == '/api/hello':
                response = create_api_response(
                    success=True,
                    message='Halo Dunia!',
                    data={
                        'greeting': 'Hello World',
                        'description': 'API endpoint untuk testing'
                    }
                )
                self.send_json_response(response)
                return
            
            if self.path == '/api/test':
                response = create_api_response(
                    success=True,
                    message='Hallo test!',
                    data={
                        'greeting': 'Hello World',
                        'description': 'Berhasil'
                    }
                )
                self.send_json_response(response)
                return
            
            if self.path == '/api/data_center':
                response = create_api_response(
                    success=True,
                    message='data center!',
                    data=self.data_center()
                )
                self.send_json_response(response)
                return
            if self.path.startswith('/api/data_center/'):
                # Ambil kode center dari URL
                center = self.path.split('/')[-1]  # hasilnya "0001"
                
                response = create_api_response(
                    success=True,
                    message=f'data center {center}!',
                    data=self.data_center_anggota(center)
                )
                self.send_json_response(response)
                return
            if self.path.startswith('/api/anggota'):
                # Ambil kode anggota dari URL
                id_anggota = self.path.split('/')[-1]  # hasilnya "0001"
                
                response = create_api_response(
                    success=True,
                    message=f'data anggota {id_anggota}!',
                    data=self.data_center_anggotaByID(id_anggota)
                )
                self.send_json_response(response)
                return

            # API: /api/data_arsip_all - New optimized endpoint
            if self.path == '/api/data_arsip_all':
                data = self.get_data_arsip_all()
                if data.get('status') == 'error':
                    response = create_api_response(
                        success=False,
                        message=data.get('message', 'Failed to retrieve data'),
                        error=data.get('error', 'Unknown error')
                    )
                    self.send_json_response(response, status_code=500)
                else:
                    response = create_api_response(
                        success=True,
                        message='Data arsip retrieved successfully',
                        data=data
                    )
                    self.send_json_response(response)
                return
            
            # API: /api/data_arsip_summary - Get summary statistics
            if self.path == '/api/data_arsip_summary':
                response = create_api_response(
                    success=True,
                    message='Summary retrieved successfully',
                    data=self.get_data_arsip_summary()
                )
                self.send_json_response(response)
                return
                
            # API: /api/cache/status - Get cache status
            if self.path == '/api/cache/status':
                response = create_api_response(
                    success=True,
                    message='Cache status retrieved',
                    data={
                        'cached_items': len(data_cache.cache),
                        'max_size': data_cache.max_size,
                        'ttl_seconds': data_cache.ttl
                    }
                )
                self.send_json_response(response)
                return
                
            # API: /api/cache/clear - Clear cache
            if self.path == '/api/cache/clear':
                data_cache.invalidate()
                response = create_api_response(
                    success=True,
                    message='Cache cleared successfully',
                    data={'status': 'cleared'}
                )
                self.send_json_response(response)
                return
            
            # API: /api/status
            if self.path == '/api/status':
                response = create_api_response(
                    success=True,
                    message='Server running',
                    data={
                        'status': 'online',
                        'server_time': datetime.now().strftime('%d %B %Y, %H:%M:%S'),
                        'os_info': f"{platform.system()} {platform.release()}",
                        'python_version': platform.python_version(),
                        'hostname': platform.node(),
                        'local_ip': server_info.get('local_ip', 'N/A'),
                        'port': server_info.get('port', 'N/A')
                    }
                )
                self.send_json_response(response)
                return
            
            # API: /api/server-info
            if self.path == '/api/server-info':
                response = create_api_response(
                    success=True,
                    message='Server information retrieved',
                    data={
                        'server': {
                            'ip': server_info.get('local_ip', 'N/A'),
                            'port': server_info.get('port', 'N/A'),
                            'status': 'running'
                        },
                        'system': {
                            'os': f"{platform.system()} {platform.release()}",
                            'python': platform.python_version(),
                            'hostname': platform.node()
                        },
                        'folder': {
                            'default': server_info.get('default_folder', 'N/A')
                        }
                    }
                )
                self.send_json_response(response)
                return


            # API endpoint tidak ditemukan
            response = create_api_response(
                success=False,
                message='API endpoint not found',
                error=f'Endpoint {self.path} tidak tersedia'
            )
            self.send_json_response(response, status_code=404)
            
        except Exception as e:
            response = create_api_response(
                success=False,
                message='Internal server error',
                error=str(e)
            )
            self.send_json_response(response, status_code=500)
    
    def serve_api_post(self):
        """Serve API POST endpoints dengan JSON response"""
        try:
            # API: POST /api/anggota_keluar
            if self.path == '/api/anggota_keluar':
                response_data = self.handle_anggota_keluar()
                
                if response_data.get('success'):
                    response = create_api_response(
                        success=True,
                        message='Data anggota keluar berhasil diproses',
                        data=response_data.get('data', {})
                    )
                    self.send_json_response(response, status_code=201)
                else:
                    response = create_api_response(
                        success=False,
                        message=response_data.get('message', 'Failed to process data'),
                        error=response_data.get('error', 'Unknown error')
                    )
                    self.send_json_response(response, status_code=400)
                return
            
            # API endpoint POST tidak ditemukan
            response = create_api_response(
                success=False,
                message='POST API endpoint not found',
                error=f'POST endpoint {self.path} tidak tersedia'
            )
            self.send_json_response(response, status_code=404)
            
        except Exception as e:
            response = create_api_response(
                success=False,
                message='Internal server error',
                error=str(e)
            )
            self.send_json_response(response, status_code=500)
    
    def handle_anggota_keluar(self):
        """Handle POST data untuk anggota keluar (termasuk upload file & konversi gambar ke PDF)"""
        import os
        import json
        import shutil
        import email
        from datetime import datetime
        from calendar import month_name
        from PIL import Image  # === NEW ===

        try:
            content_type = self.headers.get('Content-Type', '')
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)

            # === Jika multipart/form-data (ada file) ===
            if 'multipart/form-data' in content_type:
                msg = email.message_from_bytes(
                    b"Content-Type: " + content_type.encode() + b"\r\n\r\n" + body
                )

                fields = {}
                file_path = None

                # === PARSE multipart ===
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue

                    disposition = part.get("Content-Disposition", "")
                    if not disposition:
                        continue

                    params = dict(
                        item.strip().split("=", 1)
                        for item in disposition.split(";")
                        if "=" in item
                    )
                    name = params.get("name", "").strip('"')

                    if "filename" in params:
                        filename = params["filename"].strip('"')
                        if filename:
                            upload_dir = os.path.join(os.getcwd(), "uploads")
                            os.makedirs(upload_dir, exist_ok=True)
                            unique_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}"
                            file_path = os.path.join(upload_dir, unique_name)
                            with open(file_path, "wb") as f:
                                f.write(part.get_payload(decode=True))
                    else:
                        fields[name] = part.get_payload(decode=True).decode("utf-8")

                # === Ambil data penting ===
                nomor_center = fields.get("nomor_center", "").strip()
                id_anggota = fields.get("id_anggota", "").strip()
                tanggal_keluar = fields.get("tanggal_keluar", "").strip()
                folder_asal = fields.get("folder", "").strip()

                # === Validasi wajib ===
                missing = [
                    k for k, v in {
                        "nomor_center": nomor_center,
                        "id_anggota": id_anggota,
                        "tanggal_keluar": tanggal_keluar,
                        "folder": folder_asal
                    }.items() if not v
                ]
                if missing:
                    return {
                        "success": False,
                        "message": f"Missing required fields: {', '.join(missing)}",
                    }

                # === Ambil tahun dan bulan ===
                dt = datetime.strptime(tanggal_keluar, "%Y-%m-%d")
                tahun = str(dt.year)
                bulan = f"{dt.month:02d}.{month_name[dt.month].upper()}"

                # === Nama anggota dari ID ===
                try:
                    _, nama_anggota = id_anggota.split("_", 1)
                except ValueError:
                    nama_anggota = id_anggota

                # === Siapkan folder tujuan ===
                base_dir = os.getcwd()
                parent_target_dir = os.path.join(base_dir, "03.DATA_ANGGOTA_KELUAR", tahun, bulan)
                target_dir = os.path.join(parent_target_dir, id_anggota)
                os.makedirs(parent_target_dir, exist_ok=True)

                # === Pindahkan isi folder asal ke folder tujuan ===
                if os.path.exists(folder_asal):
                    if os.path.exists(target_dir):
                        print(f"📂 Folder target sudah ada, merge isi...")
                        for item in os.listdir(folder_asal):
                            src = os.path.join(folder_asal, item)
                            dst = os.path.join(target_dir, item)
                            if os.path.isdir(src):
                                shutil.copytree(src, dst, dirs_exist_ok=True)
                            else:
                                shutil.copy2(src, dst)
                        shutil.rmtree(folder_asal, ignore_errors=True)
                    else:
                        shutil.move(folder_asal, target_dir)
                else:
                    print(f"⚠️ Folder asal tidak ditemukan: {folder_asal}")

                # === Jika ada file upload, ubah ke PDF jika gambar ===
                new_file_path = None
                if file_path:
                    try:
                        ext = os.path.splitext(file_path)[1].lower()
                        new_filename = f"12.{nama_anggota}.pdf"
                        new_file_path = os.path.join(target_dir, new_filename)

                        if ext in [".jpg", ".jpeg", ".png", ".bmp", ".gif"]:  # === NEW ===
                            image = Image.open(file_path)
                            if image.mode != 'RGB':
                                image = image.convert('RGB')
                            image.save(new_file_path, "PDF", resolution=100.0)
                            os.remove(file_path)  # hapus gambar asli
                            print(f"📄 Gambar dikonversi ke PDF: {new_file_path}")
                        else:
                            # Jika bukan gambar, pindahkan langsung
                            new_filename = f"12.{nama_anggota}{ext}"
                            new_file_path = os.path.join(target_dir, new_filename)
                            shutil.move(file_path, new_file_path)

                    except Exception as e:
                        print(f"⚠️ Gagal memproses file upload: {e}")
                    finally:
                        upload_dir = os.path.join(os.getcwd(), "uploads")
                        if os.path.exists(upload_dir):
                            shutil.rmtree(upload_dir, ignore_errors=True)

                print("✅ Proses selesai:")
                print(f"Center: {nomor_center}")
                print(f"Anggota: {id_anggota}")
                print(f"Tanggal: {tanggal_keluar}")
                print(f"Folder Asal: {folder_asal}")
                print(f"Folder Tujuan: {target_dir}")
                print(f"File Upload Baru: {new_file_path}")

                return {
                    "success": True,
                    "message": "Data anggota keluar berhasil diproses",
                    "data": {
                        "nomor_center": nomor_center,
                        "id_anggota": id_anggota,
                        "tanggal_keluar": tanggal_keluar,
                        "tahun": tahun,
                        "bulan": bulan,
                        "folder_tujuan": target_dir,
                        "file_path": new_file_path
                    }
                }

            else:
                # === Jika bukan multipart (JSON biasa) ===
                body_decoded = body.decode("utf-8")
                if "application/json" in content_type:
                    data = json.loads(body_decoded)
                else:
                    from urllib.parse import parse_qs
                    parsed = parse_qs(body_decoded)
                    data = {k: v[0] if v else "" for k, v in parsed.items()}

                return {
                    "success": True,
                    "message": "Data anggota keluar diterima tanpa file",
                    "data": data
                }

        except Exception as e:
            return {
                "success": False,
                "message": "Internal server error",
                "error": str(e)
            }

    
    def update_anggota_status(self, data):
        """Update status anggota di database (jika diperlukan)"""
        try:
            # Ini adalah placeholder untuk update database
            # Bisa diimplementasikan untuk update Excel atau database lain
            
            base_dir = os.path.dirname(os.path.abspath(__file__))
            db_file = os.path.join(base_dir, "database.xlsx")
            
            if not os.path.exists(db_file):
                return {
                    'success': False,
                    'message': 'Database file not found',
                    'updated': False
                }
            
            # For now, just return success without actual update
            # Implementasi update Excel bisa ditambahkan di sini
            return {
                'success': True,
                'message': 'Database update placeholder - ready for implementation',
                'updated': False,
                'note': 'Excel update can be implemented here if needed'
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'message': 'Failed to update database'
            }
    
    def data_center(self):
        """Get list of center numbers with caching"""
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            file_path_db = os.path.join(base_dir, "database.xlsx")
            
            if not os.path.exists(file_path_db):
                return []
            
            # Check cache
            file_mtime = os.path.getmtime(file_path_db)
            cache_key = f"data_center_{file_mtime}"
            cached_result = data_cache.get(cache_key)
            
            if cached_result is not None:
                return cached_result
            
            # Read from Excel
            sheet_name = "02.DATA_ANGGOTA"
            df = pd.read_excel(file_path_db, sheet_name=sheet_name)
            df = df.fillna('')
            
            if 'NOMOR_CENTER' not in df.columns:
                return []
                
            nomor_center = df['NOMOR_CENTER'].astype(str).str.zfill(4)
            nomor_center.drop_duplicates(inplace=True)
            result = nomor_center.tolist()
            
            # Cache the result
            data_cache.set(cache_key, result)
            
            return result
            
        except Exception as e:
            print(f"Error in data_center: {e}")
            return []
    
    def data_center_anggota(self, center):
        """Get member data for specific center with caching"""
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            file_path_db = os.path.join(base_dir, "database.xlsx")
            
            if not os.path.exists(file_path_db):
                return []
            
            # Check cache
            file_mtime = os.path.getmtime(file_path_db)
            cache_key = f"data_center_anggota_{center}_{file_mtime}"
            cached_result = data_cache.get(cache_key)
            
            if cached_result is not None:
                return cached_result

            sheet_name = "02.DATA_ANGGOTA"

            # Read Excel file
            df = pd.read_excel(file_path_db, sheet_name=sheet_name)
            df = df.fillna('')

            # Check required columns
            required_cols = ['NOMOR_CENTER', 'ID_NAMA_ANGGOTA', 'TYPE']
            if not all(col in df.columns for col in required_cols):
                return []

            # Get important columns & convert types
            anggota = df[required_cols].astype(str)
            anggota['NOMOR_CENTER'] = anggota['NOMOR_CENTER'].str.zfill(4)

            # Remove duplicates
            anggota_unik = anggota.drop_duplicates(ignore_index=True)

            # Filter by center and FILE type
            filter_center = anggota_unik[
                (anggota_unik['NOMOR_CENTER'] == center) &
                (anggota_unik['TYPE'] == 'FILE')
            ].copy()

            result = filter_center.to_dict(orient="records")
            
            # Cache the result
            data_cache.set(cache_key, result)
            
            return result
            
        except Exception as e:
            print(f"Error in data_center_anggota: {e}")
            return []
    def data_center_anggotaByID(self, id):
        """Get member data by ID with file count and caching"""
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            file_path_db = os.path.join(base_dir, "database.xlsx")
            
            if not os.path.exists(file_path_db):
                return []
            
            # Check cache
            file_mtime = os.path.getmtime(file_path_db)
            cache_key = f"data_center_anggotaByID_{id}_{file_mtime}"
            cached_result = data_cache.get(cache_key)
            
            if cached_result is not None:
                return cached_result

            sheet_name = "02.DATA_ANGGOTA"

            # Read Excel file
            df = pd.read_excel(file_path_db, sheet_name=sheet_name)
            df = df.fillna('')

            # Check required columns
            required_cols = ['ID_NAMA_ANGGOTA', 'TYPE']
            if not all(col in df.columns for col in required_cols):
                return []

            # Find specific folder
            cari_anggota = df[
                (df['ID_NAMA_ANGGOTA'] == id) &
                (df['TYPE'] == 'FOLDER')
            ]

            if cari_anggota.empty:
                result = []
            else:
                # Get folder data
                data_ketemu = df.loc[cari_anggota.index].copy()

                # Count files with same member ID
                hitung_file = df[
                    (df['ID_NAMA_ANGGOTA'] == id) &
                    (df['TYPE'] == 'FILE')
                ]
                total_file = len(hitung_file)
                data_ketemu["TOTAL_FILE"] = total_file
                result = data_ketemu.to_dict(orient="records")
            
            # Cache the result
            data_cache.set(cache_key, result)
            
            return result
            
        except Exception as e:
            print(f"Error in data_center_anggotaByID: {e}")
            return []

    def get_data_arsip_all(self):
        """Optimized method to get all archive data from database.xlsx with caching"""
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
            # Check for database.xlsx in root project first
            db_files = [
                os.path.join(base_dir, "database.xlsx"),
                os.path.join(base_dir, "arsip_database.xlsx"),  # fallback name
                os.path.join(base_dir, "data_arsip.xlsx")        # alternative name
            ]
            
            db_file = None
            for file_path in db_files:
                if os.path.exists(file_path):
                    db_file = file_path
                    break
            
            if not db_file:
                return {
                    "error": "Database file not found",
                    "message": "Please generate database first using Universal Scan feature",
                    "expected_files": [os.path.basename(f) for f in db_files]
                }
            
            # Get file modification time for cache validation
            file_mtime = os.path.getmtime(db_file)
            cache_key = f"data_arsip_all_{file_mtime}"
            
            # Check cache first
            cached_data = data_cache.get(cache_key)
            if cached_data is not None:
                return cached_data
            
            # Read Excel with optimizations
            try:
                # Try to read the main sheet first
                df = pd.read_excel(db_file, sheet_name=0, engine='openpyxl')
            except Exception as e:
                # Try alternative sheet names
                sheet_names = ["Data_Arsip", "Sheet1", "Arsip", "Database"]
                df = None
                last_error = str(e)
                
                for sheet_name in sheet_names:
                    try:
                        df = pd.read_excel(db_file, sheet_name=sheet_name, engine='openpyxl')
                        break
                    except:
                        continue
                
                if df is None:
                    return {
                        "error": "Cannot read database sheet",
                        "message": f"Failed to read any sheet. Last error: {last_error}",
                        "file_path": os.path.basename(db_file)
                    }
            
            # Clean and optimize data
            df = df.fillna('')
            
            # Convert to records with optimized memory usage
            records = []
            for _, row in df.iterrows():
                record = {}
                for col in df.columns:
                    value = row[col]
                    # Convert numpy types to native Python types for JSON serialization
                    if hasattr(value, 'item'):
                        value = value.item()
                    elif hasattr(value, '__class__') and 'Timestamp' in str(type(value)):
                        value = str(value) if pd.notna(value) else ''
                    elif pd.isna(value):
                        value = ''
                    record[col] = value
                records.append(record)
            
            # Prepare optimized response with metadata
            result = {
                "data": records,
                "metadata": {
                    "total_records": len(records),
                    "columns": list(df.columns),
                    "file_path": os.path.basename(db_file),
                    "last_modified": file_mtime,
                    "file_size_mb": round(os.path.getsize(db_file) / 1024 / 1024, 2),
                    "generated_at": datetime.now().isoformat(),
                    "cache_enabled": True
                },
                "status": "success"
            }
            
            # Cache the result
            data_cache.set(cache_key, result)
            
            return result
            
        except Exception as e:
            return {
                "error": "Failed to retrieve archive data",
                "message": str(e),
                "status": "error"
            }
    
    def get_data_arsip_summary(self):
        """Get summary statistics of archive data"""
        try:
            # Try to get data from cache or database
            data_result = self.get_data_arsip_all()
            
            if data_result.get('status') == 'error':
                return data_result
            
            records = data_result.get('data', [])
            metadata = data_result.get('metadata', {})
            
            if not records:
                return {
                    "total_files": 0,
                    "total_folders": 0,
                    "total_size_mb": 0,
                    "file_extensions": {},
                    "status": "success"
                }
            
            # Calculate statistics
            total_files = 0
            total_folders = 0
            total_size_bytes = 0
            extensions = {}
            
            for record in records:
                # Check if it's file or folder based on common column names
                item_type = ''
                if 'TYPE' in record:
                    item_type = str(record['TYPE']).upper()
                elif 'Type' in record:
                    item_type = str(record['Type']).upper()
                elif 'type' in record:
                    item_type = str(record['type']).upper()
                
                if item_type == 'FILE':
                    total_files += 1
                    # Try to get file size
                    size_value = 0
                    for size_col in ['Size', 'SIZE', 'Ukuran', 'File_Size']:
                        if size_col in record and record[size_col]:
                            try:
                                size_str = str(record[size_col]).replace(',', '').replace(' ', '')
                                if size_str.replace('.', '').isdigit():
                                    size_value = float(size_str)
                                    break
                            except:
                                pass
                    total_size_bytes += size_value
                    
                    # Get file extension
                    filename = ''
                    for name_col in ['Nama', 'Name', 'Filename', 'File_Name']:
                        if name_col in record and record[name_col]:
                            filename = str(record[name_col])
                            break
                    
                    if filename and '.' in filename:
                        ext = filename.split('.')[-1].lower()
                        extensions[ext] = extensions.get(ext, 0) + 1
                        
                elif item_type == 'FOLDER':
                    total_folders += 1
            
            # Convert bytes to MB
            total_size_mb = round(total_size_bytes / (1024 * 1024), 2) if total_size_bytes > 0 else 0
            
            return {
                "total_files": total_files,
                "total_folders": total_folders,
                "total_items": len(records),
                "total_size_mb": total_size_mb,
                "file_extensions": dict(sorted(extensions.items(), key=lambda x: x[1], reverse=True)[:10]),  # Top 10
                "database_info": metadata,
                "status": "success"
            }
            
        except Exception as e:
            return {
                "error": "Failed to generate summary",
                "message": str(e),
                "status": "error"
            }


    def send_json_response(self, response_dict, status_code=200, enable_compression=True):
        """Send optimized JSON response with optional gzip compression"""
        json_content = json_response(response_dict)
        json_bytes = json_content.encode('utf-8')
        
        # Check if client accepts gzip compression
        accept_encoding = self.headers.get('Accept-Encoding', '')
        use_gzip = enable_compression and 'gzip' in accept_encoding.lower() and len(json_bytes) > 1024
        
        if use_gzip:
            # Compress response
            json_bytes = gzip.compress(json_bytes)
            
        self.send_response(status_code)
        self.send_header('Content-type', 'application/json; charset=utf-8')
        
        if use_gzip:
            self.send_header('Content-Encoding', 'gzip')
            
        # Add cache headers for API responses
        if status_code == 200:
            self.send_header('Cache-Control', 'public, max-age=60')  # Cache for 1 minute
            
        # Add CORS headers for API access
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Accept, Accept-Encoding')
        
        self.send_header('Content-Length', str(len(json_bytes)))
        self.end_headers()
        self.wfile.write(json_bytes)

    
    def serve_error_page(self, error_message):
        """Serve halaman error"""
        try:
            context = {'error_message': error_message}
            html_content = self.template_loader.render('error.html', context)
            
            self.send_response(500)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(html_content.encode('utf-8'))
        except:
            # Fallback jika template error juga gagal
            fallback_html = f"""
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>Error</title></head>
<body>
    <h1>Error</h1>
    <pre>{error_message}</pre>
</body>
</html>
            """
            self.send_response(500)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(fallback_html.encode('utf-8'))



class WebServerManager:
    """Manager untuk mengelola web server sederhana"""
    
    def __init__(self, config_file="app_config.json"):
        self.server = None
        self.server_thread = None
        self.is_running = False
        self.config_file = config_file
        self.port = self.get_web_server_port()
        self.default_folder = None
    
    def load_config(self):
        """Load konfigurasi dari file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return {}
        except Exception:
            return {}
    
    def get_web_server_port(self):
        """Get web server port dari config"""
        config = self.load_config()
        return config.get("web_server_port", 8080)
    
    def get_default_folder(self):
        """Get default folder dari config"""
        config = self.load_config()
        return config.get("default_folder", "")
    
    def get_local_ip(self):
        """Dapatkan IP lokal komputer"""
        try:
            # Buat socket untuk mendapatkan IP
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.settimeout(0.1)
            try:
                # Tidak perlu benar-benar connect, hanya untuk get IP
                s.connect(("8.8.8.8", 80))
                local_ip = s.getsockname()[0]
            except Exception:
                local_ip = "127.0.0.1"
            finally:
                s.close()
            return local_ip
        except Exception:
            return "127.0.0.1"
    
    def start_server(self, port=None):
        """Start web server"""
        if self.is_running:
            return False, "Server sudah berjalan"
        
        if port:
            self.port = port
        
        try:
            # Get local IP first
            local_ip = self.get_local_ip()
            
            # Get default folder path
            default_folder = self.get_default_folder()
            
            if not default_folder or not os.path.exists(default_folder):
                return False, "Default folder belum diset atau tidak valid.\nSilakan set default folder di Pengaturan terlebih dahulu."
            
            # Save current directory
            self.original_dir = os.getcwd()
            
            # Change directory ke default folder
            try:
                os.chdir(default_folder)
                self.default_folder = default_folder
            except Exception as e:
                return False, f"Gagal akses folder: {str(e)}"
            
            # Create server with HelloWorldHandler
            handler = HelloWorldHandler
            try:
                self.server = HTTPServer(("0.0.0.0", self.port), handler)
                
                # Set server info untuk ditampilkan di halaman
                self.server.server_info = {
                    'local_ip': local_ip,
                    'port': self.port,
                    'default_folder': default_folder
                }
            except OSError as e:
                # Restore directory
                os.chdir(self.original_dir)
                if "Address already in use" in str(e) or "WinError 10048" in str(e):
                    return False, f"Port {self.port} sudah digunakan.\nCoba gunakan port lain (misal: 8081, 8888, 9000)"
                return False, f"Error binding port: {str(e)}"
            
            # Start server di thread terpisah
            self.server_thread = threading.Thread(target=self.server.serve_forever, daemon=True)
            self.server_thread.start()
            
            self.is_running = True
            local_ip = self.get_local_ip()
            
            success_msg = (
                f"✅ Server berhasil dijalankan!\n\n"
                f"📁 Folder: {os.path.basename(default_folder)}\n"
                f"🌐 Port: {self.port}\n\n"
                f"URL Akses:\n"
                f"• Lokal: http://localhost:{self.port}\n"
                f"• Network: http://{local_ip}:{self.port}\n\n"
                f"💡 Buka URL di browser untuk akses file"
            )
            
            return True, success_msg
            
        except Exception as e:
            # Cleanup jika ada error
            if hasattr(self, 'original_dir'):
                try:
                    os.chdir(self.original_dir)
                except:
                    pass
            return False, f"Gagal start server: {str(e)}"
    
    def stop_server(self):
        """Stop web server"""
        if not self.is_running:
            return False, "Server tidak sedang berjalan"
        
        try:
            # Set flag dulu
            self.is_running = False
            
            # Shutdown server
            if self.server:
                try:
                    self.server.shutdown()
                except Exception as e:
                    print(f"Error saat shutdown: {e}")
                
                try:
                    self.server.server_close()
                except Exception as e:
                    print(f"Error saat close: {e}")
                
                self.server = None
            
            # Wait thread selesai (dengan timeout)
            if self.server_thread and self.server_thread.is_alive():
                self.server_thread.join(timeout=2.0)
            
            self.server_thread = None
            
            # Restore directory
            if hasattr(self, 'original_dir'):
                try:
                    os.chdir(self.original_dir)
                except Exception as e:
                    print(f"Error restore directory: {e}")
            
            return True, "✅ Server berhasil dihentikan"
            
        except Exception as e:
            # Pastikan flag diset meskipun error
            self.is_running = False
            self.server = None
            self.server_thread = None
            return False, f"Error saat stop server: {str(e)}\n\nServer sudah dihentikan paksa."
    
    def get_server_info(self):
        """Get server info"""
        local_ip = self.get_local_ip()
        
        return {
            "status": "Running" if self.is_running else "Stopped",
            "is_running": self.is_running,
            "port": self.port,
            "url_local": f"http://localhost:{self.port}",
            "url_network": f"http://{local_ip}:{self.port}",
            "local_ip": local_ip,
            "default_folder": self.default_folder if self.is_running else self.get_default_folder()
        }
    
    def __del__(self):
        """Destructor - pastikan server dihentikan"""
        if self.is_running:
            try:
                self.stop_server()
            except:
                pass


# Singleton instance
_web_server_instance = None

def get_web_server_manager():
    """Get singleton instance of WebServerManager"""
    global _web_server_instance
    if _web_server_instance is None:
        _web_server_instance = WebServerManager()
    return _web_server_instance
