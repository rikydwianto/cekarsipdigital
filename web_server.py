"""
Web Server Module untuk Arsip Owncloud
Menyediakan HTTP server sederhana untuk akses file arsip melalui browser
"""

import os
import socket
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
import json
from datetime import datetime
import platform


class HelloWorldHandler(BaseHTTPRequestHandler):
    """Custom HTTP Request Handler untuk Hello World"""
    
    def log_message(self, format, *args):
        """Override log_message untuk suppress console output"""
        pass
    
    def do_GET(self):
        """Handle GET request"""
        try:
            # Get server info
            server_info = self.server.server_info if hasattr(self.server, 'server_info') else {}
            
            # HTML response
            html_content = f"""
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Arsip Owncloud - Web Server</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }}
        
        .container {{
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 40px;
            max-width: 600px;
            width: 100%;
            animation: fadeIn 0.5s ease-in;
        }}
        
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(-20px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        h1 {{
            color: #667eea;
            text-align: center;
            margin-bottom: 10px;
            font-size: 2.5em;
        }}
        
        .emoji {{
            font-size: 3em;
            text-align: center;
            margin-bottom: 20px;
        }}
        
        .subtitle {{
            text-align: center;
            color: #666;
            margin-bottom: 30px;
            font-size: 1.1em;
        }}
        
        .info-box {{
            background: #f8f9fa;
            border-left: 4px solid #667eea;
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
        }}
        
        .info-label {{
            font-weight: bold;
            color: #333;
            margin-bottom: 5px;
        }}
        
        .info-value {{
            color: #666;
            font-family: 'Courier New', monospace;
            font-size: 0.95em;
        }}
        
        .status {{
            background: #d4edda;
            color: #155724;
            padding: 10px;
            border-radius: 5px;
            text-align: center;
            margin: 20px 0;
            font-weight: bold;
        }}
        
        .footer {{
            text-align: center;
            margin-top: 30px;
            color: #999;
            font-size: 0.9em;
        }}
        
        .badge {{
            display: inline-block;
            padding: 5px 10px;
            background: #667eea;
            color: white;
            border-radius: 15px;
            font-size: 0.85em;
            margin: 5px;
        }}
        
        .pulse {{
            animation: pulse 2s infinite;
        }}
        
        @keyframes pulse {{
            0%, 100% {{ opacity: 1; }}
            50% {{ opacity: 0.5; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="emoji">🌐</div>
        <h1>Hello World!</h1>
        <p class="subtitle">Arsip Owncloud Web Server</p>
        
        <div class="status pulse">
            ✅ Server Aktif dan Berjalan
        </div>
        
        <div class="info-box">
            <div class="info-label">📅 Waktu Server</div>
            <div class="info-value">{datetime.now().strftime('%d %B %Y, %H:%M:%S')}</div>
        </div>
        
        <div class="info-box">
            <div class="info-label">🖥️ Sistem Operasi</div>
            <div class="info-value">{platform.system()} {platform.release()}</div>
        </div>
        
        <div class="info-box">
            <div class="info-label">🐍 Python Version</div>
            <div class="info-value">{platform.python_version()}</div>
        </div>
        
        <div class="info-box">
            <div class="info-label">💻 Hostname</div>
            <div class="info-value">{platform.node()}</div>
        </div>
        
        <div class="info-box">
            <div class="info-label">🌐 IP Address</div>
            <div class="info-value">{server_info.get('local_ip', 'N/A')}</div>
        </div>
        
        <div class="info-box">
            <div class="info-label">🔌 Port</div>
            <div class="info-value">{server_info.get('port', 'N/A')}</div>
        </div>
        
        <div class="info-box">
            <div class="info-label">📁 Document Root</div>
            <div class="info-value">{server_info.get('default_folder', 'N/A')}</div>
        </div>
        
        <div style="text-align: center; margin-top: 20px;">
            <span class="badge">Web Server</span>
            <span class="badge">Python {platform.python_version()}</span>
            <span class="badge">Arsip Digital</span>
        </div>
        
        <div class="footer">
            <p>🚀 Powered by Arsip Owncloud Application</p>
            <p style="margin-top: 5px;">💡 Fitur browse file akan segera hadir!</p>
        </div>
    </div>
</body>
</html>
            """
            
            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(html_content.encode('utf-8'))))
            self.end_headers()
            self.wfile.write(html_content.encode('utf-8'))
            
        except Exception as e:
            # Error response
            error_html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Error</title>
    <style>
        body {{ font-family: Arial; padding: 40px; background: #f5f5f5; }}
        .error {{ background: white; padding: 30px; border-radius: 10px; max-width: 600px; margin: 0 auto; }}
        h1 {{ color: #d32f2f; }}
    </style>
</head>
<body>
    <div class="error">
        <h1>❌ Error</h1>
        <p>Terjadi kesalahan saat memproses request:</p>
        <pre>{str(e)}</pre>
    </div>
</body>
</html>
            """
            self.send_response(500)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(error_html.encode('utf-8'))


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
