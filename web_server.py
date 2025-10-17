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
import io

# QR Code library
import qrcode

# Import template loader
from src_web.template_loader import TemplateLoader, get_default_context, create_api_response, json_response


class HelloWorldHandler(BaseHTTPRequestHandler):
    """Custom HTTP Request Handler dengan template system dan API support"""
    
    def __init__(self, *args, **kwargs):
        self.template_loader = TemplateLoader()
        super().__init__(*args, **kwargs)
    
    def log_message(self, format, *args):
        """Override log_message untuk suppress console output"""
        pass
    
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
    
    def send_json_response(self, response_dict, status_code=200):
        """Send JSON response"""
        json_content = json_response(response_dict)
        
        self.send_response(status_code)
        self.send_header('Content-type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(json_content.encode('utf-8'))))
        self.end_headers()
        self.wfile.write(json_content.encode('utf-8'))

    
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
