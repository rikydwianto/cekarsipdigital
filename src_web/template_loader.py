"""
Template Loader untuk Web Server
Sistem sederhana untuk load dan render template HTML dengan support partial
"""

import os
import json
from datetime import datetime
import platform


class TemplateLoader:
    """Simple template loader dengan sistem replace variable dan partial support"""
    
    def __init__(self, template_dir=None):
        if template_dir is None:
            # Get path relatif dari file ini
            current_dir = os.path.dirname(os.path.abspath(__file__))
            template_dir = os.path.join(current_dir, 'templates')
        
        self.template_dir = template_dir
        self.partials_dir = os.path.join(template_dir, 'partials')
    
    def load_template(self, template_name):
        """Load template file"""
        template_path = os.path.join(self.template_dir, template_name)
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template tidak ditemukan: {template_name}")
        
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    
    def load_partial(self, partial_name):
        """Load partial template file"""
        partial_path = os.path.join(self.partials_dir, partial_name)
        
        if not os.path.exists(partial_path):
            raise FileNotFoundError(f"Partial tidak ditemukan: {partial_name}")
        
        with open(partial_path, 'r', encoding='utf-8') as f:
            return f.read()
    
    def render(self, template_name, context=None, active_page='home'):
        """Render template dengan context variables dan partials"""
        template_content = self.load_template(template_name)
        
        if context is None:
            context = {}
        
        # Load partials jika ada
        try:
            header_content = self.load_partial('header.html')
            footer_content = self.load_partial('footer.html')
            
            # Set active menu
            context['active_home'] = 'active' if active_page == 'home' else ''
            context['active_about'] = 'active' if active_page == 'about' else ''
            context['active_api'] = 'active' if active_page == 'api' else ''
            
            # Replace partials dulu
            template_content = template_content.replace('{{header}}', header_content)
            template_content = template_content.replace('{{footer}}', footer_content)
        except FileNotFoundError:
            # Jika partial tidak ada, lanjutkan tanpa partial
            pass
        
        # Replace semua {{variable}} dengan nilai dari context
        for key, value in context.items():
            placeholder = f"{{{{{key}}}}}"
            template_content = template_content.replace(placeholder, str(value))
        
        return template_content
    
    def get_static_file(self, file_path):
        """Get static file content"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        static_dir = os.path.join(current_dir, 'static')
        
        # Security: prevent directory traversal
        safe_path = os.path.normpath(file_path).lstrip(os.sep).lstrip('/')
        full_path = os.path.join(static_dir, safe_path)
        
        # Pastikan file berada di dalam static dir
        if not full_path.startswith(static_dir):
            return None, None
        
        if not os.path.exists(full_path):
            return None, None
        
        # Detect content type
        ext = os.path.splitext(full_path)[1].lower()
        content_types = {
            '.css': 'text/css',
            '.js': 'application/javascript',
            '.html': 'text/html',
            '.json': 'application/json',
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.svg': 'image/svg+xml',
            '.ico': 'image/x-icon'
        }
        
        content_type = content_types.get(ext, 'application/octet-stream')
        
        # Read file
        mode = 'rb' if content_type.startswith('image') else 'r'
        encoding = None if mode == 'rb' else 'utf-8'
        
        with open(full_path, mode, encoding=encoding) as f:
            content = f.read()
        
        return content, content_type


def get_default_context(server_info=None):
    """Get context default untuk template"""
    if server_info is None:
        server_info = {}
    
    return {
        'page_title': server_info.get('page_title', 'Web Server'),
        'server_time': datetime.now().strftime('%d %B %Y, %H:%M:%S'),
        'os_info': f"{platform.system()} {platform.release()}",
        'python_version': platform.python_version(),
        'hostname': platform.node(),
        'local_ip': server_info.get('local_ip', 'N/A'),
        'port': server_info.get('port', 'N/A'),
        'default_folder': server_info.get('default_folder', 'N/A'),
        'author': server_info.get('author', 'RIKY DWIANTO')
    }


def create_api_response(success=True, message='', data=None, error=None):
    """
    Membuat response JSON yang konsisten untuk API
    
    Args:
        success (bool): Status success atau error
        message (str): Pesan deskriptif
        data (dict): Data yang akan dikembalikan (untuk success)
        error (str): Detail error (untuk error response)
    
    Returns:
        dict: Response dictionary yang siap di-encode ke JSON
    """
    response = {
        'success': success,
        'message': message,
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    if success:
        if data is not None:
            response['data'] = data
    else:
        if error is not None:
            response['error'] = error
    
    return response


def json_response(response_dict):
    """
    Convert dictionary ke JSON string dengan formatting yang bagus
    
    Args:
        response_dict (dict): Dictionary response
    
    Returns:
        str: JSON string yang sudah di-format
    """
    return json.dumps(response_dict, indent=2, ensure_ascii=False)
