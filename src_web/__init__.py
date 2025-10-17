"""
Arsip Owncloud Web Server Package
Menyediakan template system dan static files untuk web server
"""

from .template_loader import TemplateLoader, get_default_context

__all__ = ['TemplateLoader', 'get_default_context']
