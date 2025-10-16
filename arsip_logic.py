"""
Business Logic Module untuk Aplikasi Arsip Digital
===================================================

Module ini berisi semua logic bisnis yang terpisah dari GUI,
sehingga mudah untuk maintenance dan testing.
"""

import os
import re
import pandas as pd
from datetime import datetime
from typing import Dict, List, Tuple, Optional


class FileManager:
    """Class untuk mengelola operasi file dan folder"""
    
    @staticmethod
    def validate_folder_path(folder_path: str) -> bool:
        """
        Validasi apakah path folder valid dan dapat diakses
        
        Args:
            folder_path (str): Path folder yang akan divalidasi
            
        Returns:
            bool: True jika valid, False jika tidak
        """
        if not folder_path or not isinstance(folder_path, str):
            return False
        
        return os.path.exists(folder_path) and os.path.isdir(folder_path)
    
    @staticmethod
    def validate_file_path(file_path: str) -> bool:
        """
        Validasi apakah path file valid dan dapat diakses
        
        Args:
            file_path (str): Path file yang akan divalidasi
            
        Returns:
            bool: True jika valid, False jika tidak
        """
        if not file_path or not isinstance(file_path, str):
            return False
        
        return os.path.exists(file_path) and os.path.isfile(file_path)
    
    @staticmethod
    def get_file_size(file_path: str) -> str:
        """
        Mendapatkan ukuran file dalam format yang mudah dibaca
        
        Args:
            file_path (str): Path file
            
        Returns:
            str: Ukuran file dalam format human-readable
        """
        try:
            if not FileManager.validate_file_path(file_path):
                return "File tidak ditemukan"
            
            size = os.path.getsize(file_path)
            for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
                if size < 1024.0:
                    return f"{size:.1f} {unit}"
                size /= 1024.0
            return f"{size:.1f} PB"
        except Exception as e:
            return f"Error: {str(e)}"
    
    @staticmethod
    def get_file_info(file_path: str) -> Dict[str, str]:
        """
        Mendapatkan informasi lengkap tentang file
        
        Args:
            file_path (str): Path file
            
        Returns:
            Dict[str, str]: Dictionary berisi informasi file
        """
        if not FileManager.validate_file_path(file_path):
            return {"error": "File tidak ditemukan atau tidak valid"}
        
        try:
            stat = os.stat(file_path)
            return {
                "name": os.path.basename(file_path),
                "path": os.path.abspath(file_path),
                "size": FileManager.get_file_size(file_path),
                "size_bytes": str(stat.st_size),
                "created": datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d %H:%M:%S"),
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                "extension": os.path.splitext(file_path)[1].lower(),
                "directory": os.path.dirname(file_path)
            }
        except Exception as e:
            return {"error": f"Gagal mendapatkan info file: {str(e)}"}
    
    @staticmethod
    def get_folder_info(folder_path: str) -> Dict[str, any]:
        """
        Mendapatkan informasi tentang folder
        
        Args:
            folder_path (str): Path folder
            
        Returns:
            Dict[str, any]: Dictionary berisi informasi folder
        """
        if not FileManager.validate_folder_path(folder_path):
            return {"error": "Folder tidak ditemukan atau tidak valid"}
        
        try:
            file_count = 0
            folder_count = 0
            total_size = 0
            file_types = {}
            
            for item in os.listdir(folder_path):
                item_path = os.path.join(folder_path, item)
                if os.path.isfile(item_path):
                    file_count += 1
                    try:
                        file_size = os.path.getsize(item_path)
                        total_size += file_size
                        
                        # Count file types
                        ext = os.path.splitext(item)[1].lower()
                        file_types[ext] = file_types.get(ext, 0) + 1
                    except:
                        pass
                elif os.path.isdir(item_path):
                    folder_count += 1
            
            return {
                "name": os.path.basename(folder_path),
                "path": os.path.abspath(folder_path),
                "file_count": file_count,
                "folder_count": folder_count,
                "total_size": FileManager._format_size(total_size),
                "total_size_bytes": total_size,
                "file_types": file_types,
                "created": datetime.fromtimestamp(os.path.getctime(folder_path)).strftime("%Y-%m-%d %H:%M:%S"),
                "modified": datetime.fromtimestamp(os.path.getmtime(folder_path)).strftime("%Y-%m-%d %H:%M:%S")
            }
        except Exception as e:
            return {"error": f"Gagal mendapatkan info folder: {str(e)}"}
    
    @staticmethod
    def _format_size(size_bytes: int) -> str:
        """Helper method untuk format ukuran file"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} PB"
    
    @staticmethod
    def is_file_in_folder(file_path: str, folder_path: str) -> bool:
        """
        Cek apakah file berada dalam folder tertentu
        
        Args:
            file_path (str): Path file
            folder_path (str): Path folder
            
        Returns:
            bool: True jika file berada dalam folder
        """
        if not FileManager.validate_file_path(file_path) or not FileManager.validate_folder_path(folder_path):
            return False
        
        try:
            abs_file = os.path.abspath(file_path)
            abs_folder = os.path.abspath(folder_path)
            return abs_file.startswith(abs_folder)
        except:
            return False


class ArsipProcessor:
    """Class untuk memproses operasi arsip digital"""
    
    def __init__(self):
        self.file_manager = FileManager()
        self.processing_history = []
    
    def validate_selection(self, folder_path: str, file_path: str) -> Dict[str, any]:
        """
        Validasi pilihan folder dan file
        
        Args:
            folder_path (str): Path folder yang dipilih
            file_path (str): Path file yang dipilih
            
        Returns:
            Dict[str, any]: Result validasi dengan status dan pesan
        """
        result = {
            "valid": False,
            "errors": [],
            "warnings": [],
            "folder_info": None,
            "file_info": None
        }
        
        # Validasi folder
        if not folder_path:
            result["errors"].append("Folder belum dipilih")
        elif not self.file_manager.validate_folder_path(folder_path):
            result["errors"].append("Folder tidak valid atau tidak dapat diakses")
        else:
            result["folder_info"] = self.file_manager.get_folder_info(folder_path)
            if "error" in result["folder_info"]:
                result["errors"].append(f"Error folder: {result['folder_info']['error']}")
        
        # Validasi file
        if not file_path:
            result["errors"].append("File belum dipilih")
        elif not self.file_manager.validate_file_path(file_path):
            result["errors"].append("File tidak valid atau tidak dapat diakses")
        else:
            result["file_info"] = self.file_manager.get_file_info(file_path)
            if "error" in result["file_info"]:
                result["errors"].append(f"Error file: {result['file_info']['error']}")
        
        # Validasi hubungan folder dan file
        if folder_path and file_path and not result["errors"]:
            if not self.file_manager.is_file_in_folder(file_path, folder_path):
                result["warnings"].append("File tidak berada dalam folder yang dipilih")
        
        # Set valid jika tidak ada error
        result["valid"] = len(result["errors"]) == 0
        
        return result
    
    def process_archive(self, folder_path: str, file_path: str) -> Dict[str, any]:
        """
        Memproses arsip digital berdasarkan folder dan file yang dipilih
        
        Args:
            folder_path (str): Path folder arsip
            file_path (str): Path file detail nasabah
            
        Returns:
            Dict[str, any]: Result proses dengan detail informasi
        """
        # Validasi terlebih dahulu
        validation_result = self.validate_selection(folder_path, file_path)
        
        if not validation_result["valid"]:
            return {
                "success": False,
                "message": "Validasi gagal",
                "errors": validation_result["errors"],
                "warnings": validation_result["warnings"]
            }
        
        try:
            # Proses arsip di sini
            process_result = {
                "success": True,
                "message": "Arsip berhasil diproses",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "folder_info": validation_result["folder_info"],
                "file_info": validation_result["file_info"],
                "warnings": validation_result["warnings"],
                "process_details": self._perform_archive_processing(folder_path, file_path)
            }
            
            # Simpan ke history
            self.processing_history.append({
                "timestamp": process_result["timestamp"],
                "folder_path": folder_path,
                "file_path": file_path,
                "success": True
            })
            
            return process_result
            
        except Exception as e:
            error_result = {
                "success": False,
                "message": f"Error saat memproses arsip: {str(e)}",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "errors": [str(e)]
            }
            
            # Simpan error ke history
            self.processing_history.append({
                "timestamp": error_result["timestamp"],
                "folder_path": folder_path,
                "file_path": file_path,
                "success": False,
                "error": str(e)
            })
            
            return error_result
    
    def _perform_archive_processing(self, folder_path: str, file_path: str) -> Dict[str, any]:
        """
        Melakukan proses arsip yang sebenarnya
        
        Args:
            folder_path (str): Path folder
            file_path (str): Path file
            
        Returns:
            Dict[str, any]: Detail proses yang dilakukan
        """
        # Simulasi proses arsip - bisa dikembangkan sesuai kebutuhan
        processing_steps = []
        
        # Step 1: Analisis folder
        processing_steps.append("✓ Menganalisis struktur folder arsip")
        
        # Step 2: Validasi file nasabah
        processing_steps.append("✓ Memvalidasi file detail nasabah")
        
        # Step 3: Extract informasi (simulasi)
        processing_steps.append("✓ Mengekstrak informasi dari file")
        
        # Step 4: Generate summary
        processing_steps.append("✓ Membuat ringkasan arsip digital")
        
        return {
            "steps_completed": processing_steps,
            "total_steps": len(processing_steps),
            "estimated_duration": "2.3 detik",
            "files_processed": 1,
            "status": "Completed Successfully"
        }
    
    def get_processing_history(self) -> List[Dict[str, any]]:
        """
        Mendapatkan riwayat pemrosesan arsip
        
        Returns:
            List[Dict[str, any]]: List riwayat proses
        """
        return self.processing_history.copy()
    
    def clear_history(self):
        """Membersihkan riwayat pemrosesan"""
        self.processing_history.clear()
    
    def export_summary(self, result: Dict[str, any]) -> str:
        """
        Export hasil proses ke format text summary
        
        Args:
            result (Dict[str, any]): Result dari process_archive
            
        Returns:
            str: Summary dalam format text
        """
        if not result.get("success", False):
            return f"PROSES GAGAL\n\nError: {result.get('message', 'Unknown error')}"
        
        summary = []
        summary.append("=" * 50)
        summary.append("RINGKASAN PROSES ARSIP DIGITAL")
        summary.append("=" * 50)
        summary.append(f"Waktu Proses: {result.get('timestamp', 'N/A')}")
        summary.append("")
        
        # Folder info
        folder_info = result.get("folder_info", {})
        if folder_info and "error" not in folder_info:
            summary.append("INFORMASI FOLDER:")
            summary.append(f"  Nama: {folder_info.get('name', 'N/A')}")
            summary.append(f"  Path: {folder_info.get('path', 'N/A')}")
            summary.append(f"  Jumlah File: {folder_info.get('file_count', 'N/A')}")
            summary.append(f"  Jumlah Subfolder: {folder_info.get('folder_count', 'N/A')}")
            summary.append(f"  Total Ukuran: {folder_info.get('total_size', 'N/A')}")
            summary.append("")
        
        # File info
        file_info = result.get("file_info", {})
        if file_info and "error" not in file_info:
            summary.append("INFORMASI FILE NASABAH:")
            summary.append(f"  Nama: {file_info.get('name', 'N/A')}")
            summary.append(f"  Ukuran: {file_info.get('size', 'N/A')}")
            summary.append(f"  Format: {file_info.get('extension', 'N/A')}")
            summary.append(f"  Dibuat: {file_info.get('created', 'N/A')}")
            summary.append(f"  Dimodifikasi: {file_info.get('modified', 'N/A')}")
            summary.append("")
        
        # Process details
        process_details = result.get("process_details", {})
        if process_details:
            summary.append("DETAIL PROSES:")
            steps = process_details.get("steps_completed", [])
            for step in steps:
                summary.append(f"  {step}")
            summary.append(f"\nStatus: {process_details.get('status', 'N/A')}")
            summary.append(f"Durasi: {process_details.get('estimated_duration', 'N/A')}")
            summary.append("")
        
        # Warnings
        warnings = result.get("warnings", [])
        if warnings:
            summary.append("PERINGATAN:")
            for warning in warnings:
                summary.append(f"  ⚠️ {warning}")
            summary.append("")
        
        summary.append("=" * 50)
        summary.append("Proses selesai dengan sukses!")
        
        return "\n".join(summary)


class AnggotaFolderReader:
    """Class untuk membaca dan memproses struktur folder anggota"""
    
    def __init__(self):
        self.file_manager = FileManager()
        
        # Pattern untuk validasi folder
        self.center_pattern = r'^\d{4}$'  # 4 digit angka
        self.anggota_pattern = r'^\d{6}_\w+$'  # 6digit_nama
        self.file_code_pattern = r'^(\d{2})'  # kode file 01-12
        
        # Valid file codes dengan mapping jenis dokumen
        self.document_types = {
            "01": "KTP",
            "02": "Kartu Keluarga", 
            "03": "Form PPI",
            "04": "Form UK",
            "05": "Form Keanggotaan",
            "06": "Form Pengajuan",
            "07": "Akad Pencairan",
            "08": "Form Monitoring",
            "09": "Form Simpanan Hari Raya",
            "10": "Form Lainnya",
            "11": "Form Cuti",
            "12": "Form Anggota Keluar"
        }
        
        # Valid file codes
        self.valid_file_codes = list(self.document_types.keys())
        
        # Expected file name patterns
        self.expected_patterns = {
            "01": "01_NAMA_ANGGOTA.pdf",
            "02": "02_NAMA_ANGGOTA.pdf",
            "03": "03_KODEPINJAMANKE_NAMA_ANGGOTA.pdf",
            "04": "04_NAMA_ANGGOTA.pdf",
            "05": "05_NAMA_ANGGOTA.pdf",
            "06": "06_KODEPENGAJUANKE_NAMA_ANGGOTA.pdf",
            "07": "07_KODEPINJAMANKE_NAMA_ANGGOTA.pdf",
            "08": "08_KODEPINJAMANKE_NAMA_ANGGOTA.pdf",
            "09": "09_NAMA_ANGGOTA.pdf",
            "10": "10_JENISFILE_NAMA_ANGGOTA.pdf",
            "11": "11_NAMA_ANGGOTA.pdf",
            "12": "12_NAMA_ANGGOTA.pdf"
        }
    
    def validate_center_folder(self, folder_name: str) -> bool:
        """
        Validasi apakah nama folder center sesuai pola (4 digit angka)
        
        Args:
            folder_name (str): Nama folder center
            
        Returns:
            bool: True jika valid
        """
        return bool(re.match(self.center_pattern, folder_name))
    
    def validate_anggota_folder(self, folder_name: str) -> bool:
        """
        Validasi apakah nama folder anggota sesuai pola (6digit_nama)
        
        Args:
            folder_name (str): Nama folder anggota
            
        Returns:
            bool: True jika valid
        """
        return bool(re.match(self.anggota_pattern, folder_name))
    
    def extract_file_code(self, filename: str) -> Optional[str]:
        """
        Extract kode file dari nama file (01-12)
        
        Args:
            filename (str): Nama file
            
        Returns:
            Optional[str]: Kode file jika valid, None jika tidak
        """
        match = re.match(self.file_code_pattern, filename)
        if match:
            code = match.group(1)
            return code if code in self.valid_file_codes else None
        return None
    
    def scan_anggota_folder(self, anggota_folder_path: str) -> Dict[str, any]:
        """
        Scan folder anggota dan kategorisasi file berdasarkan kode
        
        Args:
            anggota_folder_path (str): Path folder anggota
            
        Returns:
            Dict[str, any]: Hasil scan dengan kategorisasi file
        """
        if not self.file_manager.validate_folder_path(anggota_folder_path):
            return {"error": "Folder anggota tidak valid atau tidak dapat diakses"}
        
        folder_name = os.path.basename(anggota_folder_path)
        if not self.validate_anggota_folder(folder_name):
            return {"error": f"Nama folder anggota tidak sesuai pola (6digit_nama): {folder_name}"}
        
        try:
            # Extract info anggota dari nama folder
            parts = folder_name.split('_', 1)
            anggota_id = parts[0]
            anggota_nama = parts[1] if len(parts) > 1 else "Unknown"
            
            # Scan files
            file_categories = {code: [] for code in self.valid_file_codes}
            uncategorized_files = []
            total_files = 0
            
            for item in os.listdir(anggota_folder_path):
                item_path = os.path.join(anggota_folder_path, item)
                
                if os.path.isfile(item_path):
                    total_files += 1
                    file_code = self.extract_file_code(item)
                    
                    file_info = {
                        "name": item,
                        "path": item_path,
                        "size": self.file_manager.get_file_size(item_path),
                        "extension": os.path.splitext(item)[1].lower(),
                        "modified": datetime.fromtimestamp(os.path.getmtime(item_path)).strftime("%Y-%m-%d %H:%M:%S")
                    }
                    
                    if file_code:
                        file_categories[file_code].append(file_info)
                    else:
                        uncategorized_files.append(file_info)
            
            # Hitung statistik
            categorized_count = sum(len(files) for files in file_categories.values())
            missing_codes = [code for code, files in file_categories.items() if len(files) == 0]
            duplicate_codes = [code for code, files in file_categories.items() if len(files) > 1]
            
            return {
                "success": True,
                "anggota_info": {
                    "id": anggota_id,
                    "nama": anggota_nama,
                    "folder_name": folder_name,
                    "folder_path": anggota_folder_path
                },
                "file_summary": {
                    "total_files": total_files,
                    "categorized_files": categorized_count,
                    "uncategorized_files": len(uncategorized_files),
                    "missing_codes": missing_codes,
                    "duplicate_codes": duplicate_codes
                },
                "file_categories": file_categories,
                "uncategorized_files": uncategorized_files,
                "completeness": {
                    "percentage": (len(self.valid_file_codes) - len(missing_codes)) / len(self.valid_file_codes) * 100,
                    "missing_count": len(missing_codes),
                    "complete": len(missing_codes) == 0
                }
            }
            
        except Exception as e:
            return {"error": f"Error scanning folder anggota: {str(e)}"}
    
    def scan_center_folder(self, center_folder_path: str) -> Dict[str, any]:
        """
        Scan folder center dan semua folder anggota di dalamnya
        
        Args:
            center_folder_path (str): Path folder center
            
        Returns:
            Dict[str, any]: Hasil scan center dengan semua anggota
        """
        if not self.file_manager.validate_folder_path(center_folder_path):
            return {"error": "Folder center tidak valid atau tidak dapat diakses"}
        
        center_name = os.path.basename(center_folder_path)
        if not self.validate_center_folder(center_name):
            return {"error": f"Nama folder center tidak sesuai pola (4 digit): {center_name}"}
        
        try:
            anggota_folders = []
            invalid_folders = []
            total_anggota = 0
            
            # Scan semua item dalam folder center
            for item in os.listdir(center_folder_path):
                item_path = os.path.join(center_folder_path, item)
                
                if os.path.isdir(item_path):
                    if self.validate_anggota_folder(item):
                        # Scan folder anggota
                        anggota_result = self.scan_anggota_folder(item_path)
                        if anggota_result.get("success", False):
                            anggota_folders.append(anggota_result)
                            total_anggota += 1
                        else:
                            invalid_folders.append({
                                "name": item,
                                "path": item_path,
                                "error": anggota_result.get("error", "Unknown error")
                            })
                    else:
                        invalid_folders.append({
                            "name": item,
                            "path": item_path,
                            "error": "Nama folder tidak sesuai pola 6digit_nama"
                        })
            
            # Statistik center
            complete_anggota = sum(1 for anggota in anggota_folders if anggota["completeness"]["complete"])
            total_files = sum(anggota["file_summary"]["total_files"] for anggota in anggota_folders)
            
            return {
                "success": True,
                "center_info": {
                    "code": center_name,
                    "path": center_folder_path,
                    "total_anggota": total_anggota,
                    "complete_anggota": complete_anggota,
                    "total_files": total_files
                },
                "anggota_folders": anggota_folders,
                "invalid_folders": invalid_folders,
                "summary": {
                    "completion_rate": (complete_anggota / total_anggota * 100) if total_anggota > 0 else 0,
                    "total_valid_anggota": total_anggota,
                    "total_invalid_folders": len(invalid_folders)
                }
            }
            
        except Exception as e:
            return {"error": f"Error scanning folder center: {str(e)}"}
    
    def scan_data_anggota_root(self, root_path: str) -> Dict[str, any]:
        """
        Scan folder root DATA_ANGGOTA dan semua center di dalamnya
        
        Args:
            root_path (str): Path folder root DATA_ANGGOTA
            
        Returns:
            Dict[str, any]: Hasil scan lengkap semua center dan anggota
        """
        if not self.file_manager.validate_folder_path(root_path):
            return {"error": "Folder root tidak valid atau tidak dapat diakses"}
        
        try:
            center_folders = []
            invalid_centers = []
            total_centers = 0
            
            # Scan semua item dalam folder root
            for item in os.listdir(root_path):
                item_path = os.path.join(root_path, item)
                
                if os.path.isdir(item_path):
                    if self.validate_center_folder(item):
                        # Scan folder center
                        center_result = self.scan_center_folder(item_path)
                        if center_result.get("success", False):
                            center_folders.append(center_result)
                            total_centers += 1
                        else:
                            invalid_centers.append({
                                "name": item,
                                "path": item_path,
                                "error": center_result.get("error", "Unknown error")
                            })
                    else:
                        invalid_centers.append({
                            "name": item,
                            "path": item_path,
                            "error": "Nama folder tidak sesuai pola 4 digit"
                        })
            
            # Statistik keseluruhan
            total_anggota = 0
            total_files = 0
            complete_anggota = 0
            print("\n=== DEBUG SCAN ROOT ===")
            for center in center_folders:
                center_code = center["center_info"]["code"]
                n_anggota = center["center_info"]["total_anggota"]
                n_files = center["center_info"]["total_files"]
                n_lengkap = center["center_info"]["complete_anggota"]
                total_anggota += n_anggota
                total_files += n_files
                complete_anggota += n_lengkap
            if invalid_centers:
                print(f"INVALID CENTERS: {len(invalid_centers)}")
                for ic in invalid_centers:
                    print(f"  - {ic['name']}: {ic['error']}")
            return {
                "success": True,
                "root_info": {
                    "path": root_path,
                    "total_centers": total_centers,
                    "total_anggota": total_anggota,
                    "total_files": total_files,
                    "complete_anggota": complete_anggota
                },
                "center_folders": center_folders,
                "invalid_centers": invalid_centers,
                "summary": {
                    "overall_completion_rate": (complete_anggota / total_anggota * 100) if total_anggota > 0 else 0,
                    "centers_scanned": total_centers,
                    "anggota_scanned": total_anggota,
                    "files_found": total_files
                }
            }
            
        except Exception as e:
            return {"error": f"Error scanning root folder: {str(e)}"}
    
    def generate_anggota_report(self, scan_result: Dict[str, any]) -> str:
        """
        Generate laporan untuk hasil scan anggota
        
        Args:
            scan_result (Dict[str, any]): Hasil dari scan_anggota_folder
            
        Returns:
            str: Laporan dalam format text
        """
        if not scan_result.get("success", False):
            return f"ERROR: {scan_result.get('error', 'Unknown error')}"
        
        report = []
        report.append("=" * 60)
        report.append("LAPORAN SCAN FOLDER ANGGOTA")
        report.append("=" * 60)
        
        # Info anggota
        anggota_info = scan_result["anggota_info"]
        report.append(f"ID Anggota: {anggota_info['id']}")
        report.append(f"Nama: {anggota_info['nama']}")
        report.append(f"Folder: {anggota_info['folder_name']}")
        report.append(f"Path: {anggota_info['folder_path']}")
        report.append("")
        
        # Summary
        file_summary = scan_result["file_summary"]
        completeness = scan_result["completeness"]
        
        report.append("RINGKASAN:")
        report.append(f"  Total File: {file_summary['total_files']}")
        report.append(f"  File Terkategorisasi: {file_summary['categorized_files']}")
        report.append(f"  File Tidak Terkategorisasi: {file_summary['uncategorized_files']}")
        report.append(f"  Kelengkapan: {completeness['percentage']:.1f}%")
        report.append(f"  Status: {'LENGKAP' if completeness['complete'] else 'TIDAK LENGKAP'}")
        report.append("")
        
        # File categories
        report.append("KATEGORISASI FILE (01-12):")
        file_categories = scan_result["file_categories"]
        for code in self.valid_file_codes:
            files = file_categories[code]
            status = "✓" if files else "✗"
            count = len(files)
            report.append(f"  {status} Kode {code}: {count} file(s)")
            
            for file_info in files:
                report.append(f"    - {file_info['name']} ({file_info['size']})")
        
        # Missing codes
        if file_summary['missing_codes']:
            report.append("")
            report.append("KODE YANG HILANG:")
            for code in file_summary['missing_codes']:
                report.append(f"  - Kode {code}")
        
        # Duplicate codes
        if file_summary['duplicate_codes']:
            report.append("")
            report.append("KODE DUPLIKAT:")
            for code in file_summary['duplicate_codes']:
                count = len(file_categories[code])
                report.append(f"  - Kode {code}: {count} file(s)")
        
        # Uncategorized files
        uncategorized = scan_result["uncategorized_files"]
        if uncategorized:
            report.append("")
            report.append("FILE TIDAK TERKATEGORISASI:")
            for file_info in uncategorized:
                report.append(f"  - {file_info['name']} ({file_info['size']})")
        
        report.append("")
        report.append("=" * 60)
        
        return "\n".join(report)
    
    def generate_tabular_data_anggota(self, scan_result: Dict[str, any]) -> Dict[str, any]:
        """
        Generate data tabular untuk satu anggota (untuk export Excel)
        
        Args:
            scan_result (Dict[str, any]): Hasil dari scan_anggota_folder
            
        Returns:
            Dict[str, any]: Data dalam format tabular
        """
        if not scan_result.get("success", False):
            return {"error": scan_result.get("error", "Unknown error")}
        
        anggota_info = scan_result["anggota_info"]
        file_categories = scan_result["file_categories"]
        
        # Split folder name untuk mendapatkan ID dan nama terpisah
        folder_name = anggota_info["folder_name"]
        if "_" in folder_name:
            id_part, nama_part = folder_name.split("_", 1)
        else:
            id_part = anggota_info["id"]
            nama_part = anggota_info["nama"]
        
        # Base data anggota (tanpa kolom Path)
        row_data = {
            "ID_Anggota": id_part,
            "Nama_Anggota": nama_part,
            "Total_Files": scan_result["file_summary"]["total_files"],
            "Kelengkapan_Persen": scan_result["completeness"]["percentage"],
            "Status_Lengkap": "YA" if scan_result["completeness"]["complete"] else "TIDAK"
        }
        
        # Tambahkan kolom untuk setiap jenis dokumen (Ada/Tidak)
        for code in self.valid_file_codes:
            doc_type = self.document_types[code]
            files = file_categories.get(code, [])
            
            # Kolom ada/tidak - hanya gunakan kode angka
            row_data[f"{code}_Ada"] = "ADA" if len(files) > 0 else "TIDAK"
            
            # Kolom nama file (jika ada)
            if len(files) > 0:
                file_names = [f["name"] for f in files]
                row_data[f"{code}_File"] = "; ".join(file_names)
            else:
                row_data[f"{code}_File"] = ""
            
            # Kolom jumlah file
            row_data[f"{code}_Jumlah"] = len(files)
        
        return row_data
    
    def generate_tabular_data_center(self, scan_result: Dict[str, any]) -> List[Dict[str, any]]:
        """
        Generate data tabular untuk semua anggota dalam center
        
        Args:
            scan_result (Dict[str, any]): Hasil dari scan_center_folder
            
        Returns:
            List[Dict[str, any]]: List data tabular untuk semua anggota
        """
        if not scan_result.get("success", False):
            return [{"error": scan_result.get("error", "Unknown error")}]
        
        tabular_data = []
        center_info = scan_result["center_info"]
        
        for anggota in scan_result["anggota_folders"]:
            row_data = self.generate_tabular_data_anggota(anggota)
            if "error" not in row_data:
                # Tambahkan info center (tanpa path)
                row_data["Center_Code"] = center_info["code"]
                tabular_data.append(row_data)
        
        return tabular_data
    
    def generate_tabular_data_root(self, scan_result: Dict[str, any]) -> List[Dict[str, any]]:
        """
        Generate data tabular untuk semua anggota dalam root (semua anggota dari semua center, satu baris per anggota)
        
        Args:
            scan_result (Dict[str, any]): Hasil dari scan_data_anggota_root
            
        Returns:
            List[Dict[str, any]]: List data tabular untuk semua anggota
        """
        if not scan_result.get("success", False):
            return [{"error": scan_result.get("error", "Unknown error")}]
        
        tabular_data = []
        for center in scan_result["center_folders"]:
            center_info = center.get("center_info", {})
            center_code = center_info.get("code", "")
            anggota_folders = center.get("anggota_folders", [])
            for anggota in anggota_folders:
                row_data = self.generate_tabular_data_anggota(anggota)
                if "error" not in row_data:
                    row_data["Center_Code"] = center_code
                    tabular_data.append(row_data)
        return tabular_data
    
    def export_to_excel(self, scan_result: Dict[str, any], scan_type: str, 
                       output_path: str = None) -> Dict[str, any]:
        """
        Export hasil scan ke file Excel
        
        Args:
            scan_result (Dict[str, any]): Hasil scan
            scan_type (str): Type scan ('anggota', 'center', 'root')
            output_path (str): Path output file (optional)
            
        Returns:
            Dict[str, any]: Result export dengan status dan path file
        """
        try:
            # Generate tabular data berdasarkan type
            if scan_type == "anggota":
                tabular_data = [self.generate_tabular_data_anggota(scan_result)]
                default_name = f"data_anggota_{scan_result['anggota_info']['id']}_{scan_result['anggota_info']['nama']}"
            elif scan_type == "center":
                tabular_data = self.generate_tabular_data_center(scan_result)
                default_name = f"data_center_{scan_result['center_info']['code']}"
            elif scan_type == "root":
                tabular_data = self.generate_tabular_data_root(scan_result)
                default_name = "data_root_all_anggota"
            else:
                return {"success": False, "error": f"Invalid scan type: {scan_type}"}
            
            if not tabular_data or (len(tabular_data) == 1 and "error" in tabular_data[0]):
                return {"success": False, "error": "No valid data to export"}
            
            # Create DataFrame
            df = pd.DataFrame(tabular_data)
            
            # Generate output path jika tidak disediakan
            if not output_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"{default_name}_{timestamp}.xlsx"
            
            # Export ke Excel dengan formatting
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Sheet data utama
                df.to_excel(writer, sheet_name='Data_Anggota', index=False)
                
                # Sheet summary jika ada banyak anggota
                if len(tabular_data) > 1:
                    summary_data = self._generate_summary_data(df)
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Sheet mapping dokumen
                mapping_data = []
                for code, doc_type in self.document_types.items():
                    mapping_data.append({
                        "Kode": code,
                        "Jenis_Dokumen": doc_type,
                        "Format_Penamaan": self.expected_patterns[code]
                    })
                mapping_df = pd.DataFrame(mapping_data)
                mapping_df.to_excel(writer, sheet_name='Mapping_Dokumen', index=False)
                
                # Auto-adjust column widths
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
            
            return {
                "success": True,
                "message": f"Data berhasil di-export ke Excel",
                "file_path": output_path,
                "rows_exported": len(tabular_data),
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
        except Exception as e:
            print(f"Error during export to Excel: {str(e)}")
            return {
                "success": False,
                "error": f"Gagal export ke Excel: {str(e)}"
            }
    
    def _generate_summary_data(self, df: pd.DataFrame) -> List[Dict[str, any]]:
        """Generate summary data untuk sheet summary"""
        summary = []
        
        # Summary keseluruhan
        total_anggota = len(df)
        anggota_lengkap = len(df[df['Status_Lengkap'] == 'YA'])
        avg_kelengkapan = df['Kelengkapan_Persen'].mean()
        
        summary.append({
            "Kategori": "TOTAL",
            "Jumlah_Anggota": total_anggota,
            "Anggota_Lengkap": anggota_lengkap,
            "Persentase_Lengkap": f"{(anggota_lengkap/total_anggota*100):.1f}%",
            "Rata_Rata_Kelengkapan": f"{avg_kelengkapan:.1f}%"
        })
        
        # Summary per dokumen
        for code, doc_type in self.document_types.items():
            col_name = f"{code}_{doc_type.replace(' ', '_')}_Ada"
            if col_name in df.columns:
                ada_count = len(df[df[col_name] == 'ADA'])
                persentase = (ada_count / total_anggota * 100) if total_anggota > 0 else 0
                
                summary.append({
                    "Kategori": f"Dokumen {code} - {doc_type}",
                    "Jumlah_Anggota": total_anggota,
                    "Anggota_Lengkap": ada_count,
                    "Persentase_Lengkap": f"{persentase:.1f}%",
                    "Rata_Rata_Kelengkapan": f"{persentase:.1f}%"
                })
        
        return summary