"""
Script untuk convert image (PNG/JPG) ke ICO untuk icon aplikasi
"""
from PIL import Image
import os

def convert_to_ico(input_file, output_file="icon.ico", sizes=[(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)]):
    """
    Convert image ke ICO dengan multiple sizes
    
    Args:
        input_file: Path ke file gambar (PNG, JPG, dll)
        output_file: Path output file ICO (default: icon.ico)
        sizes: List of sizes untuk ICO
    """
    if not os.path.exists(input_file):
        print(f"❌ File {input_file} tidak ditemukan!")
        return False
    
    try:
        # Buka gambar
        img = Image.open(input_file)
        print(f"📂 Membuka: {input_file}")
        print(f"   Size: {img.size}")
        print(f"   Mode: {img.mode}")
        
        # Convert ke RGBA jika bukan
        if img.mode != 'RGBA':
            img = img.convert('RGBA')
            print("   Converted to RGBA")
        
        # Buat list images dengan berbagai size
        icon_sizes = []
        for size in sizes:
            # Resize dengan antialiasing
            resized = img.resize(size, Image.Resampling.LANCZOS)
            icon_sizes.append(resized)
            print(f"   ✅ Created {size[0]}x{size[1]}")
        
        # Save sebagai ICO
        icon_sizes[0].save(
            output_file,
            format='ICO',
            sizes=[(img.width, img.height) for img in icon_sizes],
            append_images=icon_sizes[1:]
        )
        
        print(f"\n✅ Icon berhasil dibuat: {output_file}")
        print(f"   Sizes: {', '.join([f'{s[0]}x{s[1]}' for s in sizes])}")
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False


if __name__ == "__main__":
    print("🎨 Icon Converter untuk Aplikasi Arsip Digital\n")
    print("=" * 50)
    
    # Input file
    input_file = input("📁 Masukkan path file gambar (PNG/JPG): ").strip()
    
    if not input_file:
        print("\n❌ Path tidak boleh kosong!")
        input("\nTekan Enter untuk keluar...")
        exit()
    
    # Remove quotes jika ada
    input_file = input_file.strip('"').strip("'")
    
    # Output file (default: icon.ico)
    output_file = input("💾 Nama file output (default: icon.ico): ").strip() or "icon.ico"
    
    if not output_file.endswith('.ico'):
        output_file += '.ico'
    
    print(f"\n🔄 Converting...")
    print(f"   Input:  {input_file}")
    print(f"   Output: {output_file}\n")
    
    # Convert
    if convert_to_ico(input_file, output_file):
        print(f"\n✅ Selesai! Icon siap digunakan.")
        print(f"\n📝 Cara pakai:")
        print(f"   1. Copy {output_file} ke folder project")
        print(f"   2. Pastikan code di main.py ada: root.iconbitmap('{output_file}')")
        print(f"   3. Untuk EXE, tambahkan icon di setup.py")
    else:
        print(f"\n❌ Gagal convert icon!")
    
    input("\nTekan Enter untuk keluar...")
