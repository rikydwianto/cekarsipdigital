# 🔄 Migration Note - Tool Komida v1.1.7

## ⚠️ IMPORTANT: AppData Folder Changed

Starting from v1.1.7, the application has been rebranded to **Tool Komida** and the AppData folder location has changed.

---

## 📁 Folder Location Changes

### Old Location (v1.1.6 and earlier):
```
C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\
```

### New Location (v1.1.7+):
```
C:\Users\[Username]\AppData\Local\ToolKomida\
```

---

## 🔄 How to Migrate Your Data

If you have existing data in the old location, follow these steps:

### Option 1: Manual Copy (Recommended)

1. **Close the application** if it's running

2. **Open PowerShell** and run:
   ```powershell
   # Check if old data exists
   $oldPath = "$env:LOCALAPPDATA\ArsipDigitalOwnCloud"
   $newPath = "$env:LOCALAPPDATA\ToolKomida"
   
   if (Test-Path $oldPath) {
       # Create new folder
       New-Item -ItemType Directory -Path $newPath -Force
       
       # Copy all files
       Copy-Item "$oldPath\*" $newPath -Recurse -Force
       
       Write-Host "✅ Data migrated successfully!" -ForegroundColor Green
       Write-Host "Old location: $oldPath" -ForegroundColor Yellow
       Write-Host "New location: $newPath" -ForegroundColor Green
   } else {
       Write-Host "❌ No old data found to migrate" -ForegroundColor Red
   }
   ```

3. **Verify the migration:**
   ```powershell
   # List files in new location
   Get-ChildItem "$env:LOCALAPPDATA\ToolKomida"
   ```

4. **Start the application** - it will now use the new location

5. **Optional: Delete old folder** (only after verifying everything works):
   ```powershell
   Remove-Item "$env:LOCALAPPDATA\ArsipDigitalOwnCloud" -Recurse -Force
   ```

### Option 2: Fresh Start

If you prefer to start fresh:

1. Simply run the application
2. It will create a new folder at the new location
3. Reconfigure your settings (default folder, etc.)
4. Your old data remains in the old location as backup

---

## 📋 Files to Migrate

These are the files you should copy:

| File                             | Description                    |
| -------------------------------- | ------------------------------ |
| `app_config.json`                | Application configuration      |
| `database.xlsx`                  | Arsip scan database            |
| `file_export.xlsx`               | Export data                    |
| `universal_scan_database.xlsx`   | Universal scan database        |

---

## 🧪 Verification Steps

After migration, verify everything works:

1. ✅ Open the application
2. ✅ Check Settings → Default Folder is preserved
3. ✅ Open "Cek Arsip Digital" → Database should be loaded
4. ✅ Check "Universal Scan" → Database should exist
5. ✅ All your previous configurations should work

---

## 🔍 Quick Access to New Location

### Method 1: Run Dialog
```
Win + R → %LOCALAPPDATA%\ToolKomida → Enter
```

### Method 2: PowerShell
```powershell
explorer "$env:LOCALAPPDATA\ToolKomida"
```

### Method 3: File Explorer
```
C:\Users\[YourUsername]\AppData\Local\ToolKomida
```

---

## 🆘 Troubleshooting

### Issue: Cannot find old data
**Solution**: Check if you previously used the application. New users don't need migration.

### Issue: Data not showing after migration
**Solution**: 
- Verify files exist in new location
- Check file permissions
- Restart the application

### Issue: Application creates new empty database
**Solution**: 
- Files may not have been copied correctly
- Check the new folder path
- Repeat migration steps

### Issue: Settings not preserved
**Solution**:
- Make sure `app_config.json` was copied
- Reconfigure settings manually if needed

---

## 📝 Notes

- The old folder will NOT be automatically deleted
- You can keep it as a backup
- Delete it manually only after verifying everything works
- Both folders can coexist without issues

---

## 🎯 What's New in v1.1.7

Besides the rebranding:
- ✅ Clearer application name: **Tool Komida**
- ✅ Professional subtitle: "Sistem Manajemen Arsip Digital & Tools"
- ✅ Updated window titles to reflect new branding
- ✅ AppData folder renamed for consistency

---

**Version**: 1.1.7  
**Release Date**: October 30, 2025  
**Migration Required**: Optional (for existing users)

---

Need help? Contact: Riky Dwianto
