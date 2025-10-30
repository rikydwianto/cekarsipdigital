# 📁 AppData Storage Migration (v1.1.7)

## 🎯 Overview

Version 1.1.7 memindahkan semua file konfigurasi dan database ke **AppData Local** untuk mengikuti Windows best practices.

## 📦 Files Moved to AppData

### Before (v1.1.6)
```
D:\PROJECT\PYTHON\ARSIPOWNCLOUD\
├── app_config.json                    ❌ Mixed with source code
├── universal_scan_database.xlsx       ❌ In project root
├── database.xlsx                      ✅ Already in AppData
└── file_export.xlsx                   ✅ Already in AppData
```

### After (v1.1.7)
```
C:\Users\[Username]\AppData\Local\ArsipDigitalOwnCloud\
├── app_config.json                    ✅ Configuration
├── universal_scan_database.xlsx       ✅ Universal scan data
├── database.xlsx                      ✅ Arsip scan data
└── file_export.xlsx                   ✅ Export data
```

## 🔧 Technical Changes

### 1. **app_helpers.py**

Added helper functions:
```python
def get_config_path():
    """Get full path untuk app_config.json di AppData"""
    return os.path.join(get_appdata_path(), 'app_config.json')

def get_universal_scan_database_path():
    """Get full path untuk universal_scan_database.xlsx di AppData"""
    return os.path.join(get_appdata_path(), 'universal_scan_database.xlsx')
```

Updated ConfigManager:
```python
class ConfigManager:
    def __init__(self):
        self.config_file = get_config_path()  # Uses AppData path
        # ... rest of the code
```

### 2. **app_arsip.py**

UniversalScanApp now uses AppData:
```python
def __init__(self, root, parent_window=None):
    # ...
    from app_helpers import get_universal_scan_database_path
    self.database_file = get_universal_scan_database_path()
    # ...
```

### 3. **web_server.py**

WebServerManager uses AppData config:
```python
def __init__(self, config_file=None):
    # ...
    if config_file is None:
        from app_helpers import get_config_path
        self.config_file = get_config_path()
    # ...
```

## ✨ Benefits

### 🔒 Security & Best Practices
- ✅ **No Admin Required**: User can write without elevated privileges
- ✅ **Program Files Protection**: No write to Program Files folder
- ✅ **Windows Standards**: Follows Microsoft guidelines

### 👥 Multi-User Support
- ✅ **User Isolation**: Each Windows user has separate data
- ✅ **Profile-Based**: Data follows user profile
- ✅ **No Conflicts**: Multiple users on same PC don't interfere

### 💾 Data Management
- ✅ **Automatic Backup**: Included in Windows Backup
- ✅ **Easy Location**: Standard AppData location
- ✅ **Clean Uninstall**: Remove app without leaving data in Program Files

### 📁 Project Organization
- ✅ **Clean Source**: No data files mixed with code
- ✅ **Git Friendly**: No need to gitignore data files in project
- ✅ **Portable Build**: Executable doesn't carry user data

## 🔄 Migration Path

### Automatic Migration
The application will:
1. Check if files exist in old location (project root)
2. Auto-create new files in AppData if needed
3. Users can manually move existing data if desired

### Manual Migration (Optional)

If you have existing data in project root:

1. **Close the application**

2. **Copy files to AppData:**
   ```powershell
   # Open PowerShell
   $appdata = "$env:LOCALAPPDATA\ArsipDigitalOwnCloud"
   
   # Copy config
   Copy-Item "D:\PROJECT\PYTHON\ARSIPOWNCLOUD\app_config.json" $appdata
   
   # Copy database
   Copy-Item "D:\PROJECT\PYTHON\ARSIPOWNCLOUD\universal_scan_database.xlsx" $appdata
   ```

3. **Delete old files from project root:**
   ```powershell
   Remove-Item "D:\PROJECT\PYTHON\ARSIPOWNCLOUD\app_config.json"
   Remove-Item "D:\PROJECT\PYTHON\ARSIPOWNCLOUD\universal_scan_database.xlsx"
   ```

4. **Restart the application**

## 📍 Access AppData Folder

### Method 1: Keyboard Shortcut
1. Press `Win + R`
2. Type: `%LOCALAPPDATA%\ArsipDigitalOwnCloud`
3. Press Enter

### Method 2: File Explorer
1. Open File Explorer
2. Navigate to: `C:\Users\[YourUsername]\AppData\Local\ArsipDigitalOwnCloud`
3. (Enable "Show hidden files" if needed)

### Method 3: PowerShell
```powershell
explorer "$env:LOCALAPPDATA\ArsipDigitalOwnCloud"
```

## 🧪 Testing

Tested scenarios:
- ✅ Fresh install (no existing data)
- ✅ Config creation in AppData
- ✅ Database creation in AppData
- ✅ Settings persistence
- ✅ Universal scan database
- ✅ Multi-user isolation
- ✅ Backup/restore compatibility

## 📊 File Locations Reference

| File                             | Location                                                    |
| -------------------------------- | ----------------------------------------------------------- |
| `app_config.json`                | `%LOCALAPPDATA%\ArsipDigitalOwnCloud\app_config.json`       |
| `database.xlsx`                  | `%LOCALAPPDATA%\ArsipDigitalOwnCloud\database.xlsx`         |
| `file_export.xlsx`               | `%LOCALAPPDATA%\ArsipDigitalOwnCloud\file_export.xlsx`      |
| `universal_scan_database.xlsx`   | `%LOCALAPPDATA%\ArsipDigitalOwnCloud\universal_scan_database.xlsx` |

## 🔍 Troubleshooting

### Issue: Cannot find AppData folder
**Solution**: 
- Folder is hidden by default
- Enable "Show hidden files" in File Explorer
- Or use `Win + R` → `%LOCALAPPDATA%`

### Issue: Permission denied
**Solution**:
- AppData should always be writable by current user
- Check if antivirus is blocking
- Run as administrator only if absolutely necessary

### Issue: Data not persisting
**Solution**:
- Check if folder exists: `%LOCALAPPDATA%\ArsipDigitalOwnCloud`
- Verify write permissions
- Check disk space

### Issue: Multiple users seeing same data
**Solution**:
- This should not happen with AppData
- Each user has separate `C:\Users\[Username]\AppData`
- Verify you're not using network/roaming profiles incorrectly

## 🚀 Future Considerations

### Cloud Sync (Future)
- Consider OneDrive/Dropbox integration
- Add export/import functionality
- Implement data synchronization

### Backup Strategy
- Automatic backup to secondary location
- Export configuration as JSON
- Database backup scheduling

### Performance
- Monitor AppData folder size
- Implement cleanup routines
- Add compression for large databases

## 📝 Notes

- Old files in project root are not automatically deleted
- Users can keep backup copies if desired
- Application will always use AppData location going forward
- No breaking changes - all features work as before

## 🎯 Version Info

- **Version**: 1.1.7
- **Release Date**: October 30, 2025
- **Previous Version**: 1.1.6 (Major Refactoring)
- **Migration Type**: Non-breaking (automatic)

---

**Updated**: October 30, 2025  
**Author**: Riky Dwianto
