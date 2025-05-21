#utils.py
import os
import datetime
import win32security
import platform
from datetime import timedelta


def get_file_times(file_path):
    stat_info = os.stat(file_path)
    modified = datetime.datetime.fromtimestamp(stat_info.st_mtime)
    accessed = datetime.datetime.fromtimestamp(stat_info.st_atime)
    
    # Windows: 创建时间来自 st_ctime（创建时间）
    # Linux: st_ctime 是 inode 变化时间，没有创建时间
    if platform.system().lower() == "windows":
        created = datetime.datetime.fromtimestamp(stat_info.st_ctime)
    else:
        # Linux 无法直接获取创建时间，使用最早的时间代替
        created = min(modified, accessed)

    return {
        "created": created,
        "modified": modified,
        "accessed": accessed
    }



def get_file_owner(file_path):
    if platform.system().lower() == "windows":
        try:
            # Windows
            sd = win32security.GetFileSecurity(file_path, win32security.OWNER_SECURITY_INFORMATION)
            owner_sid = sd.GetSecurityDescriptorOwner()
            name, domain, _ = win32security.LookupAccountSid(None, owner_sid)
            return f"{domain}\\{name}"
        except Exception as e:
            return f"获取失败: {e}"
    else:
        # Linux/macOS
        import pwd
        try:
            return pwd.getpwuid(os.stat(file_path).st_uid).pw_name
        except Exception as e:
            return f"获取失败: {e}"

def convert_ole_time(ole_time):
    try:
        if ole_time is None:
            return ""
        return datetime.datetime(1601, 1, 1) + timedelta(microseconds=ole_time / 10)
    except:
        return ""
