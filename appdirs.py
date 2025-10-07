#!/usr/bin/env python

import os
import sys

from platformdirs import site_data_dir, user_cache_dir, user_data_dir, user_log_dir

PY3 = sys.version_info[0] == 3

if PY3:
    unicode = str


class EnvAppDirs(object):
    def __init__(self, appname, appauthor, root_path):
        self.appname = appname
        self.appauthor = appauthor
        self.root_path = root_path

    @property
    def user_data_dir(self):
        return os.path.join(self.root_path, "data")

    @property
    def site_data_dir(self):
        return os.path.join(self.root_path, "data")

    @property
    def user_cache_dir(self):
        return os.path.join(self.root_path, "cache")

    @property
    def user_log_dir(self):
        return os.path.join(self.root_path, "log")


class AppDirs(object):
    """Convenience wrapper for getting application dirs."""

    def __init__(self, appname, appauthor, version=None, roaming=False):
        self.appname = appname
        self.appauthor = appauthor
        self.version = version
        self.roaming = roaming

    @property
    def user_data_dir(self):
        return user_data_dir(
            self.appname, self.appauthor, version=self.version, roaming=self.roaming
        )

    @property
    def site_data_dir(self):
        return site_data_dir(self.appname, self.appauthor, version=self.version)

    @property
    def user_cache_dir(self):
        return user_cache_dir(self.appname, self.appauthor, version=self.version)

    @property
    def user_log_dir(self):
        return user_log_dir(self.appname, self.appauthor, version=self.version)


def _get_win_folder_from_registry(csidl_name):
    """This is a fallback technique at best. I'm not sure if using the
    registry for this guarantees us the correct answer for all CSIDL_*
    names.
    """
    import _winreg

    shell_folder_name = {
        "CSIDL_APPDATA": "AppData",
        "CSIDL_COMMON_APPDATA": "Common AppData",
        "CSIDL_LOCAL_APPDATA": "Local AppData",
    }[csidl_name]

    key = _winreg.OpenKey(
        _winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
    )
    dir, type = _winreg.QueryValueEx(key, shell_folder_name)
    return dir


def _get_win_folder_with_pywin32(csidl_name):
    from win32com.shell import shell, shellcon

    dir = shell.SHGetFolderPath(0, getattr(shellcon, csidl_name), 0, 0)
    # Try to make this a unicode path because SHGetFolderPath does
    # not return unicode strings when there is unicode data in the
    # path.
    try:
        dir = unicode(dir)

        # Downgrade to short path name if have highbit chars. See
        # <http://bugs.activestate.com/show_bug.cgi?id=85099>.
        has_high_char = False
        for c in dir:
            if ord(c) > 255:
                has_high_char = True
                break
        if has_high_char:
            try:
                import win32api

                dir = win32api.GetShortPathName(dir)
            except ImportError:
                pass
    except UnicodeError:
        pass
    return dir


def _get_win_folder_with_ctypes(csidl_name):
    import ctypes

    csidl_const = {
        "CSIDL_APPDATA": 26,
        "CSIDL_COMMON_APPDATA": 35,
        "CSIDL_LOCAL_APPDATA": 28,
    }[csidl_name]

    buf = ctypes.create_unicode_buffer(1024)
    ctypes.windll.shell32.SHGetFolderPathW(None, csidl_const, None, 0, buf)

    # Downgrade to short path name if have highbit chars. See
    # <http://bugs.activestate.com/show_bug.cgi?id=85099>.
    has_high_char = False
    for c in buf:
        if ord(c) > 255:
            has_high_char = True
            break
    if has_high_char:
        buf2 = ctypes.create_unicode_buffer(1024)
        if ctypes.windll.kernel32.GetShortPathNameW(buf.value, buf2, 1024):
            buf = buf2

    return buf.value


if sys.platform == "win32":
    try:
        import win32com.shell

        _get_win_folder = _get_win_folder_with_pywin32
    except ImportError:
        try:
            import ctypes

            _get_win_folder = _get_win_folder_with_ctypes
        except ImportError:
            _get_win_folder = _get_win_folder_from_registry


# ---- self test code

if __name__ == "__main__":
    appname = "MyApp"
    appauthor = "MyCompany"

    props = ("user_data_dir", "site_data_dir", "user_cache_dir", "user_log_dir")

    print("-- app dirs (without optional 'version')")
    dirs = AppDirs(appname, appauthor, version="1.0")
    for prop in props:
        print("%s: %s" % (prop, getattr(dirs, prop)))

    print("\n-- app dirs (with optional 'version')")
    dirs = AppDirs(appname, appauthor)
    for prop in props:
        print("%s: %s" % (prop, getattr(dirs, prop)))
