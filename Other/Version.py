from pathlib import Path
import sys


def _app_root() -> Path:
    # Frozen (PyInstaller) builds keep resources next to the executable
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


class Version(object):
    # Major.Minor.Patch.Build, kept in a plain text file - the single source of the app version
    _version_file = _app_root() / "resources" / "version.txt"
    __instance = None

    def __init__(self):
        if self.__initialized:
            return
        self.__initialized = True
        if self._version_file.exists():
            self.version = [int(x) for x in self._version_file.read_text(encoding="utf8").strip().split(".")]
        else:
            self.version = [0, 0, 1, 0]

    def __new__(cls):
        if cls.__instance is None:
            cls.__instance = super(Version, cls).__new__(cls)
            cls.__instance.__initialized = False
        return cls.__instance

    def get_version(self):
        return self.version

    def __repr__(self):
        return ".".join(str(x) for x in self.version)

    def __str__(self):
        return self.__repr__()

    def up_build(self):
        self.version[3] += 1
        return self.get_version()

    def up_patch(self):
        major_val, minor_val, patch_val, _ = self.version
        if patch_val + 1 >= 10:
            patch_val = 0
            minor_val += 1
        else:
            patch_val += 1

        self.version = [major_val, minor_val, patch_val, 0]
        return self.get_version()

    def up_major(self):
        self.version[0] += 1
        return self.get_version()

    def set_version(self, major, minor, patch, build):
        self.version = [major, minor, patch, build]

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.save()

    def save(self):
        self._version_file.write_text(str(self) + "\n", encoding="utf8")


if __name__ == "__main__":
    ver = Version()
    ver.up_patch()
    ver.save()
