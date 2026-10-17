from pathlib import Path
import sys
import numpy as np


def _app_root() -> Path:
    # Frozen (PyInstaller) builds keep resources next to the executable
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


class Version(object):
    # Major.Minor.Patch.Build
    _version_file = _app_root() / "resources" / "version.npy"
    __instance = None

    def __init__(self):
        if self.__initialized:
            return
        self.__initialized = True
        if self._version_file.exists():
            with self._version_file.open(mode='rb') as np_file:
                self.version_np = np.load(np_file)
        else:
            self.version_np = np.zeros(4, dtype=int)
            self.version_np[2] = 1

    def __new__(cls):
        if cls.__instance is None:
            cls.__instance = super(Version, cls).__new__(cls)
            cls.__instance.__initialized = False
        return cls.__instance

    def get_version(self):
        return self.version_np

    def __repr__(self):
        return f"{self.version_np[0]}.{self.version_np[1]}.{self.version_np[2]}.{self.version_np[3]}"

    def __str__(self):
        return f"{self.version_np[0]}.{self.version_np[1]}.{self.version_np[2]}.{self.version_np[3]}"

    def up_build(self):
        self.version_np[3] += 1
        return self.get_version()

    def up_patch(self):
        major_val = self.version_np[0]
        minor_val = self.version_np[1]
        patch_val = self.version_np[2]
        if patch_val + 1 >= 10:
            patch_val = 0
            minor_val += 1
        else:
            patch_val += 1

        self.version_np[0] = major_val
        self.version_np[1] = minor_val
        self.version_np[2] = patch_val
        self.version_np[3] = 0
        return self.get_version()

    def up_major(self):
        self.version_np[0] += 1
        self.version_np[1] += 0
        self.version_np[2] += 0
        self.version_np[3] += 0
        return self.get_version()

    def set_version(self, major, minor, patch, build):
        self.version_np = np.array([major, minor, patch, build])

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.save()

    def save(self):
        with self._version_file.open(mode='wb') as output_file:
            np.save(output_file, self.version_np)


if __name__ == "__main__":
    ver = Version()
    # ver.set_version(0, 0, 1, 0)
    ver.up_patch()
    ver.save()
