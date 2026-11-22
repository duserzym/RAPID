"""Isolated settings for the packaged software checks."""
from contextlib import contextmanager
import os
from tempfile import TemporaryDirectory

from rapidpy_common.paths import absolute_io_path


@contextmanager
def offline_profile(root):
    from PySide6 import QtCore
    from rapid_main.config import AppConfig

    environment = {'RAPID_CONFIG': str(root / 'config.json'),
                   'RAPID_SAFETY_STATE': str(root / 'hardware-safety.json')}
    previous = {key: os.environ.get(key) for key in environment}
    settings = QtCore.QSettings
    # Replace only the named user settings used by RAPID. Constructing the
    # temporary QSettings object below never reads or writes the user's file.
    class IsolatedSettings(settings):
        def __init__(self, *args, **kwargs):
            if len(args) == 2 and all(isinstance(arg, str) for arg in args):
                organization, application = args
                super().__init__(str(root / organization / f'{application}.ini'), settings.IniFormat, **kwargs)
            else:
                super().__init__(*args, **kwargs)

    os.environ.update(environment)
    QtCore.QSettings = IsolatedSettings
    try:
        with TemporaryDirectory(prefix='data-', dir=absolute_io_path(root)) as data:
            config = AppConfig()
            config.general.nocomm = True
            config.general.data_dir = data
            config.save()
            yield config
    finally:
        QtCore.QSettings = settings
        for key, value in previous.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
