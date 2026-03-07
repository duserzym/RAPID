import sys

from vrm_logger.app import main


MIN_PYTHON = (3, 10)

if sys.version_info < MIN_PYTHON:
    sys.exit(
        "This application requires Python 3.10+ (for dataclasses with slots). "
        "Please run it using the 'paleomag' conda environment or upgrade your Python.\n"
        "Example: `conda activate paleomag && python RapidPy/vrm_logger/main.py`"
    )


if __name__ == "__main__":
    raise SystemExit(main())
