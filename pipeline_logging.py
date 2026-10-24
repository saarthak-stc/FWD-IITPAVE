import logging
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
LOG_PATH = SCRIPT_DIR / "pipeline.log"
_CONFIGURED = False


def _log_uncaught_exception(exception_type, exception, traceback):
    if issubclass(exception_type, KeyboardInterrupt):
        sys.__excepthook__(exception_type, exception, traceback)
        return
    logging.getLogger("pipeline").critical(
        "Uncaught exception", exc_info=(exception_type, exception, traceback)
    )
    sys.__excepthook__(exception_type, exception, traceback)


def get_logger(script_file):
    global _CONFIGURED
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    if not _CONFIGURED:
        formatter = logging.Formatter(
            "%(asctime)s %(levelname)s [%(name)s] %(message)s"
        )
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
        root_logger.addHandler(console_handler)
        _CONFIGURED = True

    logger = logging.getLogger(Path(script_file).stem)
    logger.setLevel(logging.INFO)
    logger.propagate = True

    sys.excepthook = _log_uncaught_exception
    return logger


def find_iitpave_directory(script_dir):
    script_dir = Path(script_dir).resolve()
    candidates = (
        script_dir
        / "IRC 37 2018-20261017T090045Z-1-001"
        / "IRC 37 2018"
        / "IITPAVE",
        script_dir
        / "IRC 37 2018-20261017T090045Z-1-001"
        / "IRC 37 2018"
        / "IRC 37 2018"
        / "IITPAVE",
    )
    for candidate in candidates:
        if (candidate / "IITPFILE.exe").is_file():
            return candidate
    searched = "\n".join(str(candidate) for candidate in candidates)
    raise FileNotFoundError(
        "Could not locate IITPFILE.exe. Checked these paths:\n" + searched
    )
