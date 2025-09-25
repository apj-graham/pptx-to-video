import logging

logger = logging.getLogger("ppt_to_video")
logger.setLevel(logging.INFO)

# Console handler
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)

# Formatter
formatter = logging.Formatter(
    "%(asctime)s %(filename)s %(levelname)s - %(message)s", datefmt="%H:%M:%S"
)
ch.setFormatter(formatter)

# Avoid adding multiple handlers if re-imported
if not logger.hasHandlers():
    logger.addHandler(ch)
