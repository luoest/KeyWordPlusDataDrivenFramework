import logging
import logging.config
from baseUtil.PathAndDir import logConfigFile

try:
    logging.config.fileConfig(logConfigFile)
except:
    FileNotFoundError

logger = logging.getLogger('case715')

def info(message):
    logger.info(message)

def debug(message):
    logger.debug(message)

def warning(message):
    logger.warning(message)
