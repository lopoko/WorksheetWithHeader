# encoding: utf-8

import logging


def get_internal_logger(log_level=logging.DEBUG):
    """
    Return a Logger object based on the function name who call the get internal logger function.

    :param log_level: The log level defined in logging module.
    :return: the Logger object based on the function name who call the get internal logger function.
    """
    internal_logger = logging.getLogger(__name__)
    internal_logger.setLevel(log_level)

    if not internal_logger.hasHandlers():
        # Add a file handler only if no handler before. To avoid duplicate logging record.
        file_handler = logging.FileHandler('%s.log' % __name__, 'w+')
        file_handler.setLevel(log_level)

        logging_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(module)s - %(funcName)s - %(lineno)d - %(message)s')
        file_handler.setFormatter(logging_formatter)

        internal_logger.addHandler(file_handler)

    return internal_logger
