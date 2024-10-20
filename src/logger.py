import logging

class LoggerInterface:
    def info(self, message: str):
        pass

    def error(self, message: str):
        pass

class ConsoleLogger(LoggerInterface):
    def __init__(self):
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def info(self, message: str):
        logging.info(message)

    def error(self, message: str):
        logging.error(message)