import logging
from datetime import date
import os

today = str(date.today())

try:
    os.makedirs('./errorLog/')
except:
    pass

class Logger():
    def Log():
        logging.basicConfig(filename='./errorLog/'+today+'.log', level=logging.WARNING, 
                            format='%(asctime)s %(levelname)s %(name)s %(message)s')
        logger=logging.getLogger(__name__)
        return logger