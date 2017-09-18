import threading
import time

class convertThread(threading.Thread):
    def __init__(self,que):
        threading.Thread.__init__(self)
        self.daemon = False
        self.queue = que
    def run(self):
        while True:
            if self.queue.empty():
                break
            item = self.queue.get()
            #processing the item  
            time.sleep(item) 
            self.queue.task_done()