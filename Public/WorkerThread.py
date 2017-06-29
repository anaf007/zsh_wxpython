__author__ = 'Administrator'
import threading
class WorkerThread(threading.Thread):
    def __init__(self,windows):
        threading.Thread.__init__(self)
        self.window = windows
        self.timeToQuit = threading.Event()
        self.timeToQuit.clear()