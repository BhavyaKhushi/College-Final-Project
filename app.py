import eel
import os
from queue import Queue

class ChatBot:
    started = False
    userinputQueue = Queue()

    @staticmethod
    def isUserInput():
        return not ChatBot.userinputQueue.empty()

    @staticmethod
    def popUserInput():
        return ChatBot.userinputQueue.get()

    @staticmethod
    def close_callback(route, websockets):
        # if not websockets:
        #     print('Bye!')
        exit()

    @eel.expose  # Expose this method to the front-end (JavaScript)
    @staticmethod
    def getUserInput(msg):
        ChatBot.userinputQueue.put(msg)
        print(msg)
    
    @staticmethod
    def close():
        ChatBot.started = False
    
    @eel.expose  # Expose this method to the front-end (JavaScript)
    @staticmethod
    def addUserMsg(msg):
        eel.addUserMsg(msg)  # This will call the JavaScript function from eel
    
    @eel.expose  # Expose this method to the front-end (JavaScript)
    @staticmethod
    def addAppMsg(msg):
        eel.addAppMsg(msg)  # This will call the JavaScript function from eel

    @staticmethod
    def start():
        path = os.path.dirname(os.path.abspath(__file__))
        eel.init(path + r'\web', allowed_extensions=['.js', '.html'])
        try:
            eel.start('index.html', mode='chrome',
                      host='localhost',
                      port=27005,
                      block=False,
                      size=(350, 480),
                      position=(10, 100),
                      disable_cache=True,
                      close_callback=ChatBot.close_callback)
            ChatBot.started = True
            while ChatBot.started:
                try:
                    eel.sleep(10.0)
                except:
                    # main thread exited
                    break
        except Exception as e:
            print(f"Error: {e}")
            pass
