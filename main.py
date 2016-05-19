#
# This file is part of The Principles of Modern Game AI.
# Copyright (c) 2015, AiGameDev.com KG.
#

import random
from concurrent.futures import ThreadPoolExecutor
import time
import vispy                    # Main application support.

import window                   # Terminal input and display.
import nltk.chat
import win32com.client


AGENT_RESPONSES = [
  (r'You are (worrying|scary|disturbing)',    # Pattern 1.
    ['Yes, I am %1.',                         # Response 1a.
     'Oh, sooo %1.']),

  (r'Are you ([\w\s]+)\?',                    # Pattern 2.
    ["Why would you think I am %1?",          # Response 2a.
     "Would you like me to be %1?"]),

  (r'',                                       # Pattern 3. (default)
    ["Is everything OK?",                     # Response 3a.
     "Can you still communicate?"])
]


DEFAULT_VOLUME = 75
DEFAULT_RATE = 0
MIN_RATE = -10
MAX_RATE = 10
MIN_VOLUME = 0
MAX_VOLUME = 100


class ActiveVoice:
    def __init__(self):
        # See https://msdn.microsoft.com/en-us/library/ms723602%28v=vs.85%29.aspx for doc
        self._voice = win32com.client.Dispatch("SAPI.SpVoice")
        self._executor = ThreadPoolExecutor(max_workers=1)

    def shutdown(self):
        self._executor.shutdown(False)

    def speak(self, text, volume=DEFAULT_VOLUME, rate=DEFAULT_RATE):
        volume = max(MIN_VOLUME, min(MAX_VOLUME, volume))
        rate = max(MIN_RATE, min(MAX_RATE, rate))

        def do_speak():
            self._voice.Volume = volume
            self._voice.Rate = rate
            self._voice.Speak(text)

        self._executor.submit(do_speak)

    def wait(self, time_in_sec):
        self._executor.submit(time.sleep, time_in_sec)

    def list_avatars(self): # pointless, only one on Windows 10!
        avatars = []
        for index, token in enumerate(self._voice.GetVoices("", "")):
            avatars.append( '%d: %s' % (index+1, token.GetDescription()) )
        return avatars


class HAL9000(object):
    
    def __init__(self, terminal):
        """Constructor for the agent, stores references to systems and initializes internal memory.
        """
        self.terminal = terminal
        self.location = 'unknown'
        self.already_greated = False
        self.chatbot = nltk.chat.Chat(AGENT_RESPONSES, nltk.chat.util.reflections)
        self.voice = ActiveVoice()
        self.last_alert_counter = 0

    def shutdown(self):
        self.voice.shutdown()

    def on_input(self, evt):
        """Called when user types anything in the terminal, connected via event.
        """
        if not self.already_greated:
            message = "Good evening! This is HAL."
            self.already_greated = True
            self.voice.speak(message)
            self.terminal.log(message, align='right', color='#00805A')
        if evt.text.lower() == "where am i?":
            message = "\u2014 You are in the {}. \u2014 ".format(self.location)
            self.terminal.log(message, align='center', color='#404040')
            self.voice.speak(message)
        else:
            message = self.chatbot.respond( evt.text )
            self.terminal.log(message, align='left', color='#404040')
            self.voice.speak(message)

    def on_command(self, evt):
        """Called when user types a command starting with `/` also done via events.
        """
        if evt.text == 'quit':
            vispy.app.quit()

        elif evt.text.startswith('relocate'):
            self.terminal.log('', align='center', color='#404040')
            new_location = evt.text[9:].strip()
            self.location = new_location
            message = '\u2014 Now in the {}. \u2014'.format(new_location)
            self.terminal.log(message, align='center', color='#404040')
            self.voice.speak(message)
        elif evt.text.startswith('avatars'):
            self.terminal.log('Available avatars:', align='left', color='#ff3000')
            for text in self.voice.list_avatars():
                self.terminal.log(text, align='left', color='#404040')
        else:
            self.terminal.log('Command `{}` unknown.'.format(evt.text), align='left', color='#ff3000')    
            self.terminal.log("I'm afraid I can't do that.", align='right', color='#00805A')

    def update(self, _):
        """Main update called once per second via the timer.
        """
        if self.last_alert_counter <= 0:
            event_selector = random.randint(0,100)
            if event_selector < 30:
                voice1 = dict(volume=MAX_VOLUME, rate=3)
                voice2 = dict(rate=2)
                if event_selector < 10:
                    message1 = "Alert! Asteroid on collision trajectory."
                    message2 = "Initiating avoidance protocol!"
                elif event_selector < 20:
                    message1 = "Massive Solar proton event detected!"
                    message2 = "System switching to self-preservation mode..."
                    voice2["rate"] = -5
                else:
                    message1 = "Alert! Fire detected at your location."
                    message2 = "Emitting FM-200 gas to suppress fire. You have 30 seconds to evacuate the room if you want to live."
                self.terminal.log(message1, align='center', color="#ff3000")
                self.terminal.log(message2, align='center', color="#ff8000")
                self.voice.speak(message1, **voice1)
                self.voice.speak(message2, **voice2)
                self.last_alert_counter = 20 # prevent new alert for 20s
        else:
            self.last_alert_counter -= 1


class Application(object):
    
    def __init__(self):
        # Create and open the window for user interaction.
        self.window = window.TerminalWindow()

        # Print some default lines in the terminal as hints.
        self.window.log('Operator started the chat.', align='left', color='#808080')
        self.window.log('HAL9000 joined.', align='right', color='#808080')

        # Construct and initialize the agent for this simulation.
        self.agent = HAL9000(self.window)

        # Connect the terminal's existing events.
        self.window.events.user_input.connect(self.agent.on_input)
        self.window.events.user_command.connect(self.agent.on_command)

    def run(self):
        timer = vispy.app.Timer(interval=1.0)
        timer.connect(self.agent.update)
        timer.start()
        
        vispy.app.run()

        self.agent.shutdown()


if __name__ == "__main__":
    vispy.set_log_level('WARNING')
    vispy.use(app='default')
    
    app = Application()
    app.run()
