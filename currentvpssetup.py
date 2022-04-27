'''
Created on Sep 20, 2021

See attached documentation for detailed description.

Wrapper for current_selenium.py for headless (no attached display) Linux systems. Runs current_selenium.py and notifies webhook (in this case Slack) of thrown exceptions.
Can be used with any headless Python script requiring a display to function properly.

Compatibility: Linux, Windows/MacOS requires minor edits.
'''

from subprocess import Popen
import slack
from dotenv import load_dotenv
import os

load_dotenv()

slack_token = os.environ.get("slack_token")

client = slack.WebClient(token = slack_token)
client.chat_postMessage(channel='C027YMXQM7U', text='Python Cleantech Web Scraper active!')

script = Popen(["xvfb-run", "python3", "current_selenium.py"])
script.communicate()
if script.returncode != 0:
    client.chat_postMessage(channel='C027YMXQM7U', text='Python Cleantech Web Scraper has run into an oopsie. Error code ' + str(script.returncode))