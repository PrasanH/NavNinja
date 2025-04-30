
import subprocess

from urls_to_open import url_dict 

# List of URLs
urls = [
    value for value in url_dict.values()
]

#Add your webbrowser path here 
chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"



# Build command to open all URLs in a new window
command = [chrome_path, "--new-window"] + urls

# Launch
subprocess.Popen(command)

