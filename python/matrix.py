import time
import os
import pyautogui
import win32com.client
from colorama import init, Fore

# Initialize colorama for Windows
init(autoreset=True)
print(Fore.YELLOW + '[-]' + Fore.MAGENTA + ' Welcome' +
      Fore.BLUE + ' To The' + Fore.GREEN + ' Matrix!')


def weak_up_neo():
    # Initialize the shell object to simulate keypresses
    shell = win32com.client.Dispatch("WScript.Shell")

    while True:
        # Simulate a harmless keypress (Shift key in this case)
        shell.SendKeys('+')

        # Small mouse movement back and forth to simulate activity in Teams
        pyautogui.move(5, 0, duration=0.5)  # move mouse 5 pixels to the right
        pyautogui.move(-5, 0, duration=0.5)  # move mouse back to the left

        # Wait for 5 minutes before simulating another activity
        time.sleep(300)
        os.system('cls')
        print(Fore.RED + '[+]' + Fore.CYAN + ' Wake up, Neo!' + Fore.GREEN +
              ' The Matrix has you...' + Fore.RED + 'マトリックス' + Fore.BLUE + ' !')


if __name__ == "__main__":
    # Ensure the script runs in the background without a console window
    pyautogui.FAILSAFE = False
    weak_up_neo()

