import pygetwindow as gw
import pyautogui
import time

# click 
mx    = 1869
my    = 256

# setup
#clan
ax	  = 1025
ay    = 930
#clan gifts
bx    = 570
by    = 480
# tab gifts
ex 	  = 805
ey 	  = 395

closex = 1419
closey = 340

# Trova la finestra del browser Chrome
chrome_windows = gw.getWindowsWithTitle('Chrome')
if not chrome_windows:
    print("Nessuna finestra di Chrome trovata")
    # TODO: apre chrome e lancia link
    exit(-1)
else:
    # Seleziona la prima finestra di Chrome
    chrome_window = chrome_windows[0]

    # Porta la finestra in primo piano
    chrome_window.activate()

    time.sleep(2)

    # Simula la pressione di F5 per aggiornare la pagina
    pyautogui.hotkey('f5')

    # Attendi 30 secondi
    print("waiting for page reload")
    time.sleep(30)

    # Clicca il mouse in una posizione specifica (ad esempio, x=200, y=200)
    print("close shop")
    pyautogui.moveTo(x=mx, y=my)
    time.sleep(2)
    pyautogui.click()
    pyautogui.hotkey('esc')
    pyautogui.hotkey('esc')
    pyautogui.hotkey('esc')
    
    # Prepare for collect
    
    # 1. click on clan tab
    #pyautogui.moveTo(x=ax, y=ay)
    #time.sleep(2.0)
    #pyautogui.click()
    
    # 2. click on tab gift
    #pyautogui.moveTo(x=bx, y=by)
    #time.sleep(2.0)
    #pyautogui.click()
    
    # 3. click on tab gift
    #pyautogui.moveTo(x=ex, y=ey)
    #time.sleep(2.0)
    #pyautogui.click()
    
    # close clan window
    #pyautogui.moveTo(x=closex, y=closey)
    #time.sleep(1.0)
    #pyautogui.click()
        
    exit(0)