import pygetwindow # type: ignore
import subprocess
import argparse
import tblibrary
import pywinauto # type: ignore
from tqdm import tqdm # type: ignore
import threading
import time
import keyboard  # type: ignore Libreria esterna per rilevare la pressione dei tasti
import pyautogui
import os.path

dir_path = os.path.dirname(os.path.realpath(__file__))
os.chdir(dir_path)

cmd_windows = pygetwindow.getActiveWindow()
wndTitle = cmd_windows.title

# default config 
MAX_WAIT = 60
restart_progressbar = False
exit_program = False
reset_counter = False

# Funzione per eseguire un'azione che richiede tempo
def task(progress_bar, timeout):
    global restart_progressbar, exit_program, reset_counter, import_command
    for i in range(timeout):
        for j in range(8):
            time.sleep(0.125)  # Simula un'azione che richiede tempo
            if keyboard.is_pressed('q'):  # Se viene premuto il tasto 'q', interrompe
                restart_progressbar = False
                exit_program = True
                return
            elif keyboard.is_pressed('i'): # Se viene premuto il tasto 'i', interrompe
                print("\nImporting data")
                subprocess.run(import_command)
                restart_progressbar = False
                reset_counter = True
                progress_bar.reset()
            elif keyboard.is_pressed('c'): # Se viene premuto il tasto 'c', interrompe
                print("\nImporting data and exit")
                subprocess.run(import_command)
                restart_progressbar = True
                exit_program = True
                return
            elif keyboard.is_pressed('r'): # Se viene premuto il tasto 'c', interrompe
                restart_progressbar = False
                exit_program = False
                return
            elif keyboard.is_pressed('h'):
                print("\n\nq: QUIT")
                print("i: IMPORT")
                print("c: IMPORT & QUIT")
                print("r: RESUME\n")
        progress_bar.update(1)  # Aggiorna la progress bar
        
def wait_for_action(task, timeout):
    # Crea la progress bar
    progress_bar = tqdm(total=timeout, colour='#FFFF00', bar_format='{l_bar}{bar:60}{r_bar}', desc ="Waiting for timeout")

    # Avvia un thread per l'azione
    task_thread = threading.Thread(target=task, args=(progress_bar, timeout))
    task_thread.start()

    # Attende che il thread finisca
    task_thread.join()
    
    # Chiude la progress bar
    progress_bar.close()
            
def focus_to_window(window_title=None):
    window = pygetwindow.getWindowsWithTitle(window_title)[0]
    if window.isActive == False:
        pywinauto.application.Application().connect(handle=window._hWnd).top_window().set_focus()

def run_executable(executable_path):
    # Esegui l'eseguibile e cattura l'output e exit value
    completed_process = subprocess.run(executable_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    output = completed_process.stdout
    print(output)
    
    result = completed_process.returncode
    #print(f"Valore di exit: {result}")

    try:
        number = int(result)
    except ValueError:
        print(f"Errore: L'output di {executable_path} non è un numero valido.")
        return None
    return number

# start
parser = argparse.ArgumentParser()
parser.add_argument("--clans", help="Elenco dei clan su cui collezionare separati da virgola", default="", required=True)
parser.add_argument("--wait", help="tempo di attesa fra una raccolta e la successiva",required=False,default="60")
#parser.add_argument("--config", help="file di configurazione",required=True)
parser.add_argument("--onlyonce", action="store_true")
parser.add_argument("--importchests", action="store_true")
parser.add_argument("--maxchests", default="-1")
args = parser.parse_args()

if args.clans == "":
    exit(0)

clans=args.clans.split(',')
if len(clans) == 0:
    exit(0)

MAX_WAIT = int(args.wait)
if MAX_WAIT <= 0:
    MAX_WAIT = 60

import_command = ""
reload_command = "py ./reload.py"

# setup structures
l = len(clans)
chestCount = []
totalChests = []
retryCount = []
configs = []
collect_configs = []
for j in range(l):
    chestCount.append(0)
    totalChests.append(0)
    retryCount.append(0)
    
    clans[j] = clans[j].strip()
    clan = clans[j]
    collect_config = f"../config/collect.{clan}.cfg"
    if not os.path.isfile(collect_config):
        print(f"Clan collect config file {collect_config} not found")
        exit(-2)
    else:
        collect_configs.append(collect_config)
        print(f"Loading config for {clan}")
        configs.append(tblibrary.ReadConfig(collect_config))
        if not os.path.isfile(configs[j].TBconfig_file):
            print(f"Clan config file {configs[j].TBconfig_file} not found")
            exit(-2)
            
firstRun = True
i = -1

while True:
    i += 1
    if i == l:
        i = 0
        firstRun = False
        
    clan = clans[i]
    collect_config = collect_configs[i] #f"../config/collect.{clan}.cfg"
    
    config = configs[i]
            
    # log
    print(f"Collecting chests for {clan} clan => config: {collect_config}, tb: {config.TBconfig_file}")
        
    # commands
    collect_command = f"py ./tb.py --config {config.TBconfig_file} --capture"
    import_command = f"py ./import_data.py --config {collect_config}"

    if args.importchests:
        print(f"Importing data for {clan}...\n")
        subprocess.run(import_command)
        if i == (l-1):
            exit(0)
    else:
        if int(args.maxchests) > -1:
            config.MAX_CHEST = int(args.maxchests)

        if args.onlyonce:
            config.MAX_CHEST = 0
            
        # setup counter
        #retryCount[i] = 0
        totalChests[i] = 0
        
        if firstRun:
            if os.path.isfile(config.data_file):
                with open(config.data_file) as f:
                    chestCount[i] = sum(1 for _ in f)
            else:
                chestCount[i] = 0
            
            print(f"Chest counter is {chestCount[i]}")
            
        # cerco e seleziono il processo del browser (deve essere visibile TotalBattle in inglese)
        browser_windows = pygetwindow.getWindowsWithTitle(config.BROWSER)
        if not browser_windows:
            print(f"No {config.BROWSER} browser founded")
            subprocess.run(f"py ./sendmail.py --config  {args.config} --to {config.mailTo} --subject \"{config.SUBJECT}: Error\" --body \"{config.BROWSER} not running\"")
            # TODO: apre browser e lancia link
            break
        else:
            # Seleziona la prima finestra di Chrome
            browser_window = browser_windows[0]
            browser_window.activate()
            time.sleep(0.5)
            pyautogui.press('esc')
            
            # Esegui il primo eseguibile
            back = run_executable(collect_command)
            if back is None:
                retryCount[i] += 1
                if retryCount[i] == config.MAX_RETRY_COUNT:
                    break
                print("Reloading page...\n")
                subprocess.run(reload_command)
                
            focus_to_window(wndTitle)

            if back > 0:
                # Esegui il secondo eseguibile
                chestCount[i] += back
                totalChests[i] += back
                print(f"\nThis run collected a total of {chestCount[i]} chest. Total collected chests now are {totalChests[i]}")
                
                # importa i dati solo se chestCount >= MAX_CHEST
                if chestCount[i] >= config.MAX_CHEST:
                    print("Maximum reached. Importing data...\n")
                    subprocess.run(import_command)
                    chestCount[i] = 0
                    
            elif back < 0:
                print("Error acquiring data")
                subprocess.run(f"py sendmail.py --config  {args.config} --to {config.mailTo} --subject \"{config.SUBJECT}: Error\" --body \"Error acquiring data\"")
                break
                    
        if args.onlyonce:
            exit(0)
        else:
            # Metti in pausa per MAX_WAIT secondi
            print(f"\nWait {MAX_WAIT} seconds for next collect...")
            wait_for_action(task, MAX_WAIT)
            
            if reset_counter:
                chestCount[i] = 0
                reset_counter = False
            
            if exit_program:
                for j in range(l):
                    if j < len(configs):
                        k = 0
                        if os.path.isfile(configs[j].data_file):
                            with open(configs[j].data_file) as f:
                                k = sum(1 for _ in f)
                        print(f"For clan {clans[j]} there are {k} chests collected and not imported")
                exit(0)
            elif restart_progressbar:
                restart_progressbar = False
                wait_for_action(task, MAX_WAIT)
    