import subprocess
import os

def main():
    risposta = input("Sono presenti misure 5G da analizzare? (s/n): ").strip().lower()
    if risposta == 's':
        script_to_run = "functions/EightTables_GSM_5G.py"
    else:
        script_to_run = "functions/EightTables_GSM.py"
    
    print(f"Esecuzione dello script: {script_to_run}")
    subprocess.run(["python", script_to_run])

if __name__ == "__main__":
    main()
