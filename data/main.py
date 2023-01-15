from Gen_Moy import ultime
import sys
import webbrowser


def launcher():
    webbrowser.open_new_tab('/home/etudiant/PROJETgitHUB_BM/docs/build/html/index.html') #ouvre la page de documentation de la SAE
    sys.stdout.write("Veuillez patienter pendant que le programme génère les bulletins..") #
    ultime()

def main():
    return (launcher())
    
#test unitaire
if __name__ == "__main__": 
    main()