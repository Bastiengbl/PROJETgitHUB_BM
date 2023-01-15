
Bienvenue sur la documentation de la SAE Traiter Des Donn√©es!
=============================================================
Menu:
=====
* :ref:`Usage`
* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`





Usage
-----

.. code-block:: bash

    $ python3 ./main.py 
    
    

ATTENTION, le code ayant ete redige avec des chemins absolus, le programme ne marchera pas si vous n'etes pas place dans un repertoire sous la forme  '/home/etudiant/
Les codes etant dependants les uns des autres, il est conseille de ne pas bouger un fichier d'un repertoire ‡ un autre.




    
1. Pour pouvoir lancer le programme principal afin de generer les bulletins des eleves, il vous faudra lancer le programme
'main.py' situe dans PROJETgitHUB/data

2. Pour ce faire vous pouvez lancer le programme dans le terminal linux avec la commande suivante:

.. code-block:: bash

    $ ./path/to/file/main.py

3. Vous pouvez eventuellement changer les droits sur le fichier si vous ne pouvez pas le lancer:

.. code-block:: bash

    $ chmod a+x /path/to/file/main.py

4. Si les erreurs persistent et que vous ne pouvez pas lancer le fichier, ouvrez le fichier main.py et lancez le programme depuis le terminal python.



5. Pour visualiser les rÈsultats, accedez au dossier data/ puis Bulletin/ pour les bulletins des eleves.
Si vous souhaitez acceder seulement aux notes des eleves, accedez au dossier data/ puis notes_S1 et S2.
En cas de probleme veuillez nous contacter ci-dessous:
   'bastien.gibel@etu.univ-poitiers.fr'
   'matthias.coureau@etu.univ-poitiers.fr'






Voici la documentation des modules python utilis√©s dans les programmes:
-----------------------------------------------------------------------

A propos du module openpyxl: 
----------------------------
.. automodule:: openpyxl
     :members: matthias coureau bastien gibel
     
     
openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files. 

It was born from lack of existing library to read/write natively from Python the Office Open XML format.

All kudos to the PHPExcel team as openpyxl was initially based on PHPExcel.


A propos du module random: 
--------------------------
.. automodule:: random
     :members: matthias coureau bastien gibel
     
A propos du module os: 
-----------------------
.. automodule:: os
     :members: matthias coureau bastien gibel


A propos du module shutil: 
--------------------------
.. automodule:: shutil
     :members: matthias coureau bastien gibel     




.. toctree::
   :maxdepth: 2
   :caption: Contents:



