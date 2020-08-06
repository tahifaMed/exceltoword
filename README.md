# ExcelToWordReplace
> Ce programme permet de copier du contenu source à partir d'un Excel fourni. 
chercher le contenu source dans un fichier Word et le remplacer par le contenu destination du fichier Excel

> Pour generer le jar du programme, il faut executer mvn package dans la racine du projet, apres completion
On trouvera le jar généré dans le dossier target avec le nom ExcelToWordReplace-1.0-SNAPSHOT-jar-with-dependencies.jar

> Pour avoir une idée sur les arguments necessaires pour lancer le programme , on execute le jar avec la commande :
```shell
java -jar ExcelToWordReplace-1.0-SNAPSHOT-jar-with-dependencies.jar --help

-x,--excel (Obligatoire) : le chemin absolu ou relatif du fichier excel qui contient les urls source et destinations.
-w,--word (Obligatoire)  : le chemin absolu ou relatif du fichier Word sur le quel on va modifier les urls sources par les urls destinations
-si,--sheet-index (Optionnel): l'index de la feuille excel sur laquelle est indiqué la liste source et destination, par defaut 0
-csi,--column-source-index (Optionnel): L'index de la collone source sur le tableau Excel, par defaut 2
-cdi,--column-destination-index (Optionnel): L'index de la collone destination sur le tableau Excel, par defaut 3
 ```
 > Exemple d'execution 
 ```shell
 java -jar ExcelToWordReplace-1.0-SNAPSHOT-jar-with-dependencies.jar -x D:\ExcelToWordReplace\mapping.xlsx -w D:\ExcelToWordReplace\InputFile.docx
 10:06:50.642 [main] INFO  com.cheuvreux.copernic.main.Main - ======> Program Start <=======
10:06:50.642 [main] INFO  com.cheuvreux.copernic.main.Main - start extract source and destination values from excel File
10:06:57.963 [main] INFO  com.cheuvreux.copernic.main.Main - excel values extracted with success, total lines : 6422
10:06:57.964 [main] INFO  com.cheuvreux.copernic.main.Main -  start replacing source value with destination value in Word File
Replacing  99% [================================================================================================================>] 6418/6422 (0:02:50 / 0:00:00)
10:09:50.954 [main] INFO  com.cheuvreux.copernic.main.Main -  Value change in Word File with Success : number of word changed : 0
10:09:50.954 [main] INFO  com.cheuvreux.copernic.main.Main - ======> Program Finished with success <=======
 ```
