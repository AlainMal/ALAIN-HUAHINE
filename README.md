                                 Application bus CAN et NMEA 2000 sous Excel
 
Il s'agit d'un application sous Excel qui gére les trames NMEA 2000
J'ai commencer par m'intessé au réaseau NMEA 2000 et j'ai appris le bus CAN avec son Identificateur, son nombre d'octets qui donne les valeurs et la priorité.
 Donc je me suis mis à réaliser un fichier Excel sans le VBA qui analyse l'identificateur avec l'aide des PGN venant de CSS Electronics puis aux octets dont je me suis aidé avec le site [https://canboat.github.io/](https://canboat.github.io/canboat/canboat.html) mais c'était un peu lourd, je me suis rendu à l'évidence qu'il fallait de je programme en VBA.
 
 J'ai commencé avex le décodage du PGN puis aux octets, un simple Case OF pour les PGN que j'avais grace à l'USB CAN et son logiciel USB CAN V7.0 et V8.0 que j'ai un peu décodé en regardant les octets reçus et j'ai compris tous les PGN qui sont sur mon bateau qui dispose de :
 - Vents apparent et réel (130306, 128259)
 - Vitesses SOG et surface (129026, 130577, 127250)
 - Position Latitude et Longitude (129025)
 - Tensions et courants des batteries de service et démarage moteur (127508, 127506)
 - Profondeur (128267)
 - Températures cabine et moteur (130312, 130316, 130310)
 - Niveaux d'eau douce et de gasoil (127505) 
 - Pilote automatique (127245)
 - AIS (129038, 129794, 129029, 128809, 129810)
Il y a d'autres PGN que j'ai décodé mais qui sont moins intêressent pour mon usage.

Maintenant il fallait que je capte le bus CAN en temps réel, je n'ai pas trouvé l'ideal et donc je me suis intéressé à l'Arduino qui est un appareil pour apprendre la programmation et qui fonctionne avec un micro-contôleur à 16Mhz et de mémoire restrainte de 250Kctets, j'ai acheter un MCP2515 à 8Mhz, tous ce matériel fonctionne à moitier avec des des morceaux de câbles dont je ne suis pas sûr qu'il font réellement contact.

Enfin je suis arrivé à faire fonctionner tous ce bazard, il à fallut que je programme la connexion à l'Arduino depuis mon Excel en VBA, j'ai heureusement trouvé trouver le programme qui lit les octets grace à "ModComm" dont j'ai oublié le nom du site qui utilise les API de Windows, ça n'a pas été facile même pour moi qui connait bien les réseaux indutriels car je recevez des octets en vrac mais que j'avais mis en forme dans l'Arduino avec cette représentation :
" .Identificateur;NombreOctets:octet,octet,octet  .... ,octet? "
Je ne m'éternise pas plus pour l'instant vous pouvez voir mon pravec une Dogramme dans le module "Communication.bas" ou plus simplement dans "Huahine PGN 3.5.xlsm"
J'ai trouvé un USBCAN de Lawicel qui fonctionne trés bien avec une DLL voyez mon programme dans "COM_CAN.xslm"
