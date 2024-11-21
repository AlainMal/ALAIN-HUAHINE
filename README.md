Il s'agit d'un application sous Excel qui gére les trames NMEA 2000
J'ai commencer par m'intessé au réaseau NMEA 2000 et j'ai appris le bus CAN avec son Identificateur, son d'ombre d'octets et les fameux 8 octets
 qui donne les valeurs.
 Donc je me suis mis à réaliser un fichier Excel sans le VBA qui analyse l'identificateur avec l'aide des PGN venant de CSS Electronics puis aux octets dont je me suis aidé avec le site https://canboat.github.io/ mais c'était un peu lourd, je me rendu à l'évidence qu'il fallait de je le programme en VBA.
 
 J'ai commencé avex le décodage du PGN puis aux octets, un simple Case OF pour les PGN que j'avais grace à l'USB CAN et son logiciel USB CAN V7.0 et V8.0 que j'ai à moitié décodé en regardant les octets reçus, j'ai compris tous les PGN qui sont sur mon bateau qui dispose de :
 - Vent apparent et réel
 - Vitesse SOG et surface
 - Position Latitud et Longitude
 - Tensions et courants des batteries de service et démarage moteur
 - Profondeur
 - Température cabine et moteur
 - Niveau d'eau douce et de gasoil
 - Pilote automatique

J'ai aussi un émetteur-recepteur AIS donc je me suis intéressé aux données des bateaux environnants, après cela, j'avais tout fini, maintenant il fallait que je capte le bus CAN en temps réel
 
