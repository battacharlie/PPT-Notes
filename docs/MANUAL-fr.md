# MEFISTOFELE DeckWhisperer – Manuel (Français)

**Remote notes & control for PowerPoint**
[English](MANUAL-en.md) · [Italiano](MANUAL-it.md) · [Français](MANUAL-fr.md) · [Español](MANUAL-es.md) · [العربية](MANUAL-ar.md) · [中文](MANUAL-zh.md)

Dans ce manuel, la **page de gestion** est la page qui s'ouvre sur l'ordinateur où tourne PowerPoint ; la **page distante** est celle que vous voyez sur le téléphone.

MEFISTOFELE DeckWhisperer lit les **notes du présentateur** de PowerPoint pendant la présentation et les affiche en temps réel
sur une page web ouverte sur votre téléphone (ou tout appareil du même réseau). Depuis le téléphone, vous pouvez aussi avancer et reculer dans les diapos,
aller directement à une diapo et voir la diapo en cours et la suivante.

## Prérequis
- **Windows 10/11** avec **PowerPoint** de bureau, ou un **Mac** (macOS 12 ou plus récent) avec **PowerPoint pour Mac**.
- L'ordinateur et le téléphone sur le **même réseau** (Wi-Fi ou câble).

## Démarrage rapide
**Windows**
1. Téléchargez `mefistofele.exe` depuis la page [Releases](https://github.com/battacharlie/PPT-Notes/releases) (dernière version).
2. Double-cliquez dessus. Si Windows affiche « Windows a protégé votre ordinateur » : **Informations complémentaires → Exécuter quand même**. Au premier lancement, autorisez les **réseaux privés**.

**Mac**
1. Téléchargez `mefistofele-mac.zip` depuis la page [Releases](https://github.com/battacharlie/PPT-Notes/releases) et décompressez-le.
2. Ouvrez le **Terminal**, tapez `chmod +x ` (avec l'espace), glissez le fichier `mefistofele` dans la fenêtre et appuyez sur Entrée. Recommencez sans `chmod +x ` pour le lancer.
   Si macOS le bloque (« ne peut pas être ouvert car le développeur n'est pas vérifié ») : exécutez `xattr -d com.apple.quarantine ` suivi du fichier, ou **Réglages Système → Confidentialité et sécurité → Ouvrir quand même**.
3. La première fois, macOS demande si **le Terminal peut contrôler Microsoft PowerPoint** : cliquez sur **OK** (modifiable ensuite dans **Réglages Système → Confidentialité et sécurité → Automatisation**).

**Ensuite (sur les deux systèmes)**
- La **page de gestion** s'ouvre dans le navigateur avec un QR code. Scannez-le avec votre téléphone.
- Dans PowerPoint, lancez le diaporama (**Diaporama → Depuis le début**) : les notes apparaissent sur le téléphone et suivent les diapos.

Le programme tourne dans une fenêtre noire (le Terminal sur Mac) : **fermez-la pour arrêter le programme**.

## La page de gestion (sur l'ordinateur)
Adresse : `http://localhost:8080/start` (s'ouvre toute seule). Le bouton **Guide** en haut à droite ouvre ce manuel dans la langue choisie.

**Langue.** Choisissez la langue dans le menu **Langue** : English (par défaut), Italiano, Français, Español, العربية ou 中文.
Elle s'applique à la page de gestion, à la page distante sur **tous les téléphones** (ils suivent la page de gestion) et au guide. Le choix est enregistré.

**Carte réseau du QR.** Si l'ordinateur a plusieurs cartes réseau (par exemple Wi-Fi et réseau de production), choisissez dans le menu
**Carte réseau du QR** celle à laquelle les téléphones sont connectés : le QR et l'adresse changent aussitôt. *Automatique* laisse le programme choisir.

Choisissez ensuite **ce que les téléphones affichent**. Les éléments, dans l'ordre :

| Case | Effet sur le téléphone |
|---|---|
| Barre d'outils | Le téléphone affiche la barre de boutons colorés, avec la bande et le cadenas. **Si elle est désactivée**, le téléphone n'affiche que les éléments choisis ci-dessous et **rien ne peut être modifié** |
| Style du texte | Police, taille, couleurs, interligne, espacement des lettres |
| Diapo en cours | Image de la diapo que voit le public |
| Diapo suivante | Image de la diapo suivante |
| Aller à la diapo | Champ numérique et bouton **Aller** pour sauter à une diapo |
| Navigation | Grands boutons **Suivant** et **Précédent** |
| Verrouiller l'écran distant | Retire **toutes les commandes** du téléphone : il ne reste que les notes, qu'on peut seulement **faire défiler** si elles dépassent l'écran, plus les aperçus choisis. Les commandes ne sont plus acceptées |

Vous pouvez aussi définir un **PIN** pour les commandes : le téléphone le demande la première fois que vous appuyez sur Suivant/Précédent ou Aller.
Vos choix sont **enregistrés** (`mefistofele.json` à côté du programme) et arrivent **instantanément** aux téléphones connectés.
La page de gestion s'ouvre **uniquement sur l'ordinateur** qui exécute le programme.

## La page distante (sur le téléphone)
Chaque fonction a sa couleur. Bouton **plein** = activé, bouton **contour seul** = désactivé.

| Couleur | Bouton | Fonction |
|---|---|---|
| Violet | Style du texte | Affiche/masque les commandes du texte |
| Vert | Navigation | Affiche/masque Suivant / Précédent et Aller à la diapo |
| Orange | Diapo en cours | Affiche/masque l'aperçu de la diapo en cours |
| Bleu | Diapo suivante | Affiche/masque l'aperçu de la diapo suivante |
| Turquoise | Plein écran | Active/quitte le plein écran (indisponible sur iPhone) |
| Rouge | Verrouiller l'écran | Verrouille le téléphone contre les touches accidentelles |
| Gris | QR / Infos | QR pour connecter un autre téléphone / version, guide et crédits |
| Jaune | Masquer la barre | Ferme la barre |

**Taille du texte** : saisissez le nombre dans le champ, ou utilisez **a−** (plus petit) et **A+** (plus grand).

**Bande du haut.** Quand vous fermez la barre, une bande reste avec quatre boutons : **Barre** (rouvre la barre),
**Écran** (plein écran), **Navigation** (affiche/masque Suivant, Précédent et Aller) et **Verrou**.

**Verrouillage de l'écran.** Avec le cadenas, l'écran ignore tous les contacts (les notes continuent de se mettre à jour).
Pour **déverrouiller, maintenez appuyé 5 secondes le bouton rouge en haut** : une barre se remplit puis l'écran est libéré. Si vous relâchez avant, il reste verrouillé.

**Le téléphone ne se met pas en veille** tant que la page est ouverte. Le premier contact sur la page active la protection ;
si le téléphone s'éteint quand même, vérifiez que l'*économiseur de batterie* n'est pas actif.

En bas : grands boutons **Précédent** et **Suivant**, plus le champ **Aller** pour sauter à un numéro de diapo.
Si plusieurs présentations sont ouvertes, choisissez celle à suivre dans le menu **Présentation** (indisponible avec *Verrouiller l'écran distant*).
Chaque téléphone mémorise ses propres choix (style du texte, éléments affichés). La langue est celle choisie dans la page de gestion.

## Accès à distance (téléphones hors du réseau)
Pour les téléphones qui ne sont **pas** sur le même réseau que le PC (données mobiles, autre Wi-Fi). **Désactivé par défaut** : ne l'activez que si nécessaire.
1. Dans la page de gestion, dans le cadre **Accès à distance**, appuyez sur **Activer l'accès à distance**.
2. Un lien privé (`https://remote.mefistofele.com:8443/s/...`) et un QR code apparaissent. Ouvrez le lien sur le téléphone distant.
3. Pour que ce téléphone puisse utiliser **Suivant/Précédent/Aller à la diapositive**, définissez un **PIN** dans la page de gestion : sans PIN, les téléphones distants ne peuvent que lire les notes. Le PIN est toujours demandé depuis l'extérieur et, après 5 essais erronés, les commandes sont bloquées pendant 5 minutes.
4. Appuyez sur **Désactiver** (ou fermez le programme) pour couper l'accès immédiatement. Le lien change à chaque activation.

Le PC se connecte **vers l'extérieur** au serveur relais MEFISTOFELE : aucun réglage du routeur ou du pare-feu n'est nécessaire. Les notes et les aperçus transitent par le relais chiffrés (HTTPS) et **ne sont pas enregistrés**.

## Options en ligne de commande
Windows : `mefistofele.exe --port 8080 --pin 1234 --no-control --no-browser --demo`
Mac : `./mefistofele --port 8080 --pin 1234 --no-control --no-browser --demo`

- `--port` port du serveur (8080 par défaut)
- `--pin` PIN pour les commandes
- `--no-control` désactive complètement les commandes depuis le téléphone
- `--no-browser` n'ouvre pas la page de gestion sur l'ordinateur
- `--demo` essai sans PowerPoint
- `--remote` active l'accès à distance au démarrage
- `--relay` adresse du serveur relais (par défaut `https://remote.mefistofele.com:8443`)

## Dépannage
- **Le téléphone ne se connecte pas** : même réseau ? Si l'ordinateur a plusieurs cartes, choisissez la bonne dans la page de gestion. Pare-feu Windows : autorisez les *réseaux privés* (et publics si votre Wi-Fi est vu comme public). Sur Mac, autorisez les connexions entrantes si on vous le demande.
- **Aucune note affichée** : le diaporama doit être **en cours** (plein écran), pas seulement ouvert.
- **Mac : rien ne se passe** : vérifiez **Réglages Système → Confidentialité et sécurité → Automatisation** et que le Terminal peut contrôler Microsoft PowerPoint.
- **Aperçus des diapos** : sur Mac, ils dépendent de l'autorisation d'export par PowerPoint ; en cas d'échec, la raison s'affiche à la place de l'image et les notes et commandes continuent de fonctionner.
- **Pas d'aperçus (Windows)** : ouvrez `http://localhost:8080/diagnostica`, appuyez sur le bouton de test et copiez tout.
- Les erreurs sont aussi écrites dans `mefistofele.log` à côté du programme (ou dans le dossier temporaire).

---
Created and developed by [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina
