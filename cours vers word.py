# ============================================================
#  CONVERTISSEUR DE COURS → WORD  (LPA / DPFC Côte d'Ivoire)
#  Google Colab — python-docx  — VERSION 4 (style modèle v2)
#
#  UTILISATION :
#  1. Installe : !pip install python-docx -q
#  2. Colle ton cours dans la variable COURS ci-dessous
#  3. Exécute → téléchargement automatique
#
#  FORMAT DU COURS — BALISES RECONNUES :
#
#  EN-TÊTE (les 2 formats sont acceptés) :
#  ════...════  ligne titre  ════...════     ← en-tête ASCII
#  TITRE_LECON: ...  COMPETENCE: ...        ← méta-lignes optionnelles
#  PROFESSEUR: ...   ORGANISME: ...         ← (auto-extraits de l'en-tête si absents)
#
#  SOMMAIRE (2 formats) :
#  "SOMMAIRE\n  1. Avant la leçon\n  2. La leçon..."   ← ancien format numéroté
#  "SOMMAIRE\n  Section 1 — ...\n    · Phase 1..."      ← nouveau format avec ·
#
#  SECTION 1 — AVANT LA LEÇON
#  ─────────────────────────────────
#  [CAPSULE DE REPRISE — ...]        ← avec ou sans crochets
#   CAPSULE DE REPRISE — ...         ← les deux marchent
#  [ANCRAGE IVOIRIEN]                ← avec ou sans crochets
#   ANCRAGE IVOIRIEN                 ← les deux marchent
#  ▶ HARKNESS N — ...
#
#  SECTION 2 — LA LEÇON
#  Phase N · ...
#  A. TITRE / B. TITRE / C. TITRE    ← sous-sections Phase 2
#  TABLEAU DES STRUCTURES — ...      ← titre de tableau
#  | col | col | col |               ← tableau markdown (auto-converti)
#  ~~~  texte anglais                ← texte de découverte
#  ---TRADUCTION---
#  texte français
#  ~~~
#  ┌─ À RETENIR — ... ─┐ ... └─┘
#  ⚠ ATTENTION — ...
#  FORMATION DU ...                  ← boîte retenir auto
#  ◆ DÉFINITION DPFC — ...
#  ✎ TRACE ÉCRITE
#  ~ ANALOGIE CI 1 / ~ ANALOGIE CI 2
#  5 points essentiels à retenir :   ← déclenche la boîte synthèse
#  SYNTHESE_DEBUT ... SYNTHESE_FIN   ← ou balises explicites
#  ◉ EXERCICE GUIDÉ N — ...
#
#  SECTION 3 — APRÈS LA LEÇON
#  ◎ EXERCICE N — ...
#  ◈ DEVOIR MAISON — ...
#
#  SECTION 4 — CORRIGÉS COMPLETS
#  ✔ CORRIGÉ — ...
#
#  FIN DU COURS :
#  CITATION FINALE                   ← balise optionnelle
#  "Citation en anglais..."
#  (Traduction française)
#  — Auteur
#  Prof. Evelyne ASSI · ... · DPFC/MENET-FP CI · 2025-2026
# ============================================================

COURS = """✅ Anglais LV1 · Leçon 22 — The computer · en cours de génération...

---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 22 · Classe de 3ème

THE COMPUTER
*L'ordinateur*

Compétence C8 · TIC / ICT

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 22 — THE COMPUTER
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C8 · TIC / ICT
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 21 : HIV/AIDS)
    · Ancrage ivoirien
    · Questions Harkness 1, 2 et 3
  Section 2 — La leçon
    · Phase 1 — Présentation & prérequis + Tableau des structures fondamentales
    · Phase 2 — Découverte guidée
    · Phase 3 — Schémas textuels
    · Phase 4 — Définitions DPFC
    · Phase 5 — Trace écrite + Analogies CI
    · Phase 5b — Synthèse
    · Phase 6 — Exercices guidés (2)
  Section 3 — Après la leçon
    · Exercices 1 à 5 + Devoir maison
  Section 4 — Corrigés complets
    · Corrigé du devoir maison + Exercices 1 à 5

────────────────────────────────────────────────────────
SECTION 1 — AVANT LA LEÇON
────────────────────────────────────────────────────────

CAPSULE DE REPRISE — Leçon 21 : HIV/AIDS (Le VIH/SIDA)

5 points clés à retenir :
  1. HIV est le virus ; AIDS est la maladie au stade avancé — deux choses distinctes.
  2. Le VIH se transmet par le sang infecté, les rapports sexuels non protégés
     et de la mère à l'enfant.
  3. Le VIH ne se transmet PAS par les poignées de main, les câlins,
     le partage de nourriture ou les piqûres de moustiques.
  4. Il n'existe pas de remède, mais le traitement ARV permet de vivre normalement.
  5. La stigmatisation aggrave l'épidémie — respect et solidarité sont essentiels.

3 questions flash :
  Q1 : What does HIV stand for?
  R1 : HIV stands for Human Immunodeficiency Virus.
       (HIV signifie Virus de l'Immunodéficience Humaine.)

  Q2 : Name two ways HIV is NOT transmitted.
  R2 : HIV is not transmitted by handshakes or by sharing food.
       (Le VIH ne se transmet pas par les poignées de main ou le partage de nourriture.)

  Q3 : What should we do to prevent HIV?
  R3 : We should use condoms, get tested regularly and avoid sharing needles.
       (Nous devrions utiliser des préservatifs, nous faire dépister régulièrement
       et éviter de partager des aiguilles.)

Lien avec la leçon du jour :
  La leçon 21 nous a montré comment les connaissances et la sensibilisation
  permettent de lutter contre une maladie grave.
  Aujourd'hui, nous découvrons un outil qui révolutionne l'accès à la connaissance
  et à l'information dans le monde entier : L'ORDINATEUR.
  Nous allons apprendre à le décrire, à expliquer son fonctionnement
  et à en parler en anglais.

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

Issa est élève en 3ème au Lycée Technique de Bouaké.
Son lycée vient de recevoir une salle informatique neuve
grâce à un don du gouvernement ivoirien dans le cadre du projet
"École Numérique" du MENET-FP.
Le professeur d'informatique, M. Diomandé, fait découvrir aux élèves
les différentes parties de l'ordinateur.
Issa voit pour la première fois un écran, un clavier, une souris et une unité centrale.
Il se demande comment tout cela fonctionne et à quoi cela peut lui servir
pour préparer son BEPC.

→ Comment s'appelle chaque partie d'un ordinateur en anglais ?
→ À quoi sert un ordinateur ?
→ Quels types d'ordinateurs existe-t-il ?
→ C'est exactement ce que nous allons apprendre aujourd'hui.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — What is a computer?
  Q : What is a computer and what is it used for?
      (Qu'est-ce qu'un ordinateur et à quoi sert-il ?)
  Guide :
    · Pense à ce que fait un ordinateur : traiter, stocker, afficher des informations.
    · Pense aux différentes utilisations : école, travail, communication, divertissement.
    · Pense à la différence entre un ordinateur de bureau et un ordinateur portable.
  Corrigé :
    A computer is an electronic machine that can receive, store, process
    and display information. It is used for many purposes : writing documents,
    doing calculations, searching for information, communicating with others,
    watching videos and playing games.
    There are two main types of computers : the desktop computer, which stays
    in one place, and the laptop computer, which is portable and can be carried
    anywhere. In Côte d'Ivoire, computers are increasingly used in schools,
    offices and businesses to improve productivity and access to information.
    (Un ordinateur est une machine électronique qui peut recevoir, stocker,
    traiter et afficher des informations. Il est utilisé à de nombreuses fins :
    écrire des documents, faire des calculs, chercher des informations,
    communiquer avec d'autres, regarder des vidéos et jouer à des jeux.
    Il existe deux types principaux d'ordinateurs : l'ordinateur de bureau,
    qui reste en un endroit fixe, et l'ordinateur portable, qui peut être
    transporté partout. En Côte d'Ivoire, les ordinateurs sont de plus en
    plus utilisés dans les écoles, les bureaux et les entreprises pour améliorer
    la productivité et l'accès à l'information.)

▶ HARKNESS 2 — Hardware and software
  Q : What is the difference between hardware and software?
      (Quelle est la différence entre le matériel et le logiciel ?)
  Guide :
    · Pense à ce que tu peux toucher (matériel physique) et ce que tu ne peux pas toucher (programmes).
    · Donne des exemples de hardware : écran, clavier, souris, imprimante.
    · Donne des exemples de software : système d'exploitation, traitement de texte, navigateur web.
  Corrigé :
    Hardware refers to the physical parts of a computer that you can see and touch,
    such as the screen (monitor), the keyboard, the mouse, the printer and the
    central processing unit (CPU). These are the components that make up the
    physical machine.
    Software refers to the programmes and instructions that tell the computer
    what to do. You cannot touch software. Examples include the operating system
    (like Windows or Linux), word processing programmes (like Microsoft Word),
    and web browsers (like Chrome or Firefox).
    In simple terms : hardware is the body of the computer, software is its brain.
    (Le matériel désigne les parties physiques d'un ordinateur que vous pouvez
    voir et toucher, comme l'écran, le clavier, la souris, l'imprimante et
    l'unité centrale. Ce sont les composants qui constituent la machine physique.
    Le logiciel désigne les programmes et les instructions qui disent à l'ordinateur
    quoi faire. On ne peut pas toucher un logiciel. Les exemples incluent le
    système d'exploitation, les programmes de traitement de texte et les navigateurs web.
    En termes simples : le matériel est le corps de l'ordinateur, le logiciel est son cerveau.)

▶ HARKNESS 3 — Computers and education in Côte d'Ivoire
  Q : How can computers help students in Côte d'Ivoire?
      (Comment les ordinateurs peuvent-ils aider les élèves en Côte d'Ivoire ?)
  Guide :
    · Pense à l'accès aux ressources éducatives en ligne.
    · Pense aux défis : électricité, connexion internet, coût des appareils.
    · Pense aux initiatives du gouvernement ivoirien pour le numérique à l'école.
  Corrigé :
    Computers can help students in Côte d'Ivoire in many ways. They allow
    students to access educational resources online, write documents, do research
    and learn new skills. Computer literacy is increasingly important for finding
    a job in the modern world.
    However, there are challenges : not all schools have computers, electricity
    is not always available in rural areas, and internet access can be expensive.
    The Ivorian government has launched several programmes to bring computers
    to schools, such as the "École Numérique" initiative. Despite the challenges,
    more and more Ivorian students are gaining access to digital tools
    that will help them succeed in the future.
    (Les ordinateurs peuvent aider les élèves de Côte d'Ivoire de nombreuses façons.
    Ils permettent aux élèves d'accéder aux ressources éducatives en ligne,
    d'écrire des documents, de faire des recherches et d'apprendre de nouvelles
    compétences. La maîtrise de l'informatique est de plus en plus importante
    pour trouver un emploi dans le monde moderne.
    Cependant, il y a des défis : toutes les écoles n'ont pas d'ordinateurs,
    l'électricité n'est pas toujours disponible dans les zones rurales et
    l'accès à internet peut être coûteux. Le gouvernement ivoirien a lancé
    plusieurs programmes pour apporter des ordinateurs dans les écoles.
    Malgré les défis, de plus en plus d'élèves ivoiriens ont accès aux outils
    numériques qui les aideront à réussir dans le futur.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre : The computer — L'ordinateur
Compétence : C8 · TIC / ICT

Objectifs de la leçon :
  → Nommer et décrire les parties d'un ordinateur en anglais.
  → Distinguer hardware (matériel) et software (logiciel).
  → Utiliser le Present Simple pour décrire les fonctions d'un ordinateur.
  → Utiliser "used to + infinitif" / "used for + V-ing" pour exprimer l'utilité.
  → Enrichir son vocabulaire informatique en anglais.

Prérequis nécessaires :
  → Le Present Simple pour les faits généraux et les habitudes
  → Les modaux MUST / SHOULD vus en leçons 20 et 21
  → Vocabulaire de base : machine · screen · work · type · click
  → Notions générales sur la technologie

┌─ TABLEAU DES STRUCTURES FONDAMENTALES ──────────────────┐
  Expression / Notion          | Valeur, emploi et exemple traduit
  ─────────────────────────────|──────────────────────────────────
  A computer is used to + inf. | Exprimer la fonction / l'utilité
                               | "A computer is used to store data."
                               | (Un ordinateur est utilisé pour stocker
                               |  des données.)
  ─────────────────────────────|──────────────────────────────────
  A computer is used for + ing | Exprimer l'utilité (forme nominale)
                               | "A computer is used for writing documents."
                               | (Un ordinateur est utilisé pour écrire
                               |  des documents.)
  ─────────────────────────────|──────────────────────────────────
  It allows you to + inf.      | Exprimer ce que quelque chose permet
                               | "It allows you to send emails."
                               | (Ça te permet d'envoyer des e-mails.)
  ─────────────────────────────|──────────────────────────────────
  It consists of...            | Décrire la composition
                               | "A computer consists of a monitor,
                               |  a keyboard and a mouse."
                               | (Un ordinateur est composé d'un écran,
                               |  d'un clavier et d'une souris.)
  ─────────────────────────────|──────────────────────────────────
  You can use it to + inf.     | Exprimer une possibilité d'utilisation
                               | "You can use it to do your homework."
                               | (Tu peux l'utiliser pour faire tes devoirs.)
  ─────────────────────────────|──────────────────────────────────
  It is made up of...          | Décrire les composants
                               | "The CPU is made up of several components."
                               | (L'unité centrale est composée de
                               |  plusieurs composants.)
└──────────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

TEXTE DE BASE — À lire et comprendre

  "A computer is an electronic machine used to process, store and display
  information. It consists of two main parts : hardware and software.
  Hardware includes all the physical components you can touch : the monitor
  (screen), the keyboard, the mouse, the printer and the CPU (Central Processing Unit).
  Software includes all the programmes that run on the computer, such as
  the operating system, word processors and web browsers.
  Computers are used for many purposes : writing documents, sending emails,
  doing research, watching videos and playing games.
  In schools, computers help students to learn and access educational resources.
  In Côte d'Ivoire, the government is working to provide computers to all
  schools as part of the digital education programme."

  TRADUCTION FRANÇAISE :
  "Un ordinateur est une machine électronique utilisée pour traiter, stocker
  et afficher des informations. Il est composé de deux parties principales :
  le matériel et le logiciel.
  Le matériel comprend tous les composants physiques que vous pouvez toucher :
  le moniteur (écran), le clavier, la souris, l'imprimante et l'unité centrale.
  Le logiciel comprend tous les programmes qui fonctionnent sur l'ordinateur,
  comme le système d'exploitation, les traitements de texte et les navigateurs web.
  Les ordinateurs sont utilisés à de nombreuses fins : écrire des documents,
  envoyer des e-mails, faire des recherches, regarder des vidéos et jouer à des jeux.
  Dans les écoles, les ordinateurs aident les élèves à apprendre et à accéder
  aux ressources éducatives.
  En Côte d'Ivoire, le gouvernement travaille à fournir des ordinateurs à toutes
  les écoles dans le cadre du programme d'éducation numérique."

────────────────────────────────────────────────────────

┌─ À RETENIR — HARDWARE vs SOFTWARE ──────────────────────┐
  HARDWARE (Matériel) :
  → Ce que tu peux TOUCHER physiquement.
  → Prononciation : [ˈhɑːdweər]
  → Exemples :
     · Monitor / Screen (écran) [ˈmɒnɪtər]
     · Keyboard (clavier) [ˈkiːbɔːrd]
     · Mouse (souris) [maʊs]
     · Printer (imprimante) [ˈprɪntər]
     · CPU / Central Processing Unit (unité centrale) [siː piː juː]
     · Speaker (haut-parleur) [ˈspiːkər]
     · Webcam (webcam) [ˈwebkæm]

  SOFTWARE (Logiciel) :
  → Ce que tu NE PEUX PAS toucher — ce sont des programmes.
  → Prononciation : [ˈsɒftweər]
  → Exemples :
     · Operating system (système d'exploitation) — ex. Windows, Linux
     · Word processor (traitement de texte) — ex. Microsoft Word
     · Web browser (navigateur web) — ex. Chrome, Firefox
     · Antivirus (antivirus)
     · Application / App (application)

  MOYEN MNÉMOTECHNIQUE :
  HARD = dur, solide → ce qu'on peut toucher
  SOFT = doux, intangible → ce qu'on ne peut pas toucher
└──────────────────────────────────────────────────────────┘

⚠ ATTENTION — Faux amis et pièges
  · "actually" ≠ "actuellement"
    → actually = en réalité / en fait
    → actuellement = currently / nowadays
  · "library" ≠ "librairie"
    → library = bibliothèque
    → librairie = bookshop
  · "CPU" se dit lettre par lettre : [siː piː juː] — jamais "cup"
  · "mouse" au pluriel → "mice" [maɪs] — pas "mouses"
  · "keyboard" = clavier — attention : "key" = touche/clé · "board" = planche

────────────────────────────────────────────────────────

VOCABULAIRE ESSENTIEL — The computer

  ANGLAIS                    | PRONONCIATION          | TRADUCTION FR
  ───────────────────────────|────────────────────────|─────────────────────────
  computer                   | [kəmˈpjuːtər]          | ordinateur
  hardware                   | [ˈhɑːdweər]            | matériel informatique
  software                   | [ˈsɒftweər]            | logiciel
  monitor / screen           | [ˈmɒnɪtər]             | écran / moniteur
  keyboard                   | [ˈkiːbɔːrd]            | clavier
  mouse                      | [maʊs]                 | souris
  printer                    | [ˈprɪntər]             | imprimante
  CPU / central processing   | [siː piː juː]          | unité centrale
  unit                       |                        |
  speaker                    | [ˈspiːkər]             | haut-parleur
  webcam                     | [ˈwebkæm]              | webcam
  hard drive                 | [hɑːd draɪv]           | disque dur
  memory / RAM               | [ˈmeməri]              | mémoire vive / RAM
  operating system           | [ˈɒpəreɪtɪŋ ˈsɪstəm]  | système d'exploitation
  word processor             | [wɜːd ˈprəʊsesər]      | traitement de texte
  web browser                | [web ˈbraʊzər]         | navigateur web
  file                       | [faɪl]                 | fichier
  folder                     | [ˈfəʊldər]             | dossier
  to type                    | [taɪp]                 | taper (sur un clavier)
  to click                   | [klɪk]                 | cliquer
  to save                    | [seɪv]                 | sauvegarder / enregistrer
  to print                   | [prɪnt]                | imprimer
  to download                | [ˈdaʊnləʊd]            | télécharger
  to process                 | [ˈprəʊses]             | traiter (des données)
  data                       | [ˈdeɪtə]               | données
  desktop computer           | [ˈdesktɒp]             | ordinateur de bureau
  laptop computer            | [ˈlæptɒp]              | ordinateur portable

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Les parties d'un ordinateur / Parts of a computer]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Dessin simplifié d'un ordinateur de bureau au centre de la page.
  · Chaque composant est entouré d'une flèche pointant vers lui avec son nom.

  Composants à légender :
    → En haut à gauche : MONITOR / SCREEN (écran) — rectangle vertical
    → En bas au centre : KEYBOARD (clavier) — rectangle horizontal plat
    → À droite du clavier : MOUSE (souris) — forme ovale
    → À droite de l'écran : CPU / CENTRAL UNIT (unité centrale) — rectangle vertical
    → En bas à droite : PRINTER (imprimante) — rectangle horizontal
    → En haut à droite : SPEAKER (haut-parleur) — forme triangulaire

  · Chaque légende est en MAJUSCULES EN ANGLAIS
  · Traduction française entre parenthèses sous chaque légende
  · Titre en haut : PARTS OF A COMPUTER (Les parties d'un ordinateur)
  · Flèches fines reliant les étiquettes aux composants

Note générateur : Canva (diagram with labels) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Hardware vs Software : tableau de distinction]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Deux colonnes côte à côte séparées par une ligne verticale :

  Colonne gauche (fond bleu clair) — titre : HARDWARE (Matériel)
    Sous-titre : "What you CAN touch" (Ce que tu PEUX toucher)
    Icônes + noms :
      · Monitor (écran)
      · Keyboard (clavier)
      · Mouse (souris)
      · Printer (imprimante)
      · CPU (unité centrale)
      · Speaker (haut-parleur)

  Colonne droite (fond orange clair) — titre : SOFTWARE (Logiciel)
    Sous-titre : "What you CANNOT touch" (Ce que tu NE PEUX PAS toucher)
    Icônes + noms :
      · Operating system (système d'exploitation)
      · Word processor (traitement de texte)
      · Web browser (navigateur web)
      · Antivirus (antivirus)
      · Application / App (application)

  · Titre général en haut : HARDWARE vs SOFTWARE
  · Moyen mnémotechnique en bas :
    HARD = solide = on peut toucher · SOFT = doux = on ne peut pas toucher
Note générateur : Canva (Two-column infographic) ou Word simple
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Computer (Ordinateur)
  EN : A computer is an electronic machine that receives, processes, stores
  and displays information. It can perform many tasks quickly and accurately.
  FR : Un ordinateur est une machine électronique qui reçoit, traite, stocke
  et affiche des informations. Il peut accomplir de nombreuses tâches
  rapidement et avec précision.

◆ DÉFINITION DPFC — Hardware (Matériel informatique)
  EN : Hardware refers to all the physical and tangible components of a computer
  system that can be seen and touched, such as the monitor, keyboard, mouse
  and printer.
  FR : Le matériel informatique désigne tous les composants physiques et
  tangibles d'un système informatique que l'on peut voir et toucher, comme
  l'écran, le clavier, la souris et l'imprimante.

◆ DÉFINITION DPFC — Software (Logiciel)
  EN : Software refers to the programmes, instructions and data that tell
  a computer what to do. Software cannot be touched physically. Examples
  include operating systems and word processors.
  FR : Le logiciel désigne les programmes, les instructions et les données
  qui indiquent à un ordinateur quoi faire. Les logiciels ne peuvent pas être
  touchés physiquement. Les exemples comprennent les systèmes d'exploitation
  et les traitements de texte.

◆ DÉFINITION DPFC — CPU / Central Processing Unit (Unité centrale de traitement)
  EN : The CPU is the main component of a computer that executes instructions
  and processes data. It is often called the "brain" of the computer.
  FR : L'unité centrale de traitement est le composant principal d'un ordinateur
  qui exécute les instructions et traite les données. Elle est souvent appelée
  le "cerveau" de l'ordinateur.

◆ DÉFINITION DPFC — Operating system (Système d'exploitation)
  EN : An operating system is a software programme that manages all the
  hardware and software resources of a computer and allows the user
  to interact with the machine. Examples : Windows, Linux, macOS.
  FR : Un système d'exploitation est un programme logiciel qui gère toutes
  les ressources matérielles et logicielles d'un ordinateur et permet à
  l'utilisateur d'interagir avec la machine. Exemples : Windows, Linux, macOS.

◆ DÉFINITION DPFC — Data (Données)
  EN : Data refers to information in a form that can be processed or stored
  by a computer, such as numbers, text, images or sounds.
  FR : Les données désignent des informations sous une forme pouvant être
  traitée ou stockée par un ordinateur, comme des chiffres, du texte,
  des images ou des sons.

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

  ANGLAIS LV1 · Leçon 22 · 3ème
  THE COMPUTER — L'ordinateur
  Compétence C8 · TIC / ICT · Prof. Evelyne ASSI

  I. VOCABULAIRE ESSENTIEL
  ─────────────────────────
  computer [kəmˈpjuːtər] = ordinateur
  hardware [ˈhɑːdweər] = matériel informatique
  software [ˈsɒftweər] = logiciel
  monitor / screen = écran
  keyboard [ˈkiːbɔːrd] = clavier
  mouse [maʊs] = souris (pluriel : mice)
  printer = imprimante
  CPU = unité centrale (cerveau de l'ordinateur)
  hard drive = disque dur
  operating system = système d'exploitation
  web browser = navigateur web
  file = fichier · folder = dossier
  to type = taper · to click = cliquer
  to save = sauvegarder · to print = imprimer
  to download = télécharger · data = données
  desktop computer = ordinateur de bureau
  laptop computer = ordinateur portable

  II. HARDWARE vs SOFTWARE
  ─────────────────────────
  HARDWARE = matériel = ce qu'on peut TOUCHER :
    · Monitor (écran) · Keyboard (clavier) · Mouse (souris)
    · Printer (imprimante) · CPU (unité centrale) · Speaker (haut-parleur)

  SOFTWARE = logiciel = ce qu'on NE PEUT PAS toucher :
    · Operating system (système d'exploitation) — ex. Windows
    · Word processor (traitement de texte) — ex. Microsoft Word
    · Web browser (navigateur web) — ex. Chrome, Firefox
    · Antivirus (antivirus) · App / Application (application)

  Moyen mnémotechnique :
  HARD = dur = solide = on peut toucher
  SOFT = doux = intangible = on ne peut pas toucher

  III. FONCTIONS DE L'ORDINATEUR
  ────────────────────────────────
  A computer is used to... (Un ordinateur est utilisé pour...) :
    · process information (traiter des informations)
    · store data (stocker des données)
    · write documents (écrire des documents)
    · send emails (envoyer des e-mails)
    · search for information (chercher des informations)
    · watch videos (regarder des vidéos)
    · communicate with others (communiquer avec d'autres)

  IV. STRUCTURES CLÉS
  ────────────────────
  · A computer is used to store data.
    (Un ordinateur est utilisé pour stocker des données.)
  · It consists of a monitor, a keyboard and a mouse.
    (Il est composé d'un écran, d'un clavier et d'une souris.)
  · It allows you to send emails and do research.
    (Il te permet d'envoyer des e-mails et de faire des recherches.)
  · You can use it to write your homework.
    (Tu peux l'utiliser pour écrire tes devoirs.)

~ ANALOGIE CI
  Pense à la radio communautaire de ton quartier à Abidjan ou à Yamoussoukro.
  La radio, c'est le matériel — l'antenne, les câbles, les haut-parleurs : c'est le HARDWARE.
  Les émissions, la musique, les informations diffusées : c'est le SOFTWARE.
  Sans l'un, l'autre ne fonctionne pas.
  "A computer without software is like a radio without a signal."
  (Un ordinateur sans logiciel, c'est comme une radio sans signal.)

~ ANALOGIE CI 2
  Au marché de Treichville, le commerçant Kouadio utilise désormais
  une caisse enregistreuse électronique connectée à un ordinateur.
  Le boîtier de la caisse = hardware.
  Le programme qui calcule les prix et imprime les reçus = software.
  L'ordinateur a transformé le commerce en Côte d'Ivoire, de la boutique
  de quartier aux grandes entreprises d'Abidjan.
  "Technology is not the future — it is already the present."
  (La technologie n'est pas le futur — c'est déjà le présent.)

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. Un ordinateur est une machine électronique qui traite, stocke et affiche des informations.
  2. Le hardware désigne les composants physiques (écran, clavier, souris, imprimante).
  3. Le software désigne les programmes et logiciels (système d'exploitation, navigateur).
  4. Le CPU est l'unité centrale — le "cerveau" de l'ordinateur.
  5. "mouse" fait "mice" au pluriel — "keyboard" signifie clavier (key = touche, board = planche).

5 mots clés + définition courte :
  · Hardware → composants physiques et tangibles d'un ordinateur
  · Software → programmes et instructions non tangibles d'un ordinateur
  · CPU → cerveau de l'ordinateur, traite toutes les données
  · Data → informations stockées et traitées par l'ordinateur
  · Operating system → programme principal qui fait fonctionner l'ordinateur

3 erreurs fréquentes + correction :
  Erreur 1 : "The mouse is a software."
  Correction : "The mouse is hardware." — la souris est physique, on peut la toucher.

  Erreur 2 : "I have two mouses."
  Correction : "I have two mice." — le pluriel irrégulier de mouse est mice.

  Erreur 3 : Confondre "used to + inf." (utilité) et "used to + inf." (habitude passée).
  Correction : Dans cette leçon, "A computer is used to store data" = utilité présente.
               "I used to play games" = habitude dans le passé (leçon différente).
               Le contexte et le temps du verbe "to be" indiquent la différence.

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Hardware ou Software ? · Notion ciblée : Distinction hardware / software

  Énoncé :
  Classe chaque élément dans la bonne colonne : HARDWARE ou SOFTWARE.
  Traduis chaque mot en français.

  Éléments : keyboard · Windows · mouse · Chrome · printer · Microsoft Word ·
             monitor · antivirus · CPU · hard drive · Firefox · webcam

  MÉTHODE
  Étape 1 · Pour chaque élément, pose-toi la question :
            "Est-ce que je peux le toucher physiquement ?"
            → OUI : c'est du HARDWARE
            → NON : c'est du SOFTWARE
  Étape 2 · Place chaque élément dans la bonne colonne.
  Étape 3 · Traduis chaque élément en français.
  ✔ Vérification : Relis le bloc "À RETENIR" sur hardware vs software.
  ✔ Conclusion :

  HARDWARE (Matériel)          | SOFTWARE (Logiciel)
  ─────────────────────────────|──────────────────────────────
  keyboard (clavier)           | Windows (système d'exploitation)
  mouse (souris)               | Chrome (navigateur web)
  printer (imprimante)         | Microsoft Word (traitement de texte)
  monitor (écran)              | antivirus (antivirus)
  CPU (unité centrale)         | Firefox (navigateur web)
  hard drive (disque dur)      |
  webcam (webcam)              |

  ✔ Ce que cet exercice mobilise :
    · Distinction hardware / software
    · Vocabulaire informatique anglais → français
    · Capacité de classement et d'analyse
    · Connaissance des composants et programmes d'un ordinateur

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Décrire un ordinateur · Notion ciblée : Production écrite + Structures "used to / consists of / allows"

  Énoncé :
  Aminata, élève en 3ème à Abidjan, doit présenter l'ordinateur à ses camarades
  qui n'en ont jamais utilisé. Aide-la à écrire une présentation de 7 phrases
  en anglais, en utilisant les structures :
  "A computer is used to..." · "It consists of..." · "It allows you to..."
  Traduis chaque phrase en français.

  MÉTHODE
  Étape 1 · Commence par définir ce qu'est un ordinateur en une phrase.
  Étape 2 · Utilise "It consists of..." pour décrire les composants principaux.
  Étape 3 · Utilise "A computer is used to..." pour donner 2 fonctions.
  Étape 4 · Utilise "It allows you to..." pour donner 2 possibilités d'utilisation.
  Étape 5 · Termine par une phrase sur l'importance de l'ordinateur à l'école.
  Étape 6 · Traduis chaque phrase en français.
  ✔ Vérification : Vérifie que tu as utilisé les 3 structures au moins une fois chacune.
  ✔ Conclusion — Présentation d'Aminata :
    1. A computer is an electronic machine that processes and stores information.
       (Un ordinateur est une machine électronique qui traite et stocke des informations.)
    2. It consists of hardware and software.
       (Il est composé de matériel et de logiciel.)
    3. The hardware includes the monitor, the keyboard, the mouse and the CPU.
       (Le matériel comprend l'écran, le clavier, la souris et l'unité centrale.)
    4. A computer is used to write documents and send emails.
       (Un ordinateur est utilisé pour écrire des documents et envoyer des e-mails.)
    5. A computer is also used to store data and search for information online.
       (Un ordinateur est aussi utilisé pour stocker des données et chercher
       des informations en ligne.)
    6. It allows you to communicate with people all over the world.
       (Il te permet de communiquer avec des gens partout dans le monde.)
    7. In our school, computers allow us to learn more effectively and prepare
       for the modern world of work.
       (Dans notre école, les ordinateurs nous permettent d'apprendre plus
       efficacement et de nous préparer au monde du travail moderne.)
  ✔ Ce que cet exercice mobilise :
    · Production écrite guidée sur un thème TIC / ICT
    · Structures "used to / consists of / allows you to"
    · Vocabulaire informatique hardware et software
    · Traduction anglais → français
    · Ancrage ivoirien (élève d'Abidjan)

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulaire informatique · Notions travaillées : Lexique, traduction, prononciation

  Partie A — Traduis ces mots anglais en français.
    1. keyboard     →
    2. software     →
    3. printer      →
    4. to download  →
    5. hard drive   →
    6. folder       →
    7. data         →
    8. web browser  →

  Partie B — Traduis ces mots français en anglais.
    1. souris        →
    2. logiciel      →
    3. sauvegarder   →
    4. ordinateur portable →
    5. système d'exploitation →

  GUIDE
  Étape 1 · Relis le tableau de vocabulaire essentiel de la trace écrite.
  Étape 2 · Attention aux mots qui ressemblent au français mais sont différents.
  Étape 3 · Vérifie l'orthographe de chaque mot anglais.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Hardware ou Software ? · Notions travaillées : Classement, distinction matériel/logiciel

  Lis cette liste et écris H (Hardware) ou S (Software) devant chaque élément.
  Traduis chaque élément en français.

  1. __ antivirus
  2. __ monitor
  3. __ operating system
  4. __ webcam
  5. __ Microsoft Word
  6. __ keyboard
  7. __ Firefox
  8. __ printer
  9. __ CPU
  10. __ application

  GUIDE
  Étape 1 · Pour chaque élément, demande-toi : "Puis-je le toucher physiquement ?"
  Étape 2 · OUI → H (Hardware) · NON → S (Software)
  Étape 3 · Traduis chaque élément en français.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Compréhension écrite · Notions travaillées : Lecture, compréhension, réponses en anglais

  Lis ce texte, puis réponds aux questions en anglais avec des phrases complètes.
  Traduis chaque réponse en français.

  TEXTE :
  "Kouamé is a student at Lycée Technique d'Abidjan. His school has just opened
  a new computer room with twenty desktop computers. Each computer has a monitor,
  a keyboard, a mouse and a printer. The computers run on the Windows operating
  system and have Microsoft Word installed for document writing. Kouamé uses the
  computer to type his assignments and search for information online. He also
  uses it to download educational documents. His teacher, Mr. Diallo, says that
  knowing how to use a computer is an essential skill for the future."

  Q1 : Where does Kouamé study?
  Q2 : How many computers does the new computer room have?
  Q3 : Name three hardware components mentioned in the text.
  Q4 : What does Kouamé use the computer for?
  Q5 : According to Mr. Diallo, why is it important to know how to use a computer?

  GUIDE
  Étape 1 · Lis le texte une première fois pour en comprendre le sens général.
  Étape 2 · Lis chaque question et repère dans le texte la partie qui y répond.
  Étape 3 · Écris ta réponse avec une phrase complète en anglais.
  Étape 4 · Traduis ta réponse en français.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Compléter des phrases · Notions travaillées : Structures "used to / consists of / allows"

  Complète chaque phrase avec la structure correcte parmi :
  (is used to · consists of · allows you to · can be used for · is made up of)
  Traduis chaque phrase complétée en français.

  a) A computer _______ process, store and display information.
  b) The hardware of a computer _______ a monitor, a keyboard and a mouse.
  c) A laptop _______ work anywhere, even without a desk.
  d) A printer _______ printing documents and photos.
  e) The CPU _______ several electronic components that work together.

  GUIDE
  Étape 1 · Lis chaque phrase et repère le sens : fonction, composition ou possibilité ?
  Étape 2 · Choisis la structure qui convient.
  Étape 3 · Vérifie que le verbe qui suit est à la bonne forme.
  Étape 4 · Traduis chaque phrase en français.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Production écrite · Notions travaillées : Expression écrite, vocabulaire TIC, structures de la leçon

  Ton lycée vient de recevoir des ordinateurs. Le directeur te demande d'écrire
  un court article (10 à 12 phrases) pour le journal de l'école pour expliquer
  à quoi sert un ordinateur et comment il fonctionne. Tu dois :
    · Définir ce qu'est un ordinateur (1-2 phrases)
    · Expliquer la différence entre hardware et software avec des exemples (3 phrases)
    · Donner 3 utilisations de l'ordinateur avec "is used to" ou "allows you to"
    · Expliquer l'importance de l'ordinateur pour les élèves ivoiriens (2 phrases)
    · Conclure avec une phrase d'encouragement

  GUIDE
  Étape 1 · Commence par une introduction : "Our school has just received new computers..."
  Étape 2 · Définis l'ordinateur.
  Étape 3 · Explique hardware vs software avec des exemples concrets.
  Étape 4 · Utilise "A computer is used to..." et "It allows you to..." pour les fonctions.
  Étape 5 · Parle de l'importance pour les élèves ivoiriens.
  Étape 6 · Termine par une phrase d'encouragement à apprendre l'informatique.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — Describing my ideal computer room
  Durée conseillée : 30 minutes

  Énoncé :
  Ton école n'a pas encore de salle informatique. Tu écris une lettre au directeur
  (12 à 15 phrases) pour lui demander d'en créer une. Dans ta lettre, tu dois :
    · Expliquer ce qu'est un ordinateur et à quoi il sert (2 phrases)
    · Décrire le matériel (hardware) dont la salle aurait besoin (3 phrases)
    · Expliquer les logiciels (software) nécessaires (2 phrases)
    · Donner 3 raisons pour lesquelles les élèves ont besoin d'ordinateurs
    · Utiliser MUST / SHOULD au moins 2 fois chacun
    · Terminer par une formule de politesse

  Contraintes obligatoires :
    ✔ Utiliser "is used to" au moins 2 fois
    ✔ Utiliser "consists of" ou "is made up of" au moins 1 fois
    ✔ Utiliser MUST au moins 2 fois et SHOULD au moins 2 fois
    ✔ Utiliser au moins 10 mots de vocabulaire vus en classe
    ✔ Traduire la lettre complète en français
    ✔ Format lettre : lieu + date / destinataire / corps / formule de politesse / signature

  GUIDE
  Étape 1 · En-tête : Abidjan, [date] / To : The Headmaster of [nom du lycée]
  Étape 2 · Salutation : "Dear Sir / Dear Madam,"
  Étape 3 · Introduction : "I am writing to ask you to create a computer room in our school."
  Étape 4 · Explique ce qu'est un ordinateur et à quoi il sert.
  Étape 5 · Décris le hardware nécessaire avec "The room should consist of..."
  Étape 6 · Mentionne les logiciels avec "The computers must have..."
  Étape 7 · Donne 3 raisons avec "Students need computers because..."
  Étape 8 · Utilise MUST et SHOULD pour les obligations et conseils.
  Étape 9 · Formule de politesse : "Yours faithfully," + prénom + NOM
  Étape 10 · Traduis toute la lettre en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON

  Abidjan, 14th July 2026

  To : The Headmaster of Lycée Moderne de Cocody

  Dear Sir,

  I am writing to ask you to create a computer room in our school,
  because I believe it is essential for our education and our future.
  (Je vous écris pour vous demander de créer une salle informatique dans notre école,
  car je crois que c'est essentiel pour notre éducation et notre avenir.)

  First, let me explain what a computer is. A computer is an electronic machine
  that is used to process, store and display information. It is also used to
  write documents, search for information and communicate with others.
  (Premièrement, laissez-moi expliquer ce qu'est un ordinateur. Un ordinateur
  est une machine électronique utilisée pour traiter, stocker et afficher des
  informations. Il est aussi utilisé pour écrire des documents, chercher des
  informations et communiquer avec d'autres.)

  The computer room should consist of at least twenty desktop computers,
  each with a monitor, a keyboard and a mouse.
  The room must also have printers so that students can print their work.
  Each computer should have speakers and a webcam for multimedia activities.
  (La salle informatique devrait être composée d'au moins vingt ordinateurs
  de bureau, chacun avec un écran, un clavier et une souris.
  La salle doit aussi avoir des imprimantes pour que les élèves puissent
  imprimer leurs travaux. Chaque ordinateur devrait avoir des haut-parleurs
  et une webcam pour les activités multimédia.)

  The computers must have an operating system such as Windows or Linux installed.
  They should also have a word processor like Microsoft Word and a web browser
  like Firefox so that students can write and do research online.
  (Les ordinateurs doivent avoir un système d'exploitation comme Windows ou Linux
  installé. Ils devraient aussi avoir un traitement de texte comme Microsoft Word
  et un navigateur web comme Firefox pour que les élèves puissent écrire et
  faire des recherches en ligne.)

  There are three important reasons why our students need computers.
  First, students need computers because computer literacy is required in most
  modern jobs — without this skill, they will be left behind.
  Second, computers allow students to access educational resources online
  and study more effectively.
  Third, students preparing for the BEPC must develop digital skills
  that are increasingly tested in examinations.
  (Il y a trois raisons importantes pour lesquelles nos élèves ont besoin d'ordinateurs.
  Premièrement, les élèves ont besoin d'ordinateurs car la maîtrise de l'informatique
  est requise dans la plupart des emplois modernes.
  Deuxièmement, les ordinateurs permettent aux élèves d'accéder aux ressources
  éducatives en ligne et d'étudier plus efficacement.
  Troisièmement, les élèves préparant le BEPC doivent développer des compétences
  numériques de plus en plus évaluées aux examens.)

  I strongly believe that our school must invest in technology to prepare
  our students for the future. You should consider this request as a priority
  for the coming school year.
  (Je crois fermement que notre école doit investir dans la technologie pour
  préparer nos élèves pour l'avenir. Vous devriez considérer cette demande
  comme une priorité pour l'année scolaire à venir.)

  Yours faithfully,
  Issa KONÉ
  Élève de 3ème

  ─────────────────────────────────────────────────
  TRADUCTION COMPLÈTE : voir paragraphes en français ci-dessus — inclus intégralement.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1

  Partie A :
    1. keyboard → clavier
    2. software → logiciel
    3. printer → imprimante
    4. to download → télécharger
    5. hard drive → disque dur
    6. folder → dossier
    7. data → données
    8. web browser → navigateur web

  Partie B :
    1. souris → mouse
    2. logiciel → software
    3. sauvegarder → to save
    4. ordinateur portable → laptop computer
    5. système d'exploitation → operating system

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2

  1. S — antivirus (antivirus)
  2. H — monitor (écran / moniteur)
  3. S — operating system (système d'exploitation)
  4. H — webcam (webcam)
  5. S — Microsoft Word (traitement de texte)
  6. H — keyboard (clavier)
  7. S — Firefox (navigateur web)
  8. H — printer (imprimante)
  9. H — CPU (unité centrale)
  10. S — application (application)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3

  Q1 : Where does Kouamé study?
  R1 : Kouamé studies at Lycée Technique d'Abidjan.
       (Kouamé étudie au Lycée Technique d'Abidjan.)

  Q2 : How many computers does the new computer room have?
  R2 : The new computer room has twenty desktop computers.
       (La nouvelle salle informatique a vingt ordinateurs de bureau.)

  Q3 : Name three hardware components mentioned in the text.
  R3 : Three hardware components mentioned in the text are the monitor,
       the keyboard and the mouse.
       (Trois composants matériels mentionnés dans le texte sont l'écran,
       le clavier et la souris.)

  Q4 : What does Kouamé use the computer for?
  R4 : Kouamé uses the computer to type his assignments, search for information
       online and download educational documents.
       (Kouamé utilise l'ordinateur pour taper ses devoirs, chercher des
       informations en ligne et télécharger des documents éducatifs.)

  Q5 : According to Mr. Diallo, why is it important to know how to use a computer?
  R5 : According to Mr. Diallo, knowing how to use a computer is an essential
       skill for the future.
       (Selon M. Diallo, savoir utiliser un ordinateur est une compétence
       essentielle pour l'avenir.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4

  a) A computer IS USED TO process, store and display information.
     (Un ordinateur est utilisé pour traiter, stocker et afficher des informations.)

  b) The hardware of a computer CONSISTS OF a monitor, a keyboard and a mouse.
     (Le matériel d'un ordinateur est composé d'un écran, d'un clavier et d'une souris.)

  c) A laptop ALLOWS YOU TO work anywhere, even without a desk.
     (Un ordinateur portable te permet de travailler n'importe où, même sans bureau.)

  d) A printer CAN BE USED FOR printing documents and photos.
     (Une imprimante peut être utilisée pour imprimer des documents et des photos.)

  e) The CPU IS MADE UP OF several electronic components that work together.
     (L'unité centrale est composée de plusieurs composants électroniques
     qui fonctionnent ensemble.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5

  Exemple d'article pour le journal de l'école :

  "NEW COMPUTERS FOR OUR SCHOOL!
  (DE NOUVEAUX ORDINATEURS POUR NOTRE ÉCOLE !)

  Our school has just received new computers and we are very excited.
  But what exactly is a computer and how does it work?
  (Notre école vient de recevoir de nouveaux ordinateurs et nous sommes
  très enthousiastes. Mais qu'est-ce qu'un ordinateur exactement
  et comment fonctionne-t-il ?)

  A computer is an electronic machine that is used to process, store
  and display information. It is one of the most important inventions
  of the modern world.
  (Un ordinateur est une machine électronique utilisée pour traiter,
  stocker et afficher des informations. C'est l'une des inventions
  les plus importantes du monde moderne.)

  A computer consists of two main parts : hardware and software.
  Hardware includes the physical components you can touch, such as
  the monitor, the keyboard, the mouse and the printer.
  Software includes the programmes you cannot touch, such as Windows,
  Microsoft Word and Firefox.
  (Un ordinateur est composé de deux parties principales : le matériel
  et le logiciel. Le matériel comprend les composants physiques que l'on
  peut toucher, comme l'écran, le clavier, la souris et l'imprimante.
  Le logiciel comprend les programmes que l'on ne peut pas toucher,
  comme Windows, Microsoft Word et Firefox.)

  A computer is used to write documents and do homework.
  It allows you to search for information and download educational resources.
  It also allows you to communicate with teachers and classmates by email.
  (Un ordinateur est utilisé pour écrire des documents et faire des devoirs.
  Il te permet de chercher des informations et de télécharger des ressources
  éducatives. Il te permet aussi de communiquer avec les professeurs
  et les camarades par e-mail.)

  For Ivorian students, learning to use a computer is essential.
  In today's world, most jobs require computer skills, and our school
  is preparing us for that future.
  (Pour les élèves ivoiriens, apprendre à utiliser un ordinateur est essentiel.
  Dans le monde d'aujourd'hui, la plupart des emplois exigent des compétences
  informatiques, et notre école nous prépare à cet avenir.)

  Let us all learn to use these computers well — our success depends on it!
  (Apprenons tous à bien utiliser ces ordinateurs — notre succès en dépend !)"

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "A computer is a bicycle for the mind."
  (Un ordinateur est une bicyclette pour l'esprit.)
  — Steve Jobs

  Prof. Evelyne ASSI · Anglais LV1 · 3ème · DPFC/MENET-FP CI · 2025-2026
"""

# ============================================================
#  GÉNÉRATEUR — VERSION 4 (style modèle de référence v2)
# ============================================================

import re
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── COULEURS ─────────────────────────────────────────────────
C_BLEU_TITRE  = RGBColor(0x18, 0x40, 0x82)
C_OR_TITRE    = RGBColor(0xC9, 0x92, 0x25)
C_GRIS_TEXTE  = RGBColor(0x23, 0x23, 0x23)
C_GRIS_CLAIR  = RGBColor(0x55, 0x55, 0x55)
C_VERT_OK     = RGBColor(0x2A, 0x7E, 0x5B)
C_ROUGE_ERR   = RGBColor(0xA6, 0x2D, 0x2D)
C_ORANGE_ANAL = RGBColor(0xC0, 0x54, 0x1A)
C_BLANC       = RGBColor(0xFF, 0xFF, 0xFF)

# ── FONDS ────────────────────────────────────────────────────
BG_SECTION    = RGBColor(0x18, 0x3E, 0x82)
BG_CAPSULE    = RGBColor(0xF0, 0xEB, 0xF8)
BG_ANCRAGE    = RGBColor(0xE2, 0xF0, 0xD9)
BG_HARKNESS   = RGBColor(0xEA, 0xF2, 0xF8)
BG_DECOUVERTE = RGBColor(0xEA, 0xF2, 0xF8)
BG_RETENIR    = RGBColor(0xE2, 0xF0, 0xD9)
BG_ATTENTION  = RGBColor(0xFF, 0xF2, 0xCC)
BG_DEF        = RGBColor(0xEA, 0xF2, 0xF8)
BG_TRACE      = RGBColor(0xED, 0xF4, 0xFB)
BG_ANALOGIE   = RGBColor(0xFE, 0xF0, 0xE7)
BG_SYNTHESE   = RGBColor(0xD9, 0xE2, 0xF3)
BG_EXO_GUIDE  = RGBColor(0xF5, 0xF7, 0xFA)
BG_EXO_ENT    = RGBColor(0xFA, 0xFA, 0xFA)
BG_DEVOIR     = RGBColor(0xFF, 0xF9, 0xEE)
BG_CORRIGE    = RGBColor(0xF8, 0xFB, 0xF7)
BG_SCHEMA     = RGBColor(0xFA, 0xFA, 0xFA)
BG_SOMMAIRE   = RGBColor(0xEE, 0xF3, 0xFA)


def _hex(c):
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def set_cell_bg(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for e in tcPr.findall(qn('w:shd')):
        tcPr.remove(e)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), _hex(color))
    tcPr.append(shd)


def remove_cell_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "none")
        b.set(qn("w:sz"), "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        tcBorders.append(b)
    tcPr.append(tcBorders)


def add_cell_border_bottom(cell, color="184082", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), color)
    tcBorders.append(bottom)
    tcPr.append(tcBorders)


def add_cell_border_top(cell, color="184082", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), sz)
    top.set(qn("w:space"), "0")
    top.set(qn("w:color"), color)
    tcBorders.append(top)
    tcPr.append(tcBorders)


doc = Document()
for sec in doc.sections:
    sec.top_margin    = Cm(1.8)
    sec.bottom_margin = Cm(1.8)
    sec.left_margin   = Cm(2.0)
    sec.right_margin  = Cm(2.0)


# ── EN-TÊTE ──────────────────────────────────────────────────
def add_header(lecon_info):
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    tbl = header.add_table(rows=1, cols=3, width=Cm(17))
    tbl.style = "Table Grid"
    cells = tbl.rows[0].cells
    for cell in cells:
        remove_cell_borders(cell)
        add_cell_border_bottom(cell)
    p0 = cells[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r0 = p0.add_run("LPA · 3ème")
    r0.font.size = Pt(8); r0.font.color.rgb = C_GRIS_CLAIR
    p1 = cells[1].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(lecon_info)
    r1.bold = True; r1.font.size = Pt(8); r1.font.color.rgb = C_BLEU_TITRE
    p2 = cells[2].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r2 = p2.add_run("Anglais LV1")
    r2.font.size = Pt(8); r2.font.color.rgb = C_GRIS_CLAIR


# ── PIED DE PAGE ─────────────────────────────────────────────
def add_footer(prof, organisme):
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    tbl = footer.add_table(rows=1, cols=3, width=Cm(17))
    tbl.style = "Table Grid"
    cells = tbl.rows[0].cells
    for cell in cells:
        remove_cell_borders(cell)
        add_cell_border_top(cell)
    p0 = cells[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r0 = p0.add_run(prof); r0.font.size = Pt(8); r0.font.color.rgb = C_GRIS_CLAIR
    p1 = cells[1].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(organisme); r1.font.size = Pt(8); r1.font.color.rgb = C_GRIS_CLAIR
    p2 = cells[2].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r2 = p2.add_run("Page "); r2.font.size = Pt(8); r2.font.color.rgb = C_GRIS_CLAIR
    fldChar1 = OxmlElement("w:fldChar"); fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText"); instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar"); fldChar2.set(qn("w:fldCharType"), "end")
    r_p = p2.add_run(); r_p.font.size = Pt(8); r_p.font.color.rgb = C_GRIS_CLAIR
    r_p._r.append(fldChar1); r_p._r.append(instrText); r_p._r.append(fldChar2)


# ── PAGE DE GARDE ────────────────────────────────────────────
def add_page_garde(titre, competence, professeur, organisme, lecon_num):
    # Grand LPA
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(50)
    p.paragraph_format.space_after  = Pt(4)
    r = p.add_run("LPA")
    r.bold = True; r.font.size = Pt(60); r.font.color.rgb = C_BLEU_TITRE
    # Sous-titre
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after  = Pt(16)
    r2 = p2.add_run("Lycée Personnel Autonome")
    r2.italic = True; r2.font.size = Pt(13); r2.font.color.rgb = C_GRIS_CLAIR
    # Ligne bleue
    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.space_before = Pt(0)
    p_sep.paragraph_format.space_after  = Pt(14)
    pPr = p_sep._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '16')
    bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), '184082')
    pBdr.append(bottom); pPr.append(pBdr)
    # Matière · Leçon · Niveau
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after  = Pt(6)
    r3 = p3.add_run(f"Anglais LV1  ·  {lecon_num}  ·  Classe de 3ème")
    r3.bold = True; r3.font.size = Pt(14); r3.font.color.rgb = C_BLEU_TITRE
    # Grand titre — toujours affiché, même si vide on prend l'en-tête
    titre_affiche = titre if titre else lecon_num
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(0)
    p4.paragraph_format.space_after  = Pt(6)
    r4 = p4.add_run(titre_affiche)
    r4.bold = True; r4.font.size = Pt(32); r4.font.color.rgb = C_OR_TITRE
    # Compétence italique
    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.paragraph_format.space_before = Pt(0)
    p5.paragraph_format.space_after  = Pt(30)
    r5 = p5.add_run(competence)
    r5.italic = True; r5.font.size = Pt(12); r5.font.color.rgb = C_GRIS_CLAIR
    # Professeur
    p6 = doc.add_paragraph()
    p6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6.paragraph_format.space_before = Pt(0)
    p6.paragraph_format.space_after  = Pt(6)
    r6 = p6.add_run(professeur)
    r6.font.size = Pt(12); r6.font.color.rgb = C_GRIS_TEXTE
    # Organisme
    p7 = doc.add_paragraph()
    p7.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p7.paragraph_format.space_before = Pt(0)
    p7.paragraph_format.space_after  = Pt(6)
    r7 = p7.add_run(organisme)
    r7.font.size = Pt(11); r7.font.color.rgb = C_GRIS_CLAIR
    # Saut de page
    doc.add_page_break()


# ── SOMMAIRE ─────────────────────────────────────────────────
def add_sommaire(sommaire_lines, sections_phases):
    """
    sommaire_lines : lignes du cours (1. Avant la leçon, etc.)
    sections_phases : dict pour enrichir avec les phases
    """
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, BG_SOMMAIRE)
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top","80"),("bottom","80"),("left","140"),("right","140")]:
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), val); m.set(qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)
    # Titre Sommaire
    p0 = cell.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(8)
    r0 = p0.add_run("Sommaire")
    r0.bold = True; r0.font.size = Pt(16); r0.font.color.rgb = C_BLEU_TITRE

    # Mapping sections → sous-points enrichis
    enriched = {
        "1": ["· Capsule de reprise", "· Ancrage ivoirien", "· Questions Harkness"],
        "2": ["· Phase 1 — Présentation & prérequis", "· Phase 2 — Découverte guidée",
              "· Phase 3 — Schémas textuels", "· Phase 4 — Définitions DPFC",
              "· Phase 5 — Trace écrite + Analogies CI", "· Phase 5b — Synthèse",
              "· Phase 6 — Exercices guidés"],
        "3": ["· Exercices 1 à 5 + Devoir maison"],
        "4": ["· Corrigé du devoir maison + Exercices 1 à 5"],
    }

    num_auto = 0  # compteur si pas de numéro explicite
    for line in sommaire_lines:
        line = line.strip()
        if not line:
            continue
        # Nouveau format : "Section N — ..."
        m2 = re.match(r'^Section\s+(\d+)\s*[—–-]\s*(.*)', line, re.IGNORECASE)
        if m2:
            num = m2.group(1); label = m2.group(2).strip()
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(2)
            r = p.add_run(f"Section {num} — {label}")
            r.bold = True; r.font.size = Pt(11); r.font.color.rgb = C_BLEU_TITRE
            for sub in enriched.get(num, []):
                ps = cell.add_paragraph()
                ps.paragraph_format.space_before = Pt(0)
                ps.paragraph_format.space_after  = Pt(1)
                ps.paragraph_format.left_indent  = Pt(14)
                rs = ps.add_run(sub)
                rs.font.size = Pt(10); rs.font.color.rgb = C_GRIS_CLAIR
            continue
        # Sous-point du nouveau format : "    · Phase..."
        if re.match(r'^[·•]\s+', line):
            ps = cell.add_paragraph()
            ps.paragraph_format.space_before = Pt(0)
            ps.paragraph_format.space_after  = Pt(1)
            ps.paragraph_format.left_indent  = Pt(14)
            rs = ps.add_run(line)
            rs.font.size = Pt(10); rs.font.color.rgb = C_GRIS_CLAIR
            continue
        m = re.match(r'^(\d+)\.\s*(.*)', line)
        if m:
            num = m.group(1); label = m.group(2)
            # Titre de section
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(2)
            r = p.add_run(f"Section {num} — {label}")
            r.bold = True; r.font.size = Pt(11); r.font.color.rgb = C_BLEU_TITRE
            # Sous-points
            for sub in enriched.get(num, []):
                ps = cell.add_paragraph()
                ps.paragraph_format.space_before = Pt(0)
                ps.paragraph_format.space_after  = Pt(1)
                ps.paragraph_format.left_indent  = Pt(14)
                rs = ps.add_run(sub)
                rs.font.size = Pt(10); rs.font.color.rgb = C_GRIS_CLAIR

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    doc.add_page_break()


# ── HELPERS ───────────────────────────────────────────────────
def _run(para, text, bold=False, italic=False, color=None, size=10.5):
    r = para.add_run(text)
    r.bold = bold; r.italic = italic
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return r


def make_box(bg_color):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, bg_color)
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top","70"),("bottom","70"),("left","120"),("right","120")]:
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), val); m.set(qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)
    return tbl, cell


def space_after_box(pts=4):
    doc.add_paragraph().paragraph_format.space_after = Pt(pts)


def _add_para_in_cell(cell, text, bold=False, italic=False,
                       color=None, size=10.5, sb=0, sa=2,
                       align=WD_ALIGN_PARAGRAPH.LEFT, indent_pt=0):
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if indent_pt:
        p.paragraph_format.left_indent = Pt(indent_pt)
    _run(p, text, bold=bold, italic=italic,
         color=color or C_GRIS_TEXTE, size=size)
    return p


# ── BANNIÈRE SECTION ─────────────────────────────────────────
def add_banner(text):
    tbl, cell = make_box(BG_SECTION)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    _run(p, text, bold=True, color=C_BLANC, size=13)
    space_after_box(5)


# ── PHASE HEADER ─────────────────────────────────────────────
def add_phase(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(2)
    # Ligne de soulignement
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '2')
    bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), 'D0D8E8')
    pBdr.append(bottom); pPr.append(pBdr)
    _run(p, text, bold=True, color=C_BLEU_TITRE, size=14)


# ── SOUS-TITRE ───────────────────────────────────────────────
def add_sous_titre(text, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after  = Pt(2)
    _run(p, text, bold=True, color=color or C_OR_TITRE, size=11.5)


# ── TEXTE COURANT ────────────────────────────────────────────
def add_text(text, italic=False, color=None, size=11, sb=0, sa=3,
             align=WD_ALIGN_PARAGRAPH.LEFT, indent_pt=0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if indent_pt:
        p.paragraph_format.left_indent = Pt(indent_pt)
    _run(p, text, italic=italic, color=color or C_GRIS_TEXTE, size=size)


# ── BOÎTE GÉNÉRIQUE ──────────────────────────────────────────
def add_simple_box(lines, bg, title=None, title_color=None,
                   title_size=10, title_icon=""):
    tbl, cell = make_box(bg)
    first = True
    if title:
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(5)
        txt = (title_icon + "  " + title) if title_icon else title
        _run(p, txt, bold=True, color=title_color or C_BLEU_TITRE, size=title_size)
        first = False
    for line in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE HARKNESS enrichie ──────────────────────────────────
def add_harkness_box(title, lines):
    tbl, cell = make_box(BG_HARKNESS)
    # Titre
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(5)
    _run(p, "▶  " + title, bold=True, color=C_BLEU_TITRE, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if line.startswith("Q :") or line.startswith("Q:"):
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        elif line.startswith("Guide") or line.startswith("Piste"):
            _run(p, line, italic=True, color=C_GRIS_CLAIR, size=10.5)
        elif line.startswith("Corrigé") or line.startswith("Corrige"):
            _run(p, line, bold=True, color=C_VERT_OK, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE TEXTE DE DÉCOUVERTE ────────────────────────────────
def add_decouverte_box(eng_lines, trad_lines):
    tbl, cell = make_box(BG_DECOUVERTE)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "📖  Texte de découverte", bold=True, color=C_BLEU_TITRE, size=10)
    for line in eng_lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    if trad_lines:
        # Ligne séparatrice visuelle
        p_sep = cell.add_paragraph()
        p_sep.paragraph_format.space_before = Pt(6)
        p_sep.paragraph_format.space_after  = Pt(4)
        pPr = p_sep._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        top_b = OxmlElement('w:top')
        top_b.set(qn('w:val'), 'single'); top_b.set(qn('w:sz'), '2')
        top_b.set(qn('w:space'), '1'); top_b.set(qn('w:color'), 'B0C4D8')
        pBdr.append(top_b); pPr.append(pBdr)
        _run(p_sep, "— Traduction :", bold=True, color=C_BLEU_TITRE, size=10)
        for line in trad_lines:
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(2)
            _run(p, line, italic=True, color=C_GRIS_CLAIR, size=10.5)
    space_after_box()


# ── BOÎTE À RETENIR ──────────────────────────────────────────
def add_retenir_box(title, lines):
    tbl, cell = make_box(BG_RETENIR)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "📘  " + title, bold=True, color=C_VERT_OK, size=10)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE ATTENTION ──────────────────────────────────────────
def add_attention_box(title, lines):
    tbl, cell = make_box(BG_ATTENTION)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "⚠  " + title, bold=True, color=C_OR_TITRE, size=10)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        # ✗ en rouge, ✓ en vert
        if '✗' in line or '✓' in line:
            parts = re.split(r'(→)', line)
            for part in parts:
                part = part.strip()
                if not part:
                    continue
                if '✗' in part:
                    _run(p, part + "  ", color=C_ROUGE_ERR, size=10.5)
                elif '✓' in part:
                    _run(p, part, bold=True, color=C_VERT_OK, size=10.5)
                elif part == '→':
                    _run(p, "  →  ", italic=True, color=C_GRIS_CLAIR, size=10.5)
                else:
                    _run(p, part, color=C_GRIS_TEXTE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE DÉFINITION ─────────────────────────────────────────
def add_def_box(title, lines):
    tbl, cell = make_box(BG_DEF)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, title, bold=True, color=C_BLEU_TITRE, size=11)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE TRACE ÉCRITE ───────────────────────────────────────
def add_trace_box(lines):
    tbl, cell = make_box(BG_TRACE)
    first = True
    for line in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^(I|II|III|IV|V)\.\s', line):
            _run(p, line, bold=True, color=C_BLEU_TITRE, size=11)
        elif re.match(r'^✎', line):
            _run(p, line, bold=True, color=C_BLEU_TITRE, size=10)
        elif re.match(r'^⚠', line):
            _run(p, line, bold=True, color=C_OR_TITRE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE ANALOGIE ───────────────────────────────────────────
def add_analogie_box(title, lines):
    tbl, cell = make_box(BG_ANALOGIE)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "💡  " + title, bold=True, color=C_ORANGE_ANAL, size=10)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE SYNTHÈSE ───────────────────────────────────────────
def add_synthese_box(lines):
    tbl, cell = make_box(BG_SYNTHESE)
    first = True
    for line in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^5 points|^5 mots|^3 erreurs', line):
            _run(p, line, bold=True, color=C_BLEU_TITRE, size=11)
        elif re.match(r'^\s*\d+\.', line):
            m = re.match(r'^(\s*\d+\.\s*)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_BLEU_TITRE, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^\s*·', line):
            m = re.match(r'^(\s*·\s*)(\S+)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_OR_TITRE, size=10.5)
                _run(p, m.group(2), bold=True, color=C_GRIS_TEXTE, size=10.5)
                _run(p, m.group(3), color=C_GRIS_CLAIR, size=10.5)
        elif re.match(r'^\s*✗', line):
            m = re.match(r'^(\s*✗\s*)(.*?)\s*→\s*(✓.*)', line)
            if m:
                _run(p, m.group(1) + m.group(2) + "  ", bold=True, color=C_ROUGE_ERR, size=10.5)
                _run(p, "→  ", italic=True, color=C_GRIS_CLAIR, size=10.5)
                _run(p, m.group(3), bold=True, color=C_VERT_OK, size=10.5)
            else:
                _run(p, line, color=C_ROUGE_ERR, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE EXERCICE GUIDÉ ─────────────────────────────────────
def add_exo_guide_box(title, lines):
    tbl, cell = make_box(BG_EXO_GUIDE)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "◉  " + title, bold=True, color=C_BLEU_TITRE, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^Notion ciblée\s*:', line):
            m = re.match(r'^(Notion ciblée\s*:\s*)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_OR_TITRE, size=10.5)
                _run(p, m.group(2), italic=True, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^(Énoncé|MÉTHODE|Démarche)', line):
            _run(p, line, bold=True, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^Étape \d+', line):
            m = re.match(r'^(Étape \d+ · )(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_BLEU_TITRE, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
            else:
                _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^✔', line):
            m = re.match(r'^(✔[^:]*:\s*)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_VERT_OK, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
            else:
                _run(p, line, bold=True, color=C_VERT_OK, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE EXERCICE ENTRAÎNEMENT ──────────────────────────────
def add_exo_ent_box(title, lines):
    tbl, cell = make_box(BG_EXO_ENT)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "◎  " + title, bold=True, color=C_BLEU_TITRE, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^Notions travaillées\s*:', line):
            m = re.match(r'^(Notions travaillées\s*:\s*)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_OR_TITRE, size=10.5)
                _run(p, m.group(2), italic=True, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^GUIDE|^Guide de réussite', line):
            _run(p, line, bold=True, color=C_OR_TITRE, size=10.5)
        elif re.match(r'^Étape \d+', line):
            m = re.match(r'^(Étape \d+ · )(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_BLEU_TITRE, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
            else:
                _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE DEVOIR MAISON ──────────────────────────────────────
def add_devoir_box(title, lines):
    tbl, cell = make_box(BG_DEVOIR)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "◈  " + title, bold=True, color=C_OR_TITRE, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^(CONTEXTE|ÉNONCÉ|ENONCE|STRUCTURE|CONTRAINTES|GUIDE PAS)', line):
            _run(p, line, bold=True, color=C_OR_TITRE, size=10.5)
        elif re.match(r'^Durée', line):
            _run(p, line, italic=True, color=C_GRIS_CLAIR, size=10)
        elif re.match(r'^Étape \d+', line):
            m = re.match(r'^(Étape \d+ · )(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_OR_TITRE, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
            else:
                _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        elif line.startswith('·'):
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^\d+\.', line):
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE CORRIGÉ ────────────────────────────────────────────
def add_corrige_box(title, lines):
    tbl, cell = make_box(BG_CORRIGE)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "✔  " + title, bold=True, color=C_VERT_OK, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^✔', line):
            _run(p, line, bold=True, color=C_VERT_OK, size=10.5)
        elif re.match(r'^TRADUCTION|^TITRE PROPOSÉ|^TEXTE PROPOSÉ', line):
            _run(p, line, bold=True, color=C_GRIS_TEXTE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── TABLEAU STRUCTURES FONDAMENTALES ────────────────────────
def add_tableau_box(title, md_lines):
    """Transforme les lignes markdown | col | col | en vrai tableau Word."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as _qn

    # Parser les lignes markdown en liste de colonnes
    rows_data = []
    for line in md_lines:
        line = line.strip().strip('|')
        cols = [c.strip() for c in line.split('|')]
        if cols:
            rows_data.append(cols)

    if not rows_data:
        return

    # Titre optionnel au-dessus
    if title:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(6)
        p_t.paragraph_format.space_after  = Pt(3)
        _run(p_t, title, bold=True, color=C_BLEU_TITRE, size=11)

    nb_cols = max(len(r) for r in rows_data)
    tbl = doc.add_table(rows=len(rows_data), cols=nb_cols)
    tbl.style = "Table Grid"

    for r_idx, row_data in enumerate(rows_data):
        is_header = (r_idx == 0)
        row = tbl.rows[r_idx]
        for c_idx in range(nb_cols):
            cell = row.cells[c_idx]
            text = row_data[c_idx] if c_idx < len(row_data) else ""
            # Fond
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(_qn("w:val"), "clear")
            shd.set(_qn("w:color"), "auto")
            shd.set(_qn("w:fill"), "183E82" if is_header else ("F0F4FA" if r_idx % 2 == 0 else "FFFFFF"))
            tcPr.append(shd)
            # Texte
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            r = p.add_run(text)
            r.bold = is_header
            r.font.size = Pt(9.5)
            r.font.color.rgb = C_BLANC if is_header else C_GRIS_TEXTE

    doc.add_paragraph().paragraph_format.space_after = Pt(5)


# ── CITATION FINALE ──────────────────────────────────────────
def add_citation(text_en, text_fr, prof_org):
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(4)
    p1.paragraph_format.space_after  = Pt(2)
    _run(p1, f'"{text_en}"', italic=True, color=C_BLEU_TITRE, size=11)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after  = Pt(2)
    _run(p2, f'({text_fr})', italic=True, color=C_GRIS_CLAIR, size=10)
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(2)
    p3.paragraph_format.space_after  = Pt(0)
    _run(p3, prof_org, color=C_GRIS_CLAIR, size=9)


# ============================================================
#  PARSEUR PRINCIPAL
# ============================================================

lines     = COURS.strip().split('\n')
total     = len(lines)
meta      = {}

# ── Extraction en-tête ════...════ ──────────────────────────
header_lines = []
in_header = False; header_done = False; body_start = 0
for idx, ln in enumerate(lines):
    s = ln.strip()
    if re.match(r'^[═]{4,}', s) and not in_header and not header_done:
        in_header = True; continue
    if re.match(r'^[═]{4,}', s) and in_header:
        in_header = False; header_done = True; body_start = idx + 1; break
    if in_header and s:
        header_lines.append(s)

# ── Méta-lignes ─────────────────────────────────────────────
i = body_start
while i < total:
    s = lines[i].strip()
    m = re.match(r'^(TITRE_LECON|COMPETENCE|PROFESSEUR|ORGANISME)\s*:\s*(.+)', s)
    if m:
        meta[m.group(1)] = m.group(2).strip(); i += 1
    elif not s or s == "SOMMAIRE":
        i += 1
        if s == "SOMMAIRE": break
    else:
        break

titre_lecon = meta.get("TITRE_LECON", "")
competence  = meta.get("COMPETENCE",  "")
professeur  = meta.get("PROFESSEUR",  "")
organisme   = meta.get("ORGANISME",   "")

# Si TITRE_LECON absent : extraire depuis le sommaire (ligne "Section 2")
# ou depuis l'en-tête ════ après le "—"
if not titre_lecon and header_lines:
    # Chercher "Leçon N — TITRE" dans la première ligne d'en-tête
    m_titre = re.search(r'[—–-]\s*(.+)$', header_lines[0])
    if m_titre:
        titre_lecon = m_titre.group(1).strip()

# Si PROFESSEUR absent : chercher dans l'en-tête
if not professeur and header_lines:
    for hl in header_lines:
        m_p = re.search(r'Professeur\s*:\s*(.+)', hl)
        if m_p:
            professeur = m_p.group(1).strip(); break

# Si COMPETENCE absent : chercher dans l'en-tête
if not competence and header_lines:
    for hl in header_lines:
        m_c = re.search(r'Compétence\s*:\s*(.+)', hl)
        if m_c:
            competence = m_c.group(1).strip(); break

# Si ORGANISME absent : chercher dans l'en-tête
if not organisme and header_lines:
    for hl in header_lines:
        m_o = re.search(r'(DPFC/.+)', hl)
        if m_o:
            organisme = m_o.group(1).strip(); break

# Numéro de leçon depuis l'en-tête
lecon_num = "Leçon"
if header_lines:
    m_lec = re.search(r'(Leçon\s*\d+)', header_lines[0])
    if m_lec:
        lecon_num = m_lec.group(1)

# Sommaire
sommaire_lines = []
while i < total:
    s = lines[i].strip()
    if not s or re.match(r'^[─═]{4,}', s): break
    sommaire_lines.append(s); i += 1

# ── En-tête / pied / page de garde / sommaire ───────────────
lecon_info = header_lines[0] if header_lines else "Anglais LV1"
add_header(lecon_info)
add_footer(professeur, organisme)
add_page_garde(titre_lecon, competence, professeur, organisme, lecon_num)
add_sommaire(sommaire_lines, {})

# ── Variables d'état ────────────────────────────────────────
cur_type  = None
cur_title = ""
cur_lines = []
in_tilde  = False
tilde_lines = []
in_syn    = False
syn_lines = []
in_trad_libre  = False
trad_libre_lines = []

SOUS_TITRES = re.compile(
    r'^(Objectifs|Prérequis|MÉTHODE|GUIDE|CONTRAINTES|CONTEXTE|'
    r'CONSIGNE|STRUCTURE|GUIDE PAS|Énoncé)')


def flush():
    global cur_type, cur_title, cur_lines
    if not cur_type:
        return
    t, ti, ls = cur_type, cur_title, cur_lines

    if   t == "CAPSULE":   add_simple_box(ls, BG_CAPSULE, title=ti,
                               title_color=C_BLEU_TITRE, title_size=10, title_icon="📌")
    elif t == "ANCRAGE":   add_simple_box(ls, BG_ANCRAGE, title=ti,
                               title_color=C_BLEU_TITRE, title_size=10, title_icon="🌍")
    elif t == "HARKNESS":  add_harkness_box(ti, ls)
    elif t == "RETENIR":   add_retenir_box(ti, ls)
    elif t == "ATTENTION": add_attention_box(ti, ls)
    elif t == "DEF":       add_def_box(ti, ls)
    elif t == "SCHEMA":    add_simple_box(ls, BG_SCHEMA, title=ti,
                               title_color=C_BLEU_TITRE, title_size=10, title_icon="📊")
    elif t == "TRACE":     add_trace_box(ls)
    elif t == "ANALOGIE":  add_analogie_box(ti, ls)
    elif t == "DEVOIR":    add_devoir_box(ti, ls)
    elif t == "EXOGUIDE":  add_exo_guide_box(ti, ls)
    elif t == "EXOENT":    add_exo_ent_box(ti, ls)
    elif t == "CORRIGE":   add_corrige_box(ti, ls)
    elif t == "TABLEAU":   add_tableau_box(ti, ls)

    cur_type = None; cur_title = ""; cur_lines = []


# ── Citation finale : stockage ───────────────────────────────
citation_en = "Speak even if it's not perfect — silence teaches nothing."
citation_fr = "Parle même si c'est imparfait — le silence n'apprend rien."
citation_sig = ""

# ── Boucle principale ────────────────────────────────────────
while i < total:
    raw = lines[i]; s = raw.strip(); i += 1
    if not s:
        continue

    # Méta déjà traitées
    if re.match(r'^(TITRE_LECON|COMPETENCE|PROFESSEUR|ORGANISME)\s*:', s):
        continue

    # Bloc ~~~ texte de découverte
    if s == "~~~":
        if not in_tilde:
            in_tilde = True; tilde_lines = []
        else:
            in_tilde = False
            flush()
            eng_lines = []; trad_lines = []; in_trad = False
            for tl in tilde_lines:
                if tl == "---TRADUCTION---":
                    in_trad = True; continue
                (trad_lines if in_trad else eng_lines).append(tl)
            # Si la traduction n'était pas dans les ~~~,
            # on active le mode "capture traduction libre"
            add_decouverte_box(eng_lines, trad_lines)
            if not trad_lines:
                in_trad_libre = True
                trad_libre_lines = []
        continue
    if in_tilde:
        tilde_lines.append(s); continue

    # Capture traduction libre après ~~~ sans ---TRADUCTION---
    if in_trad_libre:
        # On arrête à la prochaine boîte ou séparateur
        if (re.match(r'^[┌⚠◆◉◎◈✔✎~▶►\[]', s) or
                re.match(r'^(SECTION|Phase|SYNTHESE)', s) or
                re.match(r'^[─═]{4,}', s)):
            # Injecter les lignes de traduction dans la dernière boîte découverte
            # (on les ajoute via add_text en style traduit)
            in_trad_libre = False
            # Relancer le traitement de cette ligne (ne pas ignorer)
            # on laisse tomber dans la suite du parseur
        else:
            # Ignorer "TRADUCTION COMPLÈTE :" et absorber les lignes
            if not re.match(r'^TRADUCTION', s):
                trad_libre_lines.append(s)
            continue

    # SYNTHESE — avec balises explicites OU détection automatique
    if s == "SYNTHESE_DEBUT" or re.match(r'^5 points essentiels', s):
        flush(); in_syn = True; syn_lines = []
        if s != "SYNTHESE_DEBUT":
            syn_lines.append(s)  # inclure la ligne déclencheuse
        continue
    if s == "SYNTHESE_FIN":
        in_syn = False; add_synthese_box(syn_lines); syn_lines = []; continue
    if in_syn:
        # Fin automatique de synthèse à la prochaine Phase ou Section
        if re.match(r'^(Phase \d|SECTION)', s):
            in_syn = False; add_synthese_box(syn_lines); syn_lines = []
            # Ne pas ignorer cette ligne, la traiter ci-dessous
        else:
            syn_lines.append(s); continue

    # Séparateurs
    if re.match(r'^[═]{4,}', s) or re.match(r'^[─]{4,}', s):
        flush(); continue

    # SECTION
    if re.match(r'^SECTION \d+\s*[—–-]', s):
        flush(); add_banner(s); continue

    # Phase
    if re.match(r'^Phase \d+[a-z]?\s*[·.]', s):
        flush(); add_phase(s); continue

    # [BLOC TITRE] avec crochets
    m = re.match(r'^\[(.+)\]$', s)
    if m:
        flush(); bl = m.group(1); tu = bl.upper()
        if "CAPSULE" in tu:
            cur_type = "CAPSULE";  cur_title = bl
        elif "ANCRAGE" in tu:
            cur_type = "ANCRAGE";  cur_title = bl
        elif "SCHEMA" in tu or "SCHÉMA" in tu:
            cur_type = "SCHEMA";   cur_title = bl
        else:
            cur_type = "SCHEMA";   cur_title = bl
        continue

    # CAPSULE sans crochets : "CAPSULE DE REPRISE — ..."
    if re.match(r'^CAPSULE\s+DE\s+REPRISE', s):
        flush()
        cur_type = "CAPSULE"; cur_title = s.strip(); cur_lines = []
        continue

    # ANCRAGE sans crochets : "ANCRAGE IVOIRIEN"
    if re.match(r'^ANCRAGE\s+IVOIRIEN', s):
        flush()
        cur_type = "ANCRAGE"; cur_title = "ANCRAGE IVOIRIEN"; cur_lines = []
        continue

    # Titre de tableau juste avant les | (ex: "TABLEAU DES STRUCTURES...")
    if re.match(r'^TABLEAU\s+DES\s+', s):
        flush()
        cur_type = "TABLEAU"; cur_title = s.strip(); cur_lines = []
        continue

    # TABLEAU MARKDOWN — lignes | col | col |
    if re.match(r'^\|', s) and s.count('|') >= 2:
        if cur_type != "TABLEAU":
            flush()
            cur_type = "TABLEAU"; cur_title = ""; cur_lines = []
        # Ignorer lignes séparatrices |---|---|
        if not re.match(r'^\|[-:\s|]+\|', s):
            cur_lines.append(s)
        continue

    # Sous-sections Phase 2 : "A. TITRE", "B. TITRE" etc.
    if re.match(r'^[A-Z]\.\s+[A-ZÀÉÈÊÙÛ\(]', s):
        flush()
        add_sous_titre(s, color=C_BLEU_TITRE)
        continue

    # HARKNESS
    if re.match(r'^[▶►]\s*HARKNESS', s):
        flush()
        cur_type  = "HARKNESS"
        cur_title = s.lstrip("▶► ").strip()
        continue

    # À RETENIR — bordures ┌┐└┘ OU bloc FORMATION ... sans bordures
    if re.match(r'^[┌].*À RETENIR', s):
        flush()
        tc = re.sub(r'^[┌─\s]+', '', s).rstrip('─┐').strip()
        cur_type = "RETENIR"; cur_title = tc; continue
    if re.match(r'^[└]', s):
        flush(); continue
    # Bloc de règles sans bordure (ex: "FORMATION DU PRESENT CONTINUOUS")
    if re.match(r'^FORMATION\s+DU\s+', s):
        flush()
        cur_type  = "RETENIR"
        cur_title = s.strip()
        cur_lines = []
        continue

    # ATTENTION
    if re.match(r'^⚠', s):
        flush()
        cur_type  = "ATTENTION"
        cur_title = s.lstrip("⚠ ").strip()
        continue

    # TRACE ÉCRITE
    if re.match(r'^✎', s):
        flush(); cur_type = "TRACE"; cur_lines = [s]; continue

    # ANALOGIE — flush() systématique (gère 1 ou 2 analogies)
    if re.match(r'^~\s*ANALOGIE', s):
        flush()   # flush TRACE ou ANALOGIE précédente
        cur_type  = "ANALOGIE"
        cur_title = s.lstrip("~ ").strip()
        cur_lines = []
        continue

    # DÉFINITION DPFC
    if re.match(r'^◆\s*DÉFINITION', s):
        flush(); cur_type = "DEF"; cur_title = s.lstrip("◆ ").strip(); continue

    # EXERCICE GUIDÉ
    if re.match(r'^◉', s):
        flush(); cur_type = "EXOGUIDE"; cur_title = s.lstrip("◉ ").strip(); continue

    # EXERCICE ENTRAÎNEMENT
    if re.match(r'^◎', s):
        flush(); cur_type = "EXOENT"; cur_title = s.lstrip("◎ ").strip(); continue

    # DEVOIR MAISON
    if re.match(r'^◈', s):
        flush(); cur_type = "DEVOIR"; cur_title = s.lstrip("◈ ").strip(); continue

    # CORRIGÉ
    if re.match(r'^✔\s*CORRIGÉ', s):
        flush(); cur_type = "CORRIGE"; cur_title = s.lstrip("✔ ").strip(); continue

    # Bloc ════ final (pied de cours) → on arrête le contenu
    if re.match(r'^[═]{4,}', s) and i > body_start + 10:
        flush(); continue

    # Balise "CITATION FINALE" → déclenche le mode citation
    if re.match(r'^CITATION\s+FINALE', s):
        flush(); continue

    # Citation entre guillemets (n'importe laquelle, pas forcément la citation LPA)
    if re.match(r'^["""«]', s) and not cur_type:
        flush()
        m_en = re.search(r'["""«]([^"""»]+)["""»]', s)
        if m_en:
            citation_en = m_en.group(1).strip()
        continue
    # Traduction de la citation entre parenthèses
    if re.match(r'^\(', s) and citation_en and not cur_type:
        m_fr = re.search(r'\(([^)]+)\)', s)
        if m_fr:
            citation_fr = m_fr.group(1).strip()
        continue
    # Auteur de la citation "— Prénom Nom"
    if re.match(r'^[—–-]\s+\w', s) and citation_en and not cur_type:
        citation_sig = s.lstrip("—– ").strip()
        continue

    # Mots-clés LPA par défaut
    if "silence teaches nothing" in s:
        flush()
        m_en = re.search(r'"([^"]+)"', s)
        if m_en:
            citation_en = m_en.group(1)
        continue
    if "silence n'apprend rien" in s:
        m_fr = re.search(r'\(([^)]+)\)', s)
        if m_fr:
            citation_fr = m_fr.group(1)
        continue
    # Ligne signature finale Prof. … · DPFC/…
    if re.match(r'^Prof\.?\s+\w+', s) and ('DPFC' in s or 'Anglais' in s or '3ème' in s):
        citation_sig = s
        continue

    # Contenu d'un bloc ouvert
    if cur_type is not None:
        cur_lines.append(s)
    else:
        # Hors bloc
        if SOUS_TITRES.match(s):
            add_sous_titre(s)
        elif re.match(r'^(Titre\s*:|Compétence\s*:|Professeur\s+\w)', s):
            # Lignes de contexte Phase 1 — texte bleu gris léger
            add_text(s, color=C_GRIS_CLAIR, size=11, italic=True)
        elif re.match(r'^LIS CE TEXTE', s):
            # Instruction avant le texte de découverte — texte courant
            add_text(s, color=C_GRIS_TEXTE, size=11, sb=4)
        elif re.match(r'^TRADUCTION', s):
            # Entête "TRADUCTION COMPLÈTE :" hors boîte → ignorer
            pass
        else:
            add_text(s, color=C_GRIS_TEXTE, size=11)

flush()

# Citation finale
add_citation(citation_en, citation_fr,
             citation_sig or f"{professeur}  ·  {organisme}")

# ============================================================
#  SAUVEGARDE
# ============================================================
nom = "cours_lpa.docx"
if header_lines:
    nom = re.sub(r'[^\w\s·—-]', '', header_lines[0])
    nom = re.sub(r'[\s·—]+', '_', nom).strip('_') + ".docx"

doc.save(nom)
print(f"✅ Document généré : {nom}")
try:
    from google.colab import files
    files.download(nom)
except ImportError:
    print("(Hors Colab — fichier sauvegardé localement)")
