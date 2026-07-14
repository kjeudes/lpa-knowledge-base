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

COURS = """✅ Anglais LV1 · Leçon 24 — The Internet · en cours de génération...

---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 24 · Classe de 3ème

THE INTERNET
*L'Internet*

Compétence C8 · TIC / ICT

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 24 — THE INTERNET
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C8 · TIC / ICT
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 23 : The telephone)
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

CAPSULE DE REPRISE — Leçon 23 : The telephone (Le téléphone)

5 points clés à retenir :
  1. Il existe deux types de téléphones : le fixe (landline) et le mobile (smartphone).
  2. Les expressions clés d'un appel : "Hello, this is... / Can I speak to... /
     Hold on / I'll call back later."
  3. "to hang up" = raccrocher · "to pick up" = décrocher · "to dial" = composer.
  4. Le téléphone a des avantages (communication rapide, Mobile Money, urgences)
     et des inconvénients (distraction, dépendance, arnaques).
  5. En Côte d'Ivoire, Orange Money et MTN MoMo sont des révolutions financières
     accessibles via le téléphone mobile.

3 questions flash :
  Q1 : How do you answer the phone in English?
  R1 : You say "Hello?" or "Hello, this is [name] speaking."
       (Tu dis "Allô ?" ou "Allô, c'est [prénom] à l'appareil.")

  Q2 : What does "the line is busy" mean?
  R2 : It means the person you are calling is already on the phone.
       (Ça veut dire que la personne que tu appelles est déjà au téléphone.)

  Q3 : Give one advantage and one disadvantage of the mobile phone.
  R3 : Advantage : it allows fast communication with family and friends.
       Disadvantage : it can distract students during class.
       (Avantage : il permet une communication rapide avec la famille et les amis.
       Inconvénient : il peut distraire les élèves pendant les cours.)

Lien avec la leçon du jour :
  La leçon 23 nous a montré comment le téléphone permet de communiquer à distance.
  Aujourd'hui, nous allons étudier le réseau qui connecte tous ces appareils
  et révolutionne l'accès à l'information dans le monde entier : L'INTERNET.
  C'est la dernière leçon de la compétence C8 et la dernière leçon
  de votre programme d'anglais de 3ème. Nous allons terminer en beauté !

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

Adjoua est en 3ème au Lycée Classique d'Abidjan.
Pour préparer son exposé sur la Côte d'Ivoire, elle utilise le WiFi
de son lycée pour se connecter à internet et faire des recherches sur Google.
Elle trouve des articles, des photos et même des vidéos sur YouTube.
Elle utilise aussi WhatsApp pour partager les informations trouvées
avec ses camarades Kofi et Bintou via un groupe créé par leur professeur.
Son grand-père à Man ne comprend pas internet. Il dit que c'est "comme
une grande bibliothèque invisible dans le ciel".
Adjoua lui explique que c'est un réseau mondial d'ordinateurs et de téléphones
tous connectés entre eux.

→ Qu'est-ce qu'internet exactement ?
→ À quoi sert internet ?
→ Comment utiliser internet de façon sûre et responsable ?
→ C'est exactement ce que nous allons apprendre aujourd'hui.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — What is the Internet?
  Q : What is the Internet and how does it work?
      (Qu'est-ce qu'Internet et comment fonctionne-t-il ?)
  Guide :
    · Pense à la définition : réseau mondial d'ordinateurs connectés.
    · Pense aux éléments nécessaires pour se connecter : appareil, réseau, fournisseur.
    · Pense à la différence entre Internet et le Web (World Wide Web).
  Corrigé :
    The Internet is a global network of millions of computers, smartphones
    and other devices connected to each other. It allows people to share
    and access information from anywhere in the world.
    To connect to the Internet, you need a device (computer, smartphone or tablet),
    a network connection (WiFi, mobile data or cable) and an Internet Service
    Provider (ISP) such as Orange CI or MTN in Côte d'Ivoire.
    It is important to distinguish the Internet from the World Wide Web (WWW).
    The Internet is the physical network — the cables, satellites and connections.
    The Web is the collection of websites and pages that you access through
    a browser like Chrome or Firefox. The Web uses the Internet, but the Internet
    is much more than just the Web — it also includes emails, online games
    and video calls.
    (Internet est un réseau mondial de millions d'ordinateurs, de smartphones
    et d'autres appareils connectés entre eux. Il permet aux gens de partager
    et d'accéder à des informations de n'importe où dans le monde.
    Pour se connecter à Internet, on a besoin d'un appareil, d'une connexion réseau
    et d'un fournisseur d'accès à Internet comme Orange CI ou MTN en Côte d'Ivoire.
    Il est important de distinguer Internet du World Wide Web. Internet est le réseau
    physique — les câbles, les satellites et les connexions. Le Web est l'ensemble
    des sites web accessibles via un navigateur. Le Web utilise Internet, mais Internet
    est bien plus que le Web — il comprend aussi les e-mails, les jeux en ligne
    et les appels vidéo.)

▶ HARKNESS 2 — Uses of the Internet
  Q : What are the main uses of the Internet today?
      (Quels sont les principaux usages d'Internet aujourd'hui ?)
  Guide :
    · Pense aux usages d'information : recherche, actualités, encyclopédies.
    · Pense aux usages de communication : e-mail, réseaux sociaux, appels vidéo.
    · Pense aux usages éducatifs et économiques en Côte d'Ivoire.
  Corrigé :
    The Internet has many uses today. For information, people use it to search
    for news, read articles and watch educational videos on platforms like YouTube.
    For communication, the Internet allows people to send emails, use social media
    (Facebook, Instagram, TikTok), chat on WhatsApp and make video calls on Zoom.
    For education, students can access online courses, download textbooks and
    do research for school projects. Many Ivorian teachers send homework and
    course materials to students via WhatsApp groups.
    For business, companies use the Internet to sell products, advertise services
    and manage their finances. In Côte d'Ivoire, platforms like Jumia allow
    Ivorians to shop online and have products delivered to their door.
    (Internet a de nombreux usages aujourd'hui. Pour l'information, les gens
    l'utilisent pour chercher des nouvelles, lire des articles et regarder des vidéos
    éducatives sur YouTube. Pour la communication, Internet permet d'envoyer des e-mails,
    d'utiliser les réseaux sociaux, de chatter sur WhatsApp et de faire des appels vidéo.
    Pour l'éducation, les élèves peuvent accéder à des cours en ligne, télécharger des manuels
    et faire des recherches. Pour les affaires, les entreprises utilisent Internet
    pour vendre des produits et gérer leurs finances.)

▶ HARKNESS 3 — Internet safety and responsible use
  Q : How can we use the Internet safely and responsibly?
      (Comment peut-on utiliser Internet de façon sûre et responsable ?)
  Guide :
    · Pense aux risques : fausses informations, arnaques, cyberharcèlement, vie privée.
    · Pense aux règles de sécurité : mots de passe, partage d'informations personnelles.
    · Pense à l'équilibre : temps d'écran, usage responsable pour les élèves.
  Corrigé :
    The Internet has many risks that users must be aware of. Fake news —
    false information presented as true — spreads quickly online and can mislead
    people. Online scams try to steal money or personal information. Cyberbullying
    can cause serious psychological harm to young people. Sharing personal
    information like home addresses, phone numbers or photos with strangers is dangerous.
    To use the Internet safely, you should use strong passwords and never share them.
    You must not give personal information to strangers online. You should verify
    information before sharing it, to avoid spreading fake news. You should limit
    your screen time to avoid addiction. If you are a victim of cyberbullying,
    you must tell a trusted adult immediately.
    In Côte d'Ivoire, the government and organisations like ARTCI work to
    promote safe and responsible Internet use, especially among young people.
    (Internet comporte de nombreux risques. Les fausses nouvelles se propagent
    rapidement. Les arnaques tentent de voler de l'argent ou des informations
    personnelles. Le cyberharcèlement peut causer des dommages psychologiques graves.
    Pour utiliser Internet en sécurité, il faut utiliser des mots de passe solides,
    ne jamais partager ses informations personnelles avec des inconnus, vérifier
    les informations avant de les partager et limiter son temps d'écran.
    En Côte d'Ivoire, le gouvernement et des organisations comme ARTCI travaillent
    à promouvoir un usage sûr et responsable d'Internet.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre : The Internet — L'Internet
Compétence : C8 · TIC / ICT

Objectifs de la leçon :
  → Définir et décrire Internet et ses principales utilisations en anglais.
  → Distinguer Internet, le Web et les réseaux sociaux.
  → Utiliser les modaux MUST / SHOULD pour formuler des règles de sécurité en ligne.
  → Utiliser les connecteurs logiques pour argumenter sur les avantages et risques d'Internet.
  → Enrichir son vocabulaire sur les TIC et la vie numérique.

Prérequis nécessaires :
  → Les modaux MUST / SHOULD / MUST NOT vus en leçons 20, 21, 22, 23
  → Les connecteurs logiques : on the one hand / however / moreover / in conclusion
  → Les structures : "is used to / allows you to / consists of" (leçons 22 et 23)
  → Vocabulaire de base : computer · phone · network · screen · search

┌─ TABLEAU DES STRUCTURES FONDAMENTALES ──────────────────┐
  Expression / Notion            | Valeur, emploi et exemple traduit
  ───────────────────────────────|──────────────────────────────────
  The Internet allows you to...  | Exprimer ce qu'Internet permet
                                 | "The Internet allows you to search
                                 |  for information."
                                 | (Internet te permet de chercher
                                 |  des informations.)
  ───────────────────────────────|──────────────────────────────────
  You can use the Internet to... | Possibilité d'utilisation
                                 | "You can use the Internet to watch
                                 |  educational videos."
                                 | (Tu peux utiliser Internet pour
                                 |  regarder des vidéos éducatives.)
  ───────────────────────────────|──────────────────────────────────
  You must not + base verbale    | Règle de sécurité absolue
                                 | "You must not share your password."
                                 | (Tu ne dois pas partager ton
                                 |  mot de passe.)
  ───────────────────────────────|──────────────────────────────────
  You should + base verbale      | Conseil de sécurité
                                 | "You should verify information
                                 |  before sharing it."
                                 | (Tu devrais vérifier les informations
                                 |  avant de les partager.)
  ───────────────────────────────|──────────────────────────────────
  on the one hand... on the      | Présenter deux points de vue
  other hand...                  | "On the one hand, the Internet
                                 |  is useful. On the other hand,
                                 |  it can be dangerous."
                                 | (D'un côté... D'un autre côté...)
  ───────────────────────────────|──────────────────────────────────
  It is important to + inf.      | Souligner l'importance d'une action
                                 | "It is important to use the Internet
                                 |  responsibly."
                                 | (Il est important d'utiliser Internet
                                 |  de façon responsable.)
  ───────────────────────────────|──────────────────────────────────
  to be connected to             | Être connecté à
                                 | "I am connected to the Internet
                                 |  via WiFi."
                                 | (Je suis connecté à Internet
                                 |  via le WiFi.)
└──────────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

TEXTE DE BASE — À lire et comprendre

  "The Internet is a global network that connects millions of computers,
  smartphones and other devices around the world. It allows people to
  share and access information at any time and from any place.
  The Internet has many uses : searching for information on search engines
  like Google, sending emails, using social media like Facebook and WhatsApp,
  watching videos on YouTube, shopping online and studying from home.
  In Côte d'Ivoire, Internet access has grown rapidly thanks to mobile data
  and WiFi networks. Students use the Internet to do research, download
  course materials and communicate with their teachers.
  However, the Internet also has risks : fake news, online scams, cyberbullying
  and addiction. It is important to use the Internet safely and responsibly.
  You must protect your personal information and you should always verify
  information before sharing it."

  TRADUCTION FRANÇAISE :
  "Internet est un réseau mondial qui connecte des millions d'ordinateurs,
  de smartphones et d'autres appareils dans le monde entier. Il permet aux gens
  de partager et d'accéder à des informations à tout moment et depuis n'importe où.
  Internet a de nombreux usages : chercher des informations sur des moteurs
  de recherche comme Google, envoyer des e-mails, utiliser les réseaux sociaux
  comme Facebook et WhatsApp, regarder des vidéos sur YouTube, faire des achats
  en ligne et étudier depuis chez soi.
  En Côte d'Ivoire, l'accès à Internet a rapidement progressé grâce aux données
  mobiles et aux réseaux WiFi. Les élèves utilisent Internet pour faire des
  recherches, télécharger des cours et communiquer avec leurs professeurs.
  Cependant, Internet comporte aussi des risques : les fausses nouvelles,
  les arnaques en ligne, le cyberharcèlement et la dépendance.
  Il est important d'utiliser Internet de façon sûre et responsable.
  Tu dois protéger tes informations personnelles et tu devrais toujours vérifier
  les informations avant de les partager."

────────────────────────────────────────────────────────

┌─ À RETENIR — Internet, Web et Réseaux sociaux ─────────┐

  INTERNET :
  → Le RÉSEAU physique mondial — câbles, satellites, connexions.
  → Prononciation : [ˈɪntənet]
  → Créé dans les années 1960 par l'armée américaine (ARPANET).
  → Aujourd'hui : plus de 5 milliards d'utilisateurs dans le monde.

  THE WEB (World Wide Web / WWW) :
  → L'ensemble des SITES WEB accessibles via un navigateur (browser).
  → Prononciation : [wɜːld waɪd web]
  → Le Web UTILISE Internet — mais Internet ≠ le Web.
  → Exemple : www.google.com est sur le Web, accessible via Internet.

  SOCIAL MEDIA (Réseaux sociaux) :
  → Plateformes de partage et de communication en ligne.
  → Prononciation : [ˈsəʊʃəl ˈmiːdiə]
  → Exemples : Facebook · Instagram · TikTok · WhatsApp · YouTube · Twitter/X

  RÉSUMÉ :
  Internet → réseau physique (infrastructure)
  Web → sites et pages accessibles via browser
  Social media → plateformes de partage sur le Web

└──────────────────────────────────────────────────────────┘

⚠ ATTENTION — Faux amis et pièges
  · "navigation" sur internet → "browsing" [ˈbraʊzɪŋ] — pas "navigation" seul
    → "I am browsing the Internet." (Je navigue sur Internet.) ✔
  · "to surf the Internet" est une expression correcte et courante :
    → "I surf the Internet every day." (Je navigue sur Internet tous les jours.)
  · "e-mail" peut s'écrire "email" (sans trait d'union) — les deux sont acceptés.
  · "WiFi" se prononce [ˈwaɪfaɪ] — comme "why-fy" — jamais "wee-fee" à l'anglaise.
  · "Google" s'utilise aussi comme verbe en anglais familier :
    → "I googled it." = Je l'ai cherché sur Google.
  · "fake news" [feɪk njuːz] = fausses informations — mot à mot "nouvelles fausses".

────────────────────────────────────────────────────────

VOCABULAIRE ESSENTIEL — The Internet

  ANGLAIS                    | PRONONCIATION           | TRADUCTION FR
  ───────────────────────────|─────────────────────────|──────────────────────────
  Internet                   | [ˈɪntənet]              | Internet
  World Wide Web / WWW       | [wɜːld waɪd web]        | Web mondial / toile mondiale
  website                    | [ˈwebsaɪt]              | site web
  web page                   | [web peɪdʒ]             | page web
  web browser                | [web ˈbraʊzər]          | navigateur web
  search engine              | [sɜːtʃ ˈendʒɪn]        | moteur de recherche
  to browse / to surf        | [braʊz / sɜːf]          | naviguer / surfer
  to search for              | [sɜːtʃ fɔː]             | chercher / rechercher
  to download                | [ˈdaʊnləʊd]             | télécharger
  to upload                  | [ˈʌpləʊd]               | mettre en ligne / téléverser
  to stream                  | [striːm]                | regarder en ligne (streaming)
  email                      | [ˈiːmeɪl]               | e-mail / courriel
  social media               | [ˈsəʊʃəl ˈmiːdiə]      | réseaux sociaux
  WiFi                       | [ˈwaɪfaɪ]               | WiFi
  mobile data                | [ˈməʊbaɪl ˈdeɪtə]      | données mobiles / forfait
  password                   | [ˈpɑːswɜːd]             | mot de passe
  username                   | [ˈjuːzərneɪm]           | nom d'utilisateur
  online                     | [ˈɒnlaɪn]               | en ligne
  offline                    | [ˈɒflaɪn]               | hors ligne
  fake news                  | [feɪk njuːz]            | fausses informations
  cyberbullying              | [ˈsaɪbəbʊliɪŋ]         | cyberharcèlement
  scam / online fraud        | [skæm / frɔːd]          | arnaque / fraude en ligne
  privacy                    | [ˈprɪvəsi]              | vie privée / confidentialité
  to connect                 | [kəˈnekt]               | se connecter
  bandwidth / connection     | [ˈbændwɪdθ]             | bande passante / connexion
  screen time                | [skriːn taɪm]           | temps d'écran

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Les usages d'Internet / Uses of the Internet]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Titre en haut : USES OF THE INTERNET (Les usages d'Internet)
  · Centre (cercle bleu) : THE INTERNET
  · 5 branches rayonnant du centre :

  Branche 1 (verte) : INFORMATION
    → Search engines (moteurs de recherche) : Google, Yahoo
    → Online encyclopaedias : Wikipedia
    → News websites (sites d'actualité)

  Branche 2 (orange) : COMMUNICATION
    → Email (e-mail)
    → Social media : WhatsApp, Facebook, Instagram
    → Video calls : Zoom, Google Meet

  Branche 3 (bleue) : EDUCATION
    → Online courses (cours en ligne)
    → Educational videos : YouTube
    → Research for school (recherches scolaires)

  Branche 4 (violette) : ENTERTAINMENT
    → Streaming : Netflix, YouTube
    → Online games (jeux en ligne)
    → Music (musique en ligne)

  Branche 5 (rouge) : BUSINESS / COMMERCE
    → Online shopping : Jumia CI
    → Online banking (banque en ligne)
    → E-commerce (commerce électronique)

  · Chaque branche porte sa traduction française en italique
  · Légende : couleur par domaine d'usage

Note générateur : Canva (Mind Map) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Internet : Avantages et risques]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Titre en haut : THE INTERNET — FOR AND AGAINST
  · Deux colonnes côte à côte séparées par une ligne verticale :

  Colonne gauche (fond vert clair) : ADVANTAGES (Avantages)
    Sous-titre : "The Internet is useful because..."
    Liste :
    ✔ Access to information (Accès à l'information)
    ✔ Fast communication (Communication rapide)
    ✔ Educational resources (Ressources éducatives)
    ✔ Online shopping (Achats en ligne)
    ✔ Entertainment (Divertissement)
    ✔ Business opportunities (Opportunités d'affaires)

  Colonne droite (fond rouge clair) : RISKS / DISADVANTAGES (Risques / Inconvénients)
    Sous-titre : "The Internet can be dangerous because..."
    Liste :
    ✘ Fake news (Fausses informations)
    ✘ Online scams (Arnaques en ligne)
    ✘ Cyberbullying (Cyberharcèlement)
    ✘ Addiction (Dépendance)
    ✘ Privacy risks (Risques sur la vie privée)
    ✘ Excessive screen time (Temps d'écran excessif)

  · En bas : encadré "HOW TO STAY SAFE ONLINE" avec 3 règles clés :
    → Use strong passwords. · Do not share personal information.
    → Verify information before sharing it.

Note générateur : Canva (Two-column infographic) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Internet (Internet)
  EN : The Internet is a global network of interconnected computers and devices
  that allows the sharing and accessing of information and services from
  anywhere in the world at any time.
  FR : Internet est un réseau mondial d'ordinateurs et d'appareils interconnectés
  qui permet le partage et l'accès à des informations et des services depuis
  n'importe où dans le monde, à tout moment.

◆ DÉFINITION DPFC — World Wide Web / WWW (Web mondial)
  EN : The World Wide Web is a system of websites and web pages accessible
  through a web browser via the Internet. It was invented by Tim Berners-Lee
  in 1989. The Web uses the Internet but is not the same as the Internet.
  FR : Le World Wide Web est un système de sites web et de pages web accessibles
  via un navigateur web sur Internet. Il a été inventé par Tim Berners-Lee
  en 1989. Le Web utilise Internet mais n'est pas la même chose qu'Internet.

◆ DÉFINITION DPFC — Search engine (Moteur de recherche)
  EN : A search engine is a software programme that allows users to search
  for information on the Internet by entering keywords. The most popular
  search engine in the world is Google.
  FR : Un moteur de recherche est un programme informatique qui permet aux
  utilisateurs de rechercher des informations sur Internet en entrant des
  mots-clés. Le moteur de recherche le plus populaire au monde est Google.

◆ DÉFINITION DPFC — Social media (Réseaux sociaux)
  EN : Social media are online platforms that allow users to create, share
  and interact with content and with other users. Examples include Facebook,
  Instagram, WhatsApp, TikTok and YouTube.
  FR : Les réseaux sociaux sont des plateformes en ligne qui permettent
  aux utilisateurs de créer, partager et interagir avec du contenu et
  avec d'autres utilisateurs. Exemples : Facebook, Instagram, WhatsApp,
  TikTok et YouTube.

◆ DÉFINITION DPFC — Fake news (Fausses informations)
  EN : Fake news refers to false or misleading information presented as
  real news and spread on the Internet and social media. It is important
  to verify information from reliable sources before sharing it.
  FR : Les fausses informations désignent des informations fausses ou
  trompeuses présentées comme de vraies nouvelles et diffusées sur Internet
  et les réseaux sociaux. Il est important de vérifier les informations
  auprès de sources fiables avant de les partager.

◆ DÉFINITION DPFC — Cyberbullying (Cyberharcèlement)
  EN : Cyberbullying is the use of digital technology — especially the Internet
  and social media — to harass, threaten or humiliate another person.
  It can cause serious psychological harm to victims.
  FR : Le cyberharcèlement est l'utilisation des technologies numériques —
  notamment Internet et les réseaux sociaux — pour harceler, menacer ou
  humilier une autre personne. Il peut causer de graves dommages psychologiques
  aux victimes.

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

  ANGLAIS LV1 · Leçon 24 · 3ème
  THE INTERNET — L'Internet
  Compétence C8 · TIC / ICT · Prof. Evelyne ASSI

  I. VOCABULAIRE ESSENTIEL
  ─────────────────────────
  Internet [ˈɪntənet] = Internet
  World Wide Web / WWW = Web mondial
  website = site web · web page = page web
  web browser = navigateur web
  search engine = moteur de recherche
  to browse / to surf = naviguer / surfer sur Internet
  to search for = chercher / rechercher
  to download = télécharger · to upload = mettre en ligne
  to stream = regarder en streaming
  email = e-mail / courriel
  social media = réseaux sociaux
  WiFi [ˈwaɪfaɪ] = WiFi
  mobile data = données mobiles
  password = mot de passe · username = nom d'utilisateur
  online = en ligne · offline = hors ligne
  fake news = fausses informations
  cyberbullying = cyberharcèlement
  scam = arnaque · privacy = vie privée
  screen time = temps d'écran

  II. INTERNET, WEB ET RÉSEAUX SOCIAUX
  ──────────────────────────────────────
  · Internet = réseau mondial physique (câbles, satellites, connexions)
  · Web (WWW) = ensemble des sites web — utilise Internet
  · Social media = plateformes de partage sur le Web
    (Facebook, Instagram, WhatsApp, TikTok, YouTube)

  III. LES USAGES D'INTERNET
  ───────────────────────────
  The Internet allows you to... (Internet te permet de...) :
  · search for information (chercher des informations) — ex. Google
  · send emails (envoyer des e-mails)
  · use social media (utiliser les réseaux sociaux)
  · watch videos (regarder des vidéos) — ex. YouTube
  · shop online (faire des achats en ligne) — ex. Jumia CI
  · study from home (étudier depuis chez soi)
  · make video calls (faire des appels vidéo) — ex. Zoom

  IV. AVANTAGES ET RISQUES D'INTERNET
  ─────────────────────────────────────
  ADVANTAGES (Avantages) :
  · Access to information from anywhere in the world.
    (Accès à l'information depuis n'importe où dans le monde.)
  · Fast communication with family and friends.
    (Communication rapide avec la famille et les amis.)
  · Educational resources for students.
    (Ressources éducatives pour les élèves.)
  · Online shopping and business opportunities.
    (Achats en ligne et opportunités d'affaires.)

  RISKS (Risques) :
  · Fake news — false information that spreads quickly.
    (Fausses informations qui se propagent rapidement.)
  · Online scams that steal money or personal data.
    (Arnaques qui volent de l'argent ou des données personnelles.)
  · Cyberbullying — online harassment.
    (Cyberharcèlement — harcèlement en ligne.)
  · Addiction and excessive screen time.
    (Dépendance et temps d'écran excessif.)

  V. RÈGLES DE SÉCURITÉ EN LIGNE
  ────────────────────────────────
  · You must use a strong password and never share it.
    (Tu dois utiliser un mot de passe solide et ne jamais le partager.)
  · You must not give personal information to strangers online.
    (Tu ne dois pas donner tes informations personnelles à des inconnus en ligne.)
  · You should verify information before sharing it.
    (Tu devrais vérifier les informations avant de les partager.)
  · You should limit your screen time to avoid addiction.
    (Tu devrais limiter ton temps d'écran pour éviter la dépendance.)
  · If you are a victim of cyberbullying, you must tell a trusted adult.
    (Si tu es victime de cyberharcèlement, tu dois le dire à un adulte de confiance.)

~ ANALOGIE CI
  En Côte d'Ivoire, la RTI (Radio Télévision Ivoirienne) diffuse
  les informations à tous les Ivoiriens via les ondes radio et la télévision.
  Internet, c'est comme une RTI mondiale et interactive :
  non seulement tu reçois des informations, mais tu peux aussi en envoyer,
  commenter, partager et même créer ton propre contenu.
  "Internet is the world's biggest library, television and telephone — all in one."
  (Internet est la plus grande bibliothèque, télévision et téléphone du monde
  — tous réunis en un.)

~ ANALOGIE CI 2
  Imagine le marché de Cocody : des centaines de vendeurs, des milliers
  de produits, des acheteurs qui circulent, négocient et échangent.
  Internet, c'est ce marché — mais en version mondiale et accessible
  24h/24 depuis ton téléphone.
  Mais comme au marché, il y a de bons vendeurs et des arnaqueurs (scammers).
  C'est pourquoi il faut rester vigilant et ne jamais donner son argent
  ou ses informations sans vérifier.
  "On the Internet as in the market : beware of fraudsters."
  (Sur Internet comme au marché : méfie-toi des fraudeurs.)

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. Internet est un réseau mondial d'appareils connectés — le Web en est une partie.
  2. Internet permet de chercher des informations, communiquer, étudier, faire des achats
     et se divertir.
  3. Les principaux risques sont : fake news, arnaques, cyberharcèlement et dépendance.
  4. Règles de sécurité : mot de passe solide, ne pas partager ses informations personnelles,
     vérifier avant de partager.
  5. En Côte d'Ivoire, l'accès à Internet progresse via le mobile data et le WiFi —
     un outil essentiel pour les élèves et les professionnels.

5 mots clés + définition courte :
  · Internet → réseau mondial d'appareils interconnectés
  · Search engine → programme pour chercher des informations en ligne (ex. Google)
  · Social media → plateformes de partage et communication en ligne
  · Fake news → fausses informations diffusées comme vraies sur Internet
  · Cyberbullying → harcèlement d'une personne via Internet ou les réseaux sociaux

3 erreurs fréquentes + correction :
  Erreur 1 : Confondre Internet et le Web.
  Correction : Internet est le réseau physique. Le Web (WWW) est l'ensemble
               des sites accessibles via un navigateur. Le Web utilise Internet.

  Erreur 2 : "I surf in the Internet."
  Correction : "I surf the Internet." ou "I browse the Internet."
               → Pas de préposition "in" — on dit "on the Internet" ou verbe direct.

  Erreur 3 : "WiFi" se prononce [ˈwiːfiː] (à la française).
  Correction : En anglais, WiFi se prononce [ˈwaɪfaɪ] — comme "why-fy".
               À l'écrit, les deux orthographes WiFi et Wi-Fi sont acceptées.

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Vrai ou Faux ? · Notion ciblée : Compréhension d'Internet et du Web

  Énoncé :
  Lis chaque affirmation. Écris VRAI (True) ou FAUX (False).
  Si c'est FAUX, écris la phrase correcte en anglais et traduis-la en français.

  a) The Internet and the World Wide Web are exactly the same thing.
  b) Google is a search engine used to find information online.
  c) You can use the Internet to make video calls.
  d) Fake news is always easy to recognise on the Internet.
  e) In Côte d'Ivoire, students use the Internet to download course materials
     and communicate with their teachers.

  MÉTHODE
  Étape 1 · Lis chaque affirmation et rappelle-toi les définitions et faits vus en classe.
  Étape 2 · Décide : VRAI ou FAUX selon les informations de la leçon ?
  Étape 3 · Si FAUX, réécris la phrase correcte en anglais.
  Étape 4 · Traduis la phrase correcte en français.
  ✔ Vérification : Relis la trace écrite, surtout la section "Internet, Web et réseaux sociaux".
  ✔ Conclusion :
    a) FAUX — The Internet and the World Wide Web are NOT the same thing.
              The Web uses the Internet, but the Internet is much more than the Web.
              (Internet et le Web ne sont PAS la même chose. Le Web utilise Internet,
              mais Internet est bien plus que le Web.)
    b) VRAI — Google is a search engine used to find information online.
              (Google est un moteur de recherche utilisé pour trouver des informations en ligne.)
    c) VRAI — You can use the Internet to make video calls.
              (Tu peux utiliser Internet pour faire des appels vidéo.)
    d) FAUX — Fake news is NOT always easy to recognise on the Internet.
              That is why you should always verify information before sharing it.
              (Les fausses informations ne sont PAS toujours faciles à reconnaître sur Internet.
              C'est pourquoi tu devrais toujours vérifier les informations avant de les partager.)
    e) VRAI — In Côte d'Ivoire, students use the Internet to download course materials
              and communicate with their teachers.
              (En Côte d'Ivoire, les élèves utilisent Internet pour télécharger des cours
              et communiquer avec leurs professeurs.)
  ✔ Ce que cet exercice mobilise :
    · Compréhension des notions Internet / Web / moteur de recherche
    · Distinction Internet ≠ Web
    · Notion de fake news et vérification de l'information
    · Structure "You should verify..." (modal de conseil)
    · Traduction anglais → français · Esprit critique

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Règles de sécurité en ligne · Notion ciblée : Production écrite + Modaux + Sécurité

  Énoncé :
  Le lycée de Séraphin à Abidjan organise une semaine de l'Internet responsable.
  Séraphin doit rédiger une affiche de 8 règles de sécurité en ligne
  pour ses camarades de 3ème.
  Aide-le à écrire ces 8 règles en anglais, en utilisant MUST, MUST NOT et SHOULD.
  Traduis chaque règle en français.

  MÉTHODE
  Étape 1 · Pense aux différents aspects de la sécurité en ligne :
            mots de passe, informations personnelles, fake news, cyberharcèlement,
            temps d'écran, sources fiables, inconnus en ligne, images personnelles.
  Étape 2 · Pour chaque règle, choisis le modal approprié :
            → Règle stricte de sécurité : MUST
            → Interdiction : MUST NOT
            → Conseil fortement recommandé : SHOULD
  Étape 3 · Construis chaque phrase : Sujet + Modal + Base verbale + Complément.
  Étape 4 · Traduis chaque règle en français.
  ✔ Vérification : Tu dois avoir au moins 2 MUST, 3 MUST NOT et 2 SHOULD.
  ✔ Conclusion — Affiche de Séraphin :

    INTERNET SAFETY RULES — Règles de sécurité en ligne

    Rule 1 : You must use a strong password for all your online accounts.
             (Tu dois utiliser un mot de passe solide pour tous tes comptes en ligne.)
    Rule 2 : You must not share your password with anyone, not even your best friend.
             (Tu ne dois pas partager ton mot de passe avec quiconque,
             même pas ton meilleur ami.)
    Rule 3 : You must not give your personal information (address, phone number,
             school name) to strangers online.
             (Tu ne dois pas donner tes informations personnelles à des inconnus en ligne.)
    Rule 4 : You must not send photos of yourself to people you do not know in real life.
             (Tu ne dois pas envoyer tes photos à des gens que tu ne connais pas
             dans la vraie vie.)
    Rule 5 : You must not believe everything you read online — always verify information.
             (Tu ne dois pas croire tout ce que tu lis en ligne — vérifie toujours
             les informations.)
    Rule 6 : You should tell a trusted adult if someone online makes you feel
             uncomfortable or threatens you.
             (Tu devrais le dire à un adulte de confiance si quelqu'un en ligne
             te met mal à l'aise ou te menace.)
    Rule 7 : You should limit your screen time to avoid addiction and sleep problems.
             (Tu devrais limiter ton temps d'écran pour éviter la dépendance
             et les problèmes de sommeil.)
    Rule 8 : You must report cyberbullying to a teacher or parent immediately.
             (Tu dois signaler le cyberharcèlement à un professeur ou à un parent
             immédiatement.)

  ✔ Ce que cet exercice mobilise :
    · Production écrite guidée sur un thème DPFC (TIC et sécurité)
    · Modaux MUST / MUST NOT / SHOULD dans un contexte de sécurité numérique
    · Vocabulaire de la sécurité en ligne : password / personal information /
      cyberbullying / fake news / screen time
    · Traduction anglais → français
    · Éducation civique et numérique (contenu DPFC intégré)
    · Ancrage ivoirien (lycée d'Abidjan)

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulaire d'Internet · Notions travaillées : Lexique, traduction, sigles

  Partie A — Traduis ces mots anglais en français.
    1. search engine   →
    2. to download     →
    3. fake news       →
    4. password        →
    5. to stream       →
    6. privacy         →
    7. cyberbullying   →
    8. screen time     →

  Partie B — Traduis ces mots français en anglais.
    1. navigateur web       →
    2. réseaux sociaux      →
    3. mettre en ligne      →
    4. arnaque              →
    5. hors ligne           →

  Partie C — Que signifient ces sigles ? Écris le développé complet en anglais
  et la traduction en français.
    1. WWW = _____________ → FR : _____________
    2. WiFi = _____________ → FR : _____________

  GUIDE
  Étape 1 · Relis le tableau de vocabulaire essentiel de la trace écrite.
  Étape 2 · Pour les sigles, rappelle-toi les définitions vues en classe.
  Étape 3 · Vérifie l'orthographe de chaque mot.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Internet : Avantages ou Risques ? · Notions travaillées : Classement, argumentation

  Lis cette liste. Écris A (Advantage / Avantage) ou R (Risk / Risque)
  devant chaque élément. Traduis chaque élément en français.

  1. __ Access to educational resources from home.
  2. __ Cyberbullying among teenagers.
  3. __ Fast communication with family abroad.
  4. __ Fake news spreading on social media.
  5. __ Online shopping on platforms like Jumia.
  6. __ Addiction and excessive screen time.
  7. __ Video calls with teachers and classmates.
  8. __ Online scams targeting young people.
  9. __ Downloading course materials for free.
  10. __ Privacy risks from sharing personal information.

  GUIDE
  Étape 1 · Lis chaque élément et décide si c'est quelque chose de positif (A)
            ou de négatif (R) concernant l'usage d'Internet.
  Étape 2 · Traduis chaque élément en français.
  Étape 3 · Relis la trace écrite pour vérifier tes réponses.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Compréhension écrite · Notions travaillées : Lecture, compréhension, réponses en anglais

  Lis ce texte, puis réponds aux questions en anglais avec des phrases complètes.
  Traduis chaque réponse en français.

  TEXTE :
  "Mariama is a 15-year-old student from Abidjan. She uses the Internet every day
  for school and for fun. In the morning, she watches educational videos on YouTube
  to prepare for her BEPC exam. In the afternoon, she chats with her friends on
  WhatsApp and sometimes browses Facebook. Last week, she received a suspicious
  message from a stranger who claimed to offer her a prize of 500,000 francs CFA.
  Mariama did not reply to the message. She told her mother, who warned her that
  it was an online scam. Mariama's teacher, Mrs. Assi, also reminded the class
  that they must never give personal information to strangers online and that they
  should always verify information before believing it."

  Q1 : How old is Mariama and where does she live?
  Q2 : What does Mariama use YouTube for?
  Q3 : What happened to Mariama last week?
  Q4 : What did Mariama do when she received the suspicious message?
  Q5 : What rules did Mrs. Assi remind the class about?

  GUIDE
  Étape 1 · Lis le texte une première fois pour en comprendre le sens général.
  Étape 2 · Lis chaque question et repère dans le texte la partie qui y répond.
  Étape 3 · Écris ta réponse avec une phrase complète en anglais.
  Étape 4 · Traduis ta réponse en français.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Compléter un texte lacunaire · Notions travaillées : Vocabulaire, structures, modaux

  Complète le texte avec les mots suivants. Traduis le texte complet en français.
  (Mots à utiliser : search engine · social media · fake news · password ·
   download · browsing · online · cyberbullying · must · should)

  "The Internet is a powerful tool used for _______ websites, communicating
  on _______ and finding information with a _______ like Google.
  Students can _______ educational documents for free and study _______.
  However, the Internet also has risks. _______ spreads false information
  quickly. _______ causes serious harm to young people.
  To stay safe, you _______ use a strong _______ and never share it.
  You _______ also verify information before sharing it with others."

  GUIDE
  Étape 1 · Lis le texte entier une première fois pour comprendre le sens général.
  Étape 2 · Pour chaque blanc, choisis le mot qui convient grammaticalement et logiquement.
  Étape 3 · Relis le texte pour vérifier la cohérence.
  Étape 4 · Traduis le texte complet en français.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Production écrite · Notions travaillées : Expression écrite, argumentation, connecteurs

  Sujet : "The Internet : a useful tool or a dangerous trap for students?"
  (Internet : un outil utile ou un piège dangereux pour les élèves ?)

  Écris un texte argumentatif de 12 à 15 phrases. Tu dois :
    · Introduire le sujet (1-2 phrases)
    · Présenter 3 avantages d'Internet pour les élèves ivoiriens
    · Présenter 3 risques d'Internet pour les élèves
    · Utiliser les connecteurs : on the one hand / on the other hand /
      however / moreover / in conclusion
    · Donner ton opinion personnelle et 2 conseils avec SHOULD
    · Conclure avec une phrase d'espoir pour les élèves ivoiriens

  GUIDE
  Étape 1 · Introduction : "The Internet is one of the greatest inventions
            of the modern world. But is it always good for students?"
  Étape 2 · Avantages : information, communication, éducation, Jumia CI, YouTube.
  Étape 3 · Risques : fake news, cyberharcèlement, dépendance, arnaques.
  Étape 4 · Utilise les connecteurs pour structurer ton texte.
  Étape 5 · Conseils avec SHOULD : vérifier les informations, limiter le temps d'écran.
  Étape 6 · Conclusion positive sur l'avenir numérique des élèves ivoiriens.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — Ma lettre de fin d'année · A letter about ICT
  Durée conseillée : 40 minutes

  Énoncé :
  C'est ta DERNIÈRE LEÇON d'anglais de 3ème ! Pour marquer cet événement,
  ton professeur Evelyne ASSI te demande d'écrire une lettre (15 à 18 phrases)
  à un(e) élève imaginaire dans un pays anglophone (Ghana, Nigeria, Liberia)
  pour lui parler des TIC (Technologies de l'Information et de la Communication)
  en Côte d'Ivoire.

  Dans ta lettre, tu dois :
    · Te présenter et présenter ton pays (2 phrases)
    · Parler de l'ordinateur et de ses usages en CI (3 phrases) — Leçon 22
    · Parler du téléphone et du Mobile Money en CI (3 phrases) — Leçon 23
    · Parler d'Internet et de ses avantages pour les élèves ivoiriens (3 phrases) — Leçon 24
    · Mentionner 2 risques d'Internet et donner 2 conseils de sécurité (3 phrases)
    · Terminer par une phrase d'amitié et une formule de politesse

  Contraintes obligatoires :
    ✔ Utiliser "is used to" au moins 2 fois
    ✔ Utiliser "allows you to" au moins 1 fois
    ✔ Utiliser MUST ou MUST NOT au moins 2 fois
    ✔ Utiliser SHOULD au moins 2 fois
    ✔ Utiliser au moins 12 mots de vocabulaire des leçons 22, 23 et 24
    ✔ Traduire la lettre complète en français
    ✔ Format lettre : lieu + date / salutation / corps / formule de politesse / signature

  GUIDE
  Étape 1 · En-tête : Abidjan, Côte d'Ivoire, 14th July 2026
  Étape 2 · Salutation : "Dear [prénom],"
  Étape 3 · Présentation : "My name is... I am a student in 3ème at... in Abidjan, Côte d'Ivoire."
  Étape 4 · Ordinateur : "In my school, we use computers to... A computer is used to..."
  Étape 5 · Téléphone : "In Côte d'Ivoire, the mobile phone is very important because..."
  Étape 6 · Internet : "The Internet allows us to... We use it to..."
  Étape 7 · Risques + conseils : "However, the Internet has risks... You must... You should..."
  Étape 8 · Conclusion : "I hope we can communicate by email soon..."
  Étape 9 · Formule : "Yours sincerely," + prénom + NOM
  Étape 10 · Traduis toute la lettre en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON

  Abidjan, Côte d'Ivoire, 14th July 2026

  Dear Kwame,

  My name is Adjoua Koné. I am a 15-year-old student in 3ème at Lycée
  Classique d'Abidjan, in Côte d'Ivoire. I am writing to tell you about
  the role of ICT — Information and Communication Technology — in my country.
  (Je m'appelle Adjoua Koné. Je suis une élève de 15 ans en 3ème au Lycée
  Classique d'Abidjan, en Côte d'Ivoire. Je t'écris pour te parler du rôle
  des TIC — Technologies de l'Information et de la Communication — dans mon pays.)

  First, let me tell you about computers in Côte d'Ivoire.
  A computer is an electronic machine that consists of hardware and software.
  In my school, we have a computer room where each computer is used to write
  documents, do research and download educational resources.
  The government has launched the "École Numérique" programme to provide
  computers to all schools across the country.
  (Premièrement, laisse-moi te parler des ordinateurs en Côte d'Ivoire.
  Un ordinateur est une machine électronique composée de matériel et de logiciel.
  Dans mon école, nous avons une salle informatique où chaque ordinateur est utilisé
  pour écrire des documents, faire des recherches et télécharger des ressources éducatives.
  Le gouvernement a lancé le programme "École Numérique" pour fournir des ordinateurs
  à toutes les écoles du pays.)

  The mobile phone is also very important in Côte d'Ivoire.
  Almost every adult owns a smartphone, which is used to make calls, send
  messages and access the Internet. Moreover, Mobile Money services like
  Orange Money and MTN MoMo allow millions of Ivorians to send and receive
  money without a bank account — it is a true financial revolution.
  (Le téléphone mobile est aussi très important en Côte d'Ivoire.
  Presque chaque adulte possède un smartphone, utilisé pour passer des appels,
  envoyer des messages et accéder à Internet. De plus, les services Mobile Money
  comme Orange Money et MTN MoMo permettent à des millions d'Ivoiriens d'envoyer
  et de recevoir de l'argent sans compte bancaire — c'est une vraie révolution financière.)

  The Internet has changed our lives as students in many ways.
  It allows us to search for information on Google, watch educational videos
  on YouTube and communicate with our teachers on WhatsApp.
  In Côte d'Ivoire, more and more students connect to the Internet via mobile
  data or WiFi, and this opens many opportunities for our education and future.
  (Internet a changé nos vies en tant qu'élèves de nombreuses façons.
  Il nous permet de chercher des informations sur Google, de regarder des vidéos
  éducatives sur YouTube et de communiquer avec nos professeurs sur WhatsApp.
  En Côte d'Ivoire, de plus en plus d'élèves se connectent à Internet via les
  données mobiles ou le WiFi, ce qui ouvre de nombreuses opportunités pour notre
  éducation et notre avenir.)

  However, the Internet also has risks that we must be aware of.
  Fake news spreads very quickly on social media and can mislead people.
  Cyberbullying is also a serious problem that harms many young people online.
  You must never give your personal information to strangers on the Internet.
  You should always verify information before sharing it, and you should limit
  your screen time to stay healthy and focused on your studies.
  (Cependant, Internet comporte aussi des risques dont nous devons être conscients.
  Les fausses informations se propagent très rapidement sur les réseaux sociaux.
  Le cyberharcèlement est aussi un problème grave qui nuit à beaucoup de jeunes en ligne.
  Tu ne dois jamais donner tes informations personnelles à des inconnus sur Internet.
  Tu devrais toujours vérifier les informations avant de les partager, et tu devrais
  limiter ton temps d'écran pour rester en bonne santé et concentré sur tes études.)

  I hope this letter has given you a good picture of ICT in Côte d'Ivoire.
  I would love to hear about how students use technology in Ghana too.
  I hope we can communicate by email or WhatsApp soon — the Internet makes
  friendship across borders possible!
  (J'espère que cette lettre t'a donné une bonne image des TIC en Côte d'Ivoire.
  J'aimerais entendre parler de la façon dont les élèves utilisent la technologie
  au Ghana aussi. J'espère que nous pourrons bientôt communiquer par e-mail ou WhatsApp
  — Internet rend possible l'amitié par-delà les frontières !)

  Yours sincerely,
  Adjoua KONÉ
  Élève de 3ème · Lycée Classique d'Abidjan · Côte d'Ivoire

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1

  Partie A :
    1. search engine → moteur de recherche
    2. to download → télécharger
    3. fake news → fausses informations
    4. password → mot de passe
    5. to stream → regarder en streaming
    6. privacy → vie privée / confidentialité
    7. cyberbullying → cyberharcèlement
    8. screen time → temps d'écran

  Partie B :
    1. navigateur web → web browser
    2. réseaux sociaux → social media
    3. mettre en ligne → to upload
    4. arnaque → scam
    5. hors ligne → offline

  Partie C :
    1. WWW = World Wide Web → FR : Web mondial / toile mondiale
    2. WiFi = Wireless Fidelity → FR : WiFi (réseau sans fil)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2

  1. A — Access to educational resources from home.
         (Accès aux ressources éducatives depuis chez soi.)
  2. R — Cyberbullying among teenagers.
         (Cyberharcèlement chez les adolescents.)
  3. A — Fast communication with family abroad.
         (Communication rapide avec la famille à l'étranger.)
  4. R — Fake news spreading on social media.
         (Fausses informations qui se propagent sur les réseaux sociaux.)
  5. A — Online shopping on platforms like Jumia.
         (Achats en ligne sur des plateformes comme Jumia.)
  6. R — Addiction and excessive screen time.
         (Dépendance et temps d'écran excessif.)
  7. A — Video calls with teachers and classmates.
         (Appels vidéo avec les professeurs et les camarades.)
  8. R — Online scams targeting young people.
         (Arnaques en ligne visant les jeunes.)
  9. A — Downloading course materials for free.
         (Télécharger des cours gratuitement.)
  10. R — Privacy risks from sharing personal information.
          (Risques sur la vie privée liés au partage d'informations personnelles.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3

  Q1 : How old is Mariama and where does she live?
  R1 : Mariama is 15 years old and she lives in Abidjan.
       (Mariama a 15 ans et elle vit à Abidjan.)

  Q2 : What does Mariama use YouTube for?
  R2 : Mariama uses YouTube to watch educational videos to prepare for her BEPC exam.
       (Mariama utilise YouTube pour regarder des vidéos éducatives afin de préparer
       son examen du BEPC.)

  Q3 : What happened to Mariama last week?
  R3 : Last week, Mariama received a suspicious message from a stranger who claimed
       to offer her a prize of 500,000 francs CFA.
       (La semaine dernière, Mariama a reçu un message suspect d'un inconnu qui prétendait
       lui offrir un prix de 500 000 francs CFA.)

  Q4 : What did Mariama do when she received the suspicious message?
  R4 : Mariama did not reply to the message and told her mother, who warned her
       that it was an online scam.
       (Mariama n'a pas répondu au message et l'a dit à sa mère, qui l'a prévenue
       que c'était une arnaque en ligne.)

  Q5 : What rules did Mrs. Assi remind the class about?
  R5 : Mrs. Assi reminded the class that they must never give personal information
       to strangers online and that they should always verify information before believing it.
       (Mme Assi a rappelé à la classe qu'ils ne doivent jamais donner leurs informations
       personnelles à des inconnus en ligne et qu'ils devraient toujours vérifier
       les informations avant de les croire.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4

  Texte complété :
  "The Internet is a powerful tool used for BROWSING websites, communicating
  on SOCIAL MEDIA and finding information with a SEARCH ENGINE like Google.
  Students can DOWNLOAD educational documents for free and study ONLINE.
  However, the Internet also has risks. FAKE NEWS spreads false information
  quickly. CYBERBULLYING causes serious harm to young people.
  To stay safe, you MUST use a strong PASSWORD and never share it.
  You SHOULD also verify information before sharing it with others."

  Traduction française complète :
  "Internet est un outil puissant utilisé pour naviguer sur des sites web,
  communiquer sur les réseaux sociaux et trouver des informations avec
  un moteur de recherche comme Google. Les élèves peuvent télécharger des
  documents éducatifs gratuitement et étudier en ligne.
  Cependant, Internet comporte aussi des risques. Les fausses informations
  propagent rapidement de fausses informations. Le cyberharcèlement cause
  de graves dommages aux jeunes.
  Pour rester en sécurité, tu dois utiliser un mot de passe solide et
  ne jamais le partager. Tu devrais aussi vérifier les informations
  avant de les partager avec les autres."

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5

  Exemple de texte argumentatif :

  "The Internet is one of the greatest inventions of the modern world.
  But is it always good for students in Côte d'Ivoire?
  (Internet est l'une des plus grandes inventions du monde moderne.
  Mais est-il toujours bon pour les élèves en Côte d'Ivoire ?)

  On the one hand, the Internet has many advantages for students.
  First, it allows students to search for information on Google and download
  educational resources for free — a huge advantage for students who do not
  have access to many books.
  Moreover, the Internet allows students to communicate with their teachers
  via WhatsApp and receive course materials at home.
  In addition, students can watch educational videos on YouTube to prepare
  for their BEPC exam at any time of day.
  (D'un côté, Internet a de nombreux avantages pour les élèves.
  Premièrement, il permet aux élèves de chercher des informations sur Google
  et de télécharger des ressources éducatives gratuitement.
  De plus, Internet permet aux élèves de communiquer avec leurs professeurs
  via WhatsApp et de recevoir des cours à la maison.
  En outre, les élèves peuvent regarder des vidéos éducatives sur YouTube
  pour préparer leur examen du BEPC à tout moment de la journée.)

  On the other hand, the Internet also has serious risks for students.
  Fake news spreads quickly on social media and can mislead young people
  who do not verify their sources.
  However, the biggest danger is cyberbullying, which causes serious
  psychological harm to victims and can ruin their school results.
  Moreover, spending too much time on the Internet leads to addiction,
  sleep problems and poor concentration in class.
  (D'un autre côté, Internet comporte aussi de sérieux risques pour les élèves.
  Les fausses informations se propagent rapidement sur les réseaux sociaux.
  Cependant, le plus grand danger est le cyberharcèlement, qui cause de graves
  dommages psychologiques aux victimes.
  De plus, passer trop de temps sur Internet entraîne une dépendance,
  des problèmes de sommeil et une mauvaise concentration en classe.)

  In my opinion, the Internet is neither good nor bad in itself — it depends
  on how we use it. Students should use the Internet as a learning tool,
  not as a distraction. They should also verify information before sharing it
  and protect their personal data online.
  In conclusion, if Ivorian students use the Internet wisely and responsibly,
  it will be a powerful weapon for their success and for the development
  of Côte d'Ivoire.
  (À mon avis, Internet n'est ni bon ni mauvais en lui-même — cela dépend
  de la façon dont on l'utilise. Les élèves devraient utiliser Internet
  comme un outil d'apprentissage, pas comme une distraction. Ils devraient
  aussi vérifier les informations avant de les partager et protéger leurs
  données personnelles en ligne.
  En conclusion, si les élèves ivoiriens utilisent Internet de façon sage
  et responsable, ce sera une arme puissante pour leur succès et pour le
  développement de la Côte d'Ivoire.)"

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "The Internet is becoming the town square for the global village
  of tomorrow."
  (Internet devient la place du village pour le village mondial de demain.)
  — Bill Gates

  ✦ ✦ ✦

  Félicitations ! Tu as terminé les 24 leçons d'Anglais LV1 de 3ème.
  Tu es maintenant prêt(e) pour le BEPC !
  "Speak even if it's not perfect — silence teaches nothing."
  (Parle même si c'est imparfait — le silence n'apprend rien.)

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
