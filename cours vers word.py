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

COURS = """✅ Anglais LV1 · Leçon 23 — The telephone · en cours de génération...

---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 23 · Classe de 3ème

THE TELEPHONE
*Le téléphone*

Compétence C8 · TIC / ICT

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 23 — THE TELEPHONE
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C8 · TIC / ICT
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 22 : The computer)
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

CAPSULE DE REPRISE — Leçon 22 : The computer (L'ordinateur)

5 points clés à retenir :
  1. Un ordinateur est une machine électronique qui traite, stocke et affiche des informations.
  2. Le hardware désigne les composants physiques que l'on peut toucher
     (écran, clavier, souris, imprimante).
  3. Le software désigne les programmes que l'on ne peut pas toucher
     (système d'exploitation, navigateur web, traitement de texte).
  4. Le CPU est l'unité centrale — le "cerveau" de l'ordinateur.
  5. "Mouse" fait "mice" au pluriel — jamais "mouses".

3 questions flash :
  Q1 : What is the difference between hardware and software?
  R1 : Hardware is what you can touch (monitor, keyboard, mouse).
       Software is what you cannot touch (programmes like Windows or Word).
       (Le matériel est ce qu'on peut toucher. Le logiciel est ce qu'on ne peut pas toucher.)

  Q2 : Give two uses of a computer.
  R2 : A computer is used to write documents and search for information online.
       (Un ordinateur est utilisé pour écrire des documents et chercher des informations en ligne.)

  Q3 : What does CPU stand for?
  R3 : CPU stands for Central Processing Unit — it is the brain of the computer.
       (CPU signifie Unité Centrale de Traitement — c'est le cerveau de l'ordinateur.)

Lien avec la leçon du jour :
  La leçon 22 nous a montré comment l'ordinateur révolutionne le traitement
  de l'information. Aujourd'hui, nous allons étudier un autre outil TIC
  tout aussi révolutionnaire et beaucoup plus répandu en Côte d'Ivoire :
  LE TÉLÉPHONE — et particulièrement le téléphone mobile.
  Nous allons apprendre à le décrire, à parler de ses fonctions
  et à l'utiliser dans des situations de communication en anglais.

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

Mariama est élève en 3ème à Abidjan. Sa maman, vendeuse au marché de Treichville,
utilise son téléphone mobile Orange CI pour appeler ses fournisseurs,
envoyer de l'argent via Mobile Money et consulter les prix des produits.
Son grand frère Seydou, étudiant à l'Université FHB de Cocody, utilise WhatsApp
pour communiquer avec ses professeurs et recevoir ses cours en PDF.
Mariama elle-même utilise son téléphone pour appeler sa tante à Bouaké
chaque dimanche soir.
En Côte d'Ivoire, le téléphone mobile est l'outil numérique le plus utilisé —
bien plus que l'ordinateur.

→ Quels sont les différents types de téléphones ?
→ À quoi sert un téléphone aujourd'hui ?
→ Comment parler au téléphone en anglais ?
→ C'est exactement ce que nous allons apprendre aujourd'hui.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — Types of telephones
  Q : What are the different types of telephones and how have they changed over time?
      (Quels sont les différents types de téléphones et comment ont-ils évolué ?)
  Guide :
    · Pense au téléphone fixe (landline) et au téléphone mobile (mobile phone).
    · Pense à l'évolution : téléphone à cadran → portable basique → smartphone.
    · Pense aux marques connues en Côte d'Ivoire : Samsung, Tecno, iPhone, Infinix.
  Corrigé :
    There are two main types of telephones : the landline telephone, which is
    fixed in one place and connected by a wire, and the mobile phone (or cell phone),
    which is portable and works using wireless networks.
    Over the years, telephones have changed dramatically. The first telephones
    were large machines with a dial. Then came the basic mobile phone, used
    only for calls and text messages. Today, smartphones are small computers
    that can make calls, send messages, browse the internet, take photos
    and run thousands of applications.
    In Côte d'Ivoire, brands like Samsung, Tecno and Infinix are very popular
    because they offer affordable smartphones to Ivorian users.
    (Il existe deux types principaux de téléphones : le téléphone fixe, qui est
    installé en un endroit et connecté par un fil, et le téléphone mobile,
    qui est portable et fonctionne via des réseaux sans fil.
    Au fil des années, les téléphones ont beaucoup changé. Les premiers téléphones
    étaient de grandes machines avec un cadran. Puis sont venus les téléphones
    mobiles basiques, utilisés uniquement pour les appels et les SMS.
    Aujourd'hui, les smartphones sont de petits ordinateurs qui peuvent passer
    des appels, envoyer des messages, naviguer sur internet, prendre des photos
    et faire tourner des milliers d'applications.
    En Côte d'Ivoire, des marques comme Samsung, Tecno et Infinix sont très
    populaires car elles proposent des smartphones abordables.)

▶ HARKNESS 2 — Uses of the telephone
  Q : What are the different uses of the telephone today?
      (Quels sont les différents usages du téléphone aujourd'hui ?)
  Guide :
    · Pense aux usages de base : appels, SMS.
    · Pense aux usages modernes : internet, réseaux sociaux, Mobile Money, photos.
    · Pense aux usages professionnels et éducatifs en Côte d'Ivoire.
  Corrigé :
    The telephone has many uses today. Its basic use is to make and receive
    phone calls and send text messages (SMS). However, modern smartphones
    do much more : they allow users to browse the internet, send emails,
    take photos and videos, listen to music and use social media apps
    like WhatsApp, Facebook and TikTok.
    In Côte d'Ivoire, Mobile Money services like Orange Money and MTN MoMo
    allow people to send and receive money directly from their phones —
    a revolutionary tool for people who do not have bank accounts.
    In education, students use their phones to receive course materials,
    watch educational videos and communicate with their teachers.
    (Le téléphone a de nombreux usages aujourd'hui. Son usage de base est de
    passer et recevoir des appels et d'envoyer des SMS. Cependant, les smartphones
    modernes font beaucoup plus : ils permettent de naviguer sur internet, d'envoyer
    des e-mails, de prendre des photos et des vidéos, d'écouter de la musique
    et d'utiliser des applications de réseaux sociaux comme WhatsApp, Facebook et TikTok.
    En Côte d'Ivoire, les services Mobile Money comme Orange Money et MTN MoMo
    permettent aux gens d'envoyer et de recevoir de l'argent directement depuis
    leur téléphone — un outil révolutionnaire pour les personnes sans compte bancaire.
    Dans l'éducation, les élèves utilisent leurs téléphones pour recevoir des cours,
    regarder des vidéos éducatives et communiquer avec leurs professeurs.)

▶ HARKNESS 3 — Advantages and disadvantages
  Q : What are the advantages and disadvantages of the telephone?
      (Quels sont les avantages et les inconvénients du téléphone ?)
  Guide :
    · Avantages : communication rapide, accès à l'information, Mobile Money, sécurité.
    · Inconvénients : distraction à l'école, dépendance, cyberharcèlement, arnaques.
    · Pense à l'équilibre : comment utiliser le téléphone de façon responsable ?
  Corrigé :
    The telephone has many advantages. It allows people to communicate quickly
    with family and friends, even over long distances. It gives access to information,
    educational resources and financial services like Mobile Money. In emergencies,
    a phone can save lives by allowing people to call for help.
    However, the telephone also has disadvantages. Students who use their phones
    during class are distracted and learn less. Excessive phone use can lead to
    addiction and sleep problems. Cyberbullying and online scams (arnaque) are also
    serious risks, especially for young people.
    The key is to use the telephone responsibly : as a tool for learning and
    communication, not as a source of distraction.
    (Le téléphone a de nombreux avantages. Il permet aux gens de communiquer
    rapidement avec leur famille et leurs amis, même à longue distance. Il donne
    accès à l'information, aux ressources éducatives et aux services financiers
    comme le Mobile Money. En cas d'urgence, un téléphone peut sauver des vies.
    Cependant, le téléphone a aussi des inconvénients. Les élèves qui utilisent
    leur téléphone en classe sont distraits et apprennent moins. L'utilisation
    excessive du téléphone peut entraîner une dépendance et des problèmes de sommeil.
    Le cyberharcèlement et les arnaques en ligne sont aussi des risques sérieux.
    La clé est d'utiliser le téléphone de façon responsable : comme un outil
    d'apprentissage et de communication, pas comme une source de distraction.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre : The telephone — Le téléphone
Compétence : C8 · TIC / ICT

Objectifs de la leçon :
  → Nommer et décrire les types de téléphones et leurs fonctions en anglais.
  → Utiliser les expressions téléphoniques pour simuler une conversation.
  → Utiliser le Present Simple et le Present Perfect pour parler du téléphone.
  → Distinguer avantages et inconvénients du téléphone avec des connecteurs logiques.
  → Enrichir son vocabulaire sur la communication et les TIC.

Prérequis nécessaires :
  → Le Present Simple pour les habitudes et les faits généraux
  → Le Present Perfect : "I have called / I have received" (vu en grammaire 3ème)
  → Les modaux MUST / SHOULD vus en leçons 20 et 21
  → Vocabulaire de base : call · speak · listen · message · phone
  → Structures de la leçon 22 : "is used to / allows you to"

┌─ TABLEAU DES STRUCTURES FONDAMENTALES ──────────────────┐
  Expression / Notion            | Valeur, emploi et exemple traduit
  ───────────────────────────────|──────────────────────────────────
  Hello, this is [name]          | Se présenter au téléphone
                                 | "Hello, this is Kofi speaking."
                                 | (Allô, c'est Kofi à l'appareil.)
  ───────────────────────────────|──────────────────────────────────
  Can I speak to [name], please? | Demander à parler à quelqu'un
                                 | "Can I speak to Awa, please?"
                                 | (Puis-je parler à Awa, s'il vous plaît ?)
  ───────────────────────────────|──────────────────────────────────
  Hold on, please.               | Demander de patienter
                                 | "Hold on, please. I'll get her."
                                 | (Patientez s'il vous plaît.
                                 |  Je vais la chercher.)
  ───────────────────────────────|──────────────────────────────────
  Can I leave a message?         | Laisser un message
                                 | "Can I leave a message for him?"
                                 | (Puis-je laisser un message pour lui ?)
  ───────────────────────────────|──────────────────────────────────
  I'll call back later.          | Rappeler plus tard
                                 | "She is not here. I'll call back later."
                                 | (Elle n'est pas là. Je rappellerai
                                 |  plus tard.)
  ───────────────────────────────|──────────────────────────────────
  The line is busy.              | La ligne est occupée
                                 | "I tried to call but the line is busy."
                                 | (J'ai essayé d'appeler mais la ligne
                                 |  est occupée.)
  ───────────────────────────────|──────────────────────────────────
  I'll text you.                 | Envoyer un SMS
                                 | "I can't talk now. I'll text you."
                                 | (Je ne peux pas parler maintenant.
                                 |  Je t'enverrai un SMS.)
  ───────────────────────────────|──────────────────────────────────
  It is used for + V-ing         | Exprimer l'utilité
                                 | "A phone is used for making calls."
                                 | (Un téléphone est utilisé pour
                                 |  passer des appels.)
└──────────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

DIALOGUE DE BASE — À lire et comprendre

  [Le téléphone sonne chez Awa, à Abidjan]

  AWA    : Hello?
  KOFI   : Hello, this is Kofi speaking. Can I speak to Awa, please?
  AWA    : Speaking! Hi Kofi, how are you?
  KOFI   : I'm fine, thank you. I'm calling because I need the homework
            for tomorrow's English class.
  AWA    : Oh, hold on please. Let me check my notebook.
            [pause]
            Are you still there?
  KOFI   : Yes, I'm here.
  AWA    : The homework is to read page 45 and answer the questions.
  KOFI   : Thank you very much. By the way, have you received the message
            from our teacher on WhatsApp?
  AWA    : Not yet. I'll check it now and text you if it's important.
  KOFI   : Great. I'll call back later if I have more questions.
  AWA    : No problem. Goodbye, Kofi!
  KOFI   : Goodbye, Awa!

  TRADUCTION FRANÇAISE :
  AWA    : Allô ?
  KOFI   : Allô, c'est Kofi à l'appareil. Puis-je parler à Awa, s'il te plaît ?
  AWA    : C'est moi ! Salut Kofi, comment vas-tu ?
  KOFI   : Je vais bien, merci. J'appelle parce que j'ai besoin du devoir
            pour le cours d'anglais de demain.
  AWA    : Oh, patiente s'il te plaît. Laisse-moi vérifier mon cahier.
            [pause]
            Tu es toujours là ?
  KOFI   : Oui, je suis là.
  AWA    : Le devoir c'est de lire la page 45 et de répondre aux questions.
  KOFI   : Merci beaucoup. Au fait, as-tu reçu le message de notre professeur
            sur WhatsApp ?
  AWA    : Pas encore. Je vais le vérifier maintenant et je t'enverrai un SMS
            si c'est important.
  KOFI   : Super. Je rappellerai plus tard si j'ai d'autres questions.
  AWA    : Pas de problème. Au revoir, Kofi !
  KOFI   : Au revoir, Awa !

────────────────────────────────────────────────────────

┌─ À RETENIR — Expressions téléphoniques essentielles ───┐

  POUR COMMENCER UN APPEL :
  · Hello? = Allô ?
  · Hello, this is [name] speaking. = Allô, c'est [prénom] à l'appareil.
  · Can I speak to [name], please? = Puis-je parler à [prénom], s.v.p. ?

  POUR RÉPONDRE :
  · Speaking! = C'est moi !
  · Hold on, please. = Patientez / Attendez, s.v.p.
  · He/She is not available. = Il/Elle n'est pas disponible.
  · He/She is not here at the moment. = Il/Elle n'est pas là pour l'instant.

  EN CAS DE PROBLÈME :
  · The line is busy. = La ligne est occupée.
  · I can't hear you. = Je ne t'entends pas.
  · You are breaking up. = Tu coupes. (mauvais réseau)
  · I have no signal. = Je n'ai pas de réseau.

  POUR TERMINER / LAISSER UN MESSAGE :
  · Can I leave a message? = Puis-je laisser un message ?
  · I'll call back later. = Je rappellerai plus tard.
  · I'll text you. = Je t'enverrai un SMS.
  · Goodbye! / Bye! = Au revoir !

  PRONONCIATION CLÉS :
  · "speak" [spiːk] — le "ea" se dit "ee"
  · "message" [ˈmesɪdʒ] — le "g" se dit comme dans "gym"
  · "available" [əˈveɪləbl] — 4 syllabes : a-vail-a-ble
└──────────────────────────────────────────────────────────┘

⚠ ATTENTION — Faux amis et pièges
  · "phone" est invariable dans "on the phone" — jamais "on a phone"
    → "I am on the phone." (Je suis au téléphone.) ✔
  · "call" peut être nom ou verbe :
    → "I made a call." (nom — J'ai passé un appel.)
    → "I called him." (verbe — Je l'ai appelé.)
  · "text" s'utilise aussi comme verbe en anglais moderne :
    → "I'll text you." = Je t'enverrai un SMS.
    → Ce n'est PAS un faux ami — "texter" existe aussi en français familier.
  · "mobile" se prononce [ˈməʊbaɪl] en anglais britannique
    et [ˈmoʊbəl] en anglais américain.

────────────────────────────────────────────────────────

VOCABULAIRE ESSENTIEL — The telephone

  ANGLAIS                  | PRONONCIATION           | TRADUCTION FR
  ─────────────────────────|─────────────────────────|──────────────────────────
  telephone / phone        | [ˈtelɪfəʊn]             | téléphone
  mobile phone / cell phone| [ˈməʊbaɪl fəʊn]        | téléphone mobile
  smartphone               | [ˈsmɑːtfəʊn]           | smartphone
  landline                 | [ˈlændlaɪn]             | téléphone fixe
  SIM card                 | [sɪm kɑːd]              | carte SIM
  battery                  | [ˈbætəri]               | batterie
  charger                  | [ˈtʃɑːrdʒər]           | chargeur
  screen / touchscreen     | [skri:n / ˈtʌtʃskriːn] | écran / écran tactile
  signal / network         | [ˈsɪɡnəl / ˈnetwɜːk]  | réseau / signal
  call                     | [kɔːl]                  | appel (n.) / appeler (v.)
  text message / SMS       | [tekst ˈmesɪdʒ]         | SMS / message texte
  to call / to phone       | [kɔːl]                  | appeler / téléphoner
  to hang up               | [hæŋ ʌp]               | raccrocher
  to pick up               | [pɪk ʌp]               | décrocher
  to dial                  | [ˈdaɪəl]               | composer (un numéro)
  to text                  | [tekst]                 | envoyer un SMS
  to charge                | [tʃɑːrdʒ]              | charger (la batterie)
  voicemail                | [ˈvɔɪsmeɪl]            | messagerie vocale
  ringtone                 | [ˈrɪŋtəʊn]             | sonnerie
  application / app        | [ˌæplɪˈkeɪʃən]         | application
  Mobile Money             | [ˈməʊbaɪl ˈmʌni]       | Mobile Money
  to browse                | [braʊz]                 | naviguer (sur internet)
  advantage                | [ədˈvɑːntɪdʒ]          | avantage
  disadvantage             | [ˌdɪsədˈvɑːntɪdʒ]     | inconvénient

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Les types de téléphones / Types of telephones]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Titre en haut : TYPES OF TELEPHONES (Types de téléphones)
  · Ligne horizontale centrale divisant la page en deux zones :

  Zone gauche (fond bleu clair) :
    · Titre : LANDLINE TELEPHONE (Téléphone fixe)
    · Dessin : rectangle avec combiné et fil
    · Caractéristiques listées dessous :
        - Fixed in one place (Fixe en un endroit)
        - Connected by a wire (Connecté par un fil)
        - Used mainly for calls (Utilisé surtout pour les appels)
        - Less common today (Moins courant aujourd'hui)

  Zone droite (fond orange clair) :
    · Titre : MOBILE PHONE / SMARTPHONE (Téléphone mobile)
    · Dessin : rectangle fin vertical (forme de smartphone)
    · Caractéristiques listées dessous :
        - Portable — no wire (Portable — sans fil)
        - Calls + SMS + Internet (Appels + SMS + Internet)
        - Photos, apps, Mobile Money
        - Very common in Côte d'Ivoire
          (Très répandu en Côte d'Ivoire)

  · En bas : flèche d'évolution de gauche à droite :
    "From landline → basic phone → smartphone"
    (Du fixe → téléphone basique → smartphone)

Note générateur : Canva (Two-column infographic) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Carte mentale : Avantages et inconvénients du téléphone]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Centre (cercle bleu) : THE TELEPHONE (Le téléphone)
  · Deux branches principales :

  Branche gauche (fond vert) : ADVANTAGES (Avantages)
    → Fast communication (Communication rapide)
    → Access to information (Accès à l'information)
    → Mobile Money (transferts d'argent)
    → Emergency calls (Appels d'urgence)
    → Educational use (Usage éducatif)

  Branche droite (fond rouge clair) : DISADVANTAGES (Inconvénients)
    → Distraction at school (Distraction à l'école)
    → Addiction (Dépendance)
    → Cyberbullying (Cyberharcèlement)
    → Scams / fraud (Arnaques)
    → Sleep problems (Problèmes de sommeil)

  · Chaque élément porte sa traduction française en italique
  · Titre général : THE TELEPHONE — FOR AND AGAINST
  · Légende : vert = positif · rouge = négatif

Note générateur : Canva (Mind Map) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Telephone (Téléphone)
  EN : A telephone is a device used to transmit and receive voice
  communications over a distance using electrical signals or wireless networks.
  FR : Un téléphone est un appareil utilisé pour transmettre et recevoir
  des communications vocales à distance en utilisant des signaux électriques
  ou des réseaux sans fil.

◆ DÉFINITION DPFC — Mobile phone / Smartphone (Téléphone mobile / Smartphone)
  EN : A mobile phone is a portable telephone that works without a wire,
  using wireless networks. A smartphone is an advanced mobile phone that
  can also browse the internet, run applications and take photos.
  FR : Un téléphone mobile est un téléphone portable qui fonctionne sans fil,
  en utilisant des réseaux sans fil. Un smartphone est un téléphone mobile
  avancé qui peut aussi naviguer sur internet, exécuter des applications
  et prendre des photos.

◆ DÉFINITION DPFC — SMS / Text message (SMS / Message texte)
  EN : An SMS (Short Message Service) is a short written message sent
  from one phone to another. It is also called a text message or simply a text.
  FR : Un SMS (Service de Messages Courts) est un court message écrit
  envoyé d'un téléphone à un autre. On l'appelle aussi message texte ou simplement texte.

◆ DÉFINITION DPFC — Mobile Money (Mobile Money)
  EN : Mobile Money is a financial service that allows people to send,
  receive and store money using a mobile phone, without needing a bank account.
  In Côte d'Ivoire, examples include Orange Money and MTN MoMo.
  FR : Le Mobile Money est un service financier qui permet aux personnes
  d'envoyer, de recevoir et de stocker de l'argent en utilisant un téléphone
  mobile, sans avoir besoin d'un compte bancaire. En Côte d'Ivoire,
  les exemples incluent Orange Money et MTN MoMo.

◆ DÉFINITION DPFC — Network / Signal (Réseau / Signal)
  EN : A network is a system of connected devices that allows communication.
  A signal is the wireless transmission that allows a mobile phone to connect
  to the network and make calls or use the internet.
  FR : Un réseau est un système d'appareils connectés qui permet la communication.
  Un signal est la transmission sans fil qui permet à un téléphone mobile
  de se connecter au réseau et de passer des appels ou d'utiliser internet.

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

  ANGLAIS LV1 · Leçon 23 · 3ème
  THE TELEPHONE — Le téléphone
  Compétence C8 · TIC / ICT · Prof. Evelyne ASSI

  I. VOCABULAIRE ESSENTIEL
  ─────────────────────────
  telephone / phone = téléphone
  mobile phone / cell phone = téléphone mobile
  smartphone [ˈsmɑːtfəʊn] = smartphone
  landline = téléphone fixe
  SIM card = carte SIM
  battery = batterie · charger = chargeur
  signal / network = signal / réseau
  call (n. et v.) = appel / appeler
  text message / SMS = message texte / SMS
  to hang up = raccrocher · to pick up = décrocher
  to dial = composer un numéro · to text = envoyer un SMS
  voicemail = messagerie vocale · ringtone = sonnerie
  app / application = application
  Mobile Money = Mobile Money (ex. Orange Money, MTN MoMo)
  advantage = avantage · disadvantage = inconvénient

  II. EXPRESSIONS TÉLÉPHONIQUES
  ──────────────────────────────
  · Hello, this is [name] speaking.
    (Allô, c'est [prénom] à l'appareil.)
  · Can I speak to [name], please?
    (Puis-je parler à [prénom], s.v.p. ?)
  · Speaking! = C'est moi !
  · Hold on, please. = Patientez s.v.p.
  · He/She is not available. = Il/Elle n'est pas disponible.
  · The line is busy. = La ligne est occupée.
  · Can I leave a message? = Puis-je laisser un message ?
  · I'll call back later. = Je rappellerai plus tard.
  · I'll text you. = Je t'enverrai un SMS.
  · I have no signal. = Je n'ai pas de réseau.
  · You are breaking up. = Tu coupes. (mauvais réseau)
  · Goodbye! / Bye! = Au revoir !

  III. TYPES DE TÉLÉPHONES
  ─────────────────────────
  Landline telephone (Téléphone fixe) :
    · Fixed in one place — connected by a wire.
      (Fixe en un endroit — connecté par un fil.)
    · Used mainly for voice calls.
      (Utilisé principalement pour les appels vocaux.)

  Mobile phone / Smartphone (Téléphone mobile) :
    · Portable — works without a wire.
      (Portable — fonctionne sans fil.)
    · Used for calls, SMS, internet, photos, apps, Mobile Money.
      (Utilisé pour les appels, SMS, internet, photos, applications, Mobile Money.)

  IV. AVANTAGES ET INCONVÉNIENTS
  ────────────────────────────────
  ADVANTAGES (Avantages) :
    · Fast communication with family and friends.
      (Communication rapide avec la famille et les amis.)
    · Access to information and educational resources.
      (Accès à l'information et aux ressources éducatives.)
    · Mobile Money for financial transactions.
      (Mobile Money pour les transactions financières.)
    · Emergency calls can save lives.
      (Les appels d'urgence peuvent sauver des vies.)

  DISADVANTAGES (Inconvénients) :
    · Distraction at school and during study time.
      (Distraction à l'école et pendant le temps d'étude.)
    · Addiction and sleep problems.
      (Dépendance et problèmes de sommeil.)
    · Risk of cyberbullying and online scams.
      (Risque de cyberharcèlement et d'arnaques en ligne.)

~ ANALOGIE CI
  En Côte d'Ivoire, le griot du village utilisait autrefois le tam-tam
  pour transmettre des messages importants d'un village à l'autre.
  Aujourd'hui, le téléphone mobile a remplacé le tam-tam :
  un message WhatsApp arrive en une seconde à Bouaké, à San-Pédro
  ou même à Paris.
  "The smartphone is the modern tam-tam of Côte d'Ivoire."
  (Le smartphone est le tam-tam moderne de la Côte d'Ivoire.)

~ ANALOGIE CI 2
  Au marché de Marcory, la commerçante Adèle n'a pas de compte bancaire.
  Grâce à Orange Money sur son téléphone, elle envoie de l'argent
  à sa mère à Daloa en trente secondes — sans se déplacer, sans banque.
  Le téléphone mobile a transformé la vie économique de millions
  d'Ivoiriens qui n'ont pas accès aux services bancaires traditionnels.
  "A phone in your pocket is a bank in your hand."
  (Un téléphone dans ta poche, c'est une banque dans ta main.)

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. Il existe deux types principaux de téléphones : le fixe (landline) et le mobile (smartphone).
  2. Les expressions clés d'un appel : "Hello, this is... / Can I speak to... / Hold on / I'll call back."
  3. "to hang up" = raccrocher · "to pick up" = décrocher · "to dial" = composer un numéro.
  4. Le téléphone a des avantages (communication, Mobile Money, urgences)
     et des inconvénients (distraction, dépendance, arnaques).
  5. En Côte d'Ivoire, le Mobile Money (Orange Money, MTN MoMo) est une révolution financière.

5 mots clés + définition courte :
  · Smartphone → téléphone mobile avancé avec accès à internet et applications
  · Landline → téléphone fixe connecté par un fil
  · SMS / text message → court message écrit envoyé par téléphone
  · Mobile Money → service financier via téléphone mobile, sans compte bancaire
  · Signal / network → transmission sans fil permettant la connexion téléphonique

3 erreurs fréquentes + correction :
  Erreur 1 : "I am in the phone."
  Correction : "I am on the phone." — en anglais, on dit "on the phone", pas "in".

  Erreur 2 : "Can I speak with Awa, please?"
  Correction : "Can I speak TO Awa, please?" — on utilise "to", pas "with" dans ce contexte formel.

  Erreur 3 : Confondre "hang up" (raccrocher) et "pick up" (décrocher).
  Correction : "to pick up the phone" = décrocher (quand ça sonne)
               "to hang up" = raccrocher (pour terminer l'appel)
               Moyen mnémotechnique : PICK UP = prendre · HANG UP = poser/terminer.

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Dialogue téléphonique · Notion ciblée : Expressions téléphoniques

  Énoncé :
  Remets les répliques de ce dialogue dans le bon ordre (1 à 8).
  Traduis le dialogue complet en français.

  A. "Hold on, please. I'll get him."
  B. "Goodbye, Moussa!"
  C. "Hello? Koné family."
  D. "Hello, this is Moussa speaking. Can I speak to Seydou, please?"
  E. "Speaking! Hi Moussa, what's up?"
  F. "Thanks! Goodbye!"
  G. "Hi Seydou! I'm calling to remind you about tomorrow's football match at 4pm."
  H. "Oh right! Thanks for reminding me. I'll be there!"

  MÉTHODE
  Étape 1 · Identifie qui commence l'appel et qui répond.
  Étape 2 · Repère les expressions clés :
            "Hello?" → début · "Can I speak to...?" → demande
            "Hold on" → attente · "Speaking!" → la personne demandée répond
            "Goodbye" → fin de l'appel
  Étape 3 · Reconstitue la logique de la conversation.
  Étape 4 · Traduis chaque réplique en français.
  ✔ Vérification : Le dialogue doit avoir un début clair, un échange d'informations
                   et une fin polie.
  ✔ Conclusion — Ordre correct :
    1 → C : "Hello? Koné family."
            (Allô ? Famille Koné.)
    2 → D : "Hello, this is Moussa speaking. Can I speak to Seydou, please?"
            (Allô, c'est Moussa à l'appareil. Puis-je parler à Seydou, s.v.p. ?)
    3 → A : "Hold on, please. I'll get him."
            (Patientez s.v.p. Je vais le chercher.)
    4 → E : "Speaking! Hi Moussa, what's up?"
            (C'est moi ! Salut Moussa, quoi de neuf ?)
    5 → G : "Hi Seydou! I'm calling to remind you about tomorrow's football match at 4pm."
            (Salut Seydou ! J'appelle pour te rappeler le match de foot de demain à 16h.)
    6 → H : "Oh right! Thanks for reminding me. I'll be there!"
            (Ah oui ! Merci de me le rappeler. J'y serai !)
    7 → F : "Thanks! Goodbye!"
            (Merci ! Au revoir !)
    8 → B : "Goodbye, Moussa!"
            (Au revoir, Moussa !)
  ✔ Ce que cet exercice mobilise :
    · Expressions téléphoniques en situation réelle
    · Logique de la communication orale au téléphone
    · Traduction anglais → français
    · Vocabulaire de la leçon (speaking / hold on / goodbye)
    · Ancrage ivoirien (famille ivoirienne, match de foot)

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Avantages et inconvénients · Notion ciblée : Expression écrite + Connecteurs logiques

  Énoncé :
  Fatou et son frère Issa ont un désaccord. Fatou pense que le téléphone
  est très utile pour les élèves. Issa pense que le téléphone distrait
  les élèves et nuit à leurs études.
  Écris un texte de 8 phrases qui présente les deux points de vue,
  en utilisant les connecteurs logiques suivants :
  "on the one hand... on the other hand..." · "however" · "but" · "moreover"
  Traduis chaque phrase en français.

  MÉTHODE
  Étape 1 · Commence par présenter le sujet en une phrase.
  Étape 2 · Utilise "On the one hand..." pour les avantages (point de vue de Fatou).
  Étape 3 · Utilise "On the other hand..." ou "However..." pour les inconvénients (Issa).
  Étape 4 · Utilise "Moreover" pour ajouter un argument supplémentaire.
  Étape 5 · Conclus avec ton propre avis ou un conseil (avec SHOULD).
  Étape 6 · Traduis chaque phrase en français.
  ✔ Vérification : Vérifie que tu as utilisé au moins 3 connecteurs différents.
  ✔ Conclusion — Exemple de texte :
    1. The telephone is a subject of debate among students in Côte d'Ivoire.
       (Le téléphone est un sujet de débat chez les élèves en Côte d'Ivoire.)
    2. On the one hand, Fatou thinks the phone is very useful for students.
       (D'un côté, Fatou pense que le téléphone est très utile pour les élèves.)
    3. It allows them to search for information online and communicate with their teachers.
       (Il leur permet de chercher des informations en ligne et de communiquer
       avec leurs professeurs.)
    4. Moreover, Mobile Money helps students' families to manage money easily.
       (De plus, le Mobile Money aide les familles des élèves à gérer l'argent facilement.)
    5. On the other hand, Issa believes that phones distract students in class.
       (D'un autre côté, Issa croit que les téléphones distraient les élèves en classe.)
    6. However, students who use their phones during lessons learn less and get bad results.
       (Cependant, les élèves qui utilisent leur téléphone pendant les cours
       apprennent moins et obtiennent de mauvais résultats.)
    7. Moreover, excessive phone use can lead to addiction and sleep problems.
       (De plus, l'utilisation excessive du téléphone peut entraîner une dépendance
       et des problèmes de sommeil.)
    8. Students should use their phones responsibly — as a tool for learning,
       not as a source of distraction.
       (Les élèves devraient utiliser leur téléphone de façon responsable —
       comme un outil d'apprentissage, pas comme une source de distraction.)
  ✔ Ce que cet exercice mobilise :
    · Expression écrite argumentative guidée
    · Connecteurs logiques : on the one hand / on the other hand / however / moreover
    · Vocabulaire des avantages et inconvénients du téléphone
    · Modal SHOULD pour la recommandation finale
    · Traduction anglais → français · Ancrage ivoirien

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulaire du téléphone · Notions travaillées : Lexique, traduction, prononciation

  Partie A — Traduis ces mots anglais en français.
    1. charger      →
    2. ringtone     →
    3. voicemail    →
    4. to hang up   →
    5. landline     →
    6. SIM card     →
    7. to dial      →
    8. network      →

  Partie B — Traduis ces mots français en anglais.
    1. décrocher          →
    2. appel              →
    3. batterie           →
    4. message texte      →
    5. application        →

  GUIDE
  Étape 1 · Relis le tableau de vocabulaire essentiel de la trace écrite.
  Étape 2 · Attention à la prononciation des mots difficiles.
  Étape 3 · Vérifie l'orthographe de chaque mot anglais.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Expressions téléphoniques · Notions travaillées : Communication orale, formules au téléphone

  Associe chaque situation à l'expression téléphonique correcte.
  (Écris la lettre de l'expression devant le numéro de la situation.)

  SITUATIONS :
  1. Tu veux parler à Awa mais c'est son frère qui répond.
  2. La ligne est occupée et tu vas rappeler.
  3. Tu décroches le téléphone qui sonne.
  4. Tu veux demander à quelqu'un de patienter.
  5. La personne que tu veux appeler n'est pas là et tu veux laisser un message.
  6. Tu ne peux pas parler et tu vas envoyer un SMS.
  7. Tu te présentes au téléphone.
  8. La personne demandée, c'est toi qui réponds.

  EXPRESSIONS :
  A. "Can I speak to Awa, please?"
  B. "Hello? / Hello!"
  C. "I'll call back later."
  D. "Hold on, please."
  E. "Can I leave a message?"
  F. "I'll text you."
  G. "Hello, this is [name] speaking."
  H. "Speaking!"

  GUIDE
  Étape 1 · Lis chaque situation et imagine-toi au téléphone.
  Étape 2 · Choisis l'expression qui correspond exactement à la situation.
  Étape 3 · Traduis chaque expression choisie en français.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Compréhension écrite · Notions travaillées : Lecture, compréhension, réponses en anglais

  Lis ce texte, puis réponds aux questions en anglais avec des phrases complètes.
  Traduis chaque réponse en français.

  TEXTE :
  "In Côte d'Ivoire, the mobile phone has become an essential tool in everyday life.
  According to recent statistics, more than 80% of Ivorians own a mobile phone.
  People use their phones to call family members, send text messages, browse the
  internet and use Mobile Money services like Orange Money and MTN MoMo.
  For students, the phone is useful for receiving course materials and communicating
  with teachers. However, many schools have banned phones in classrooms because
  they distract students. The government encourages young people to use their
  phones responsibly and to develop their digital skills for the future."

  Q1 : What percentage of Ivorians own a mobile phone?
  Q2 : Name three things people use their phones for in Côte d'Ivoire.
  Q3 : Why have many schools banned phones in classrooms?
  Q4 : How can phones be useful for students?
  Q5 : What does the government encourage young people to do?

  GUIDE
  Étape 1 · Lis le texte une première fois pour en comprendre le sens général.
  Étape 2 · Lis chaque question et repère dans le texte la partie qui y répond.
  Étape 3 · Écris ta réponse avec une phrase complète en anglais.
  Étape 4 · Traduis ta réponse en français.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Compléter un dialogue · Notions travaillées : Expressions téléphoniques en contexte

  Complète ce dialogue avec les expressions suivantes :
  (Hold on · Can I speak to · Speaking · I'll call back · this is · the line is busy)

  [Le téléphone sonne chez Bamba]

  BAMBA  : Hello?
  AMINATA : Hello, _______ Aminata speaking. _______ Bamba, please?
  BAMBA  : _______ ! Hi Aminata!
  AMINATA : Hi Bamba! I tried to call you earlier but _______.
  BAMBA  : Sorry, I was talking to my mum.
  AMINATA : No problem. _______ please, I need to find my notebook.
  BAMBA  : Of course, take your time.
  AMINATA : Actually, I can't find it now. _______ later, OK?
  BAMBA  : Sure! Goodbye!
  AMINATA : Goodbye!

  GUIDE
  Étape 1 · Lis le dialogue entier pour comprendre le contexte.
  Étape 2 · Pour chaque blanc, choisis l'expression qui convient logiquement.
  Étape 3 · Traduis le dialogue complet en français une fois complété.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Production écrite · Notions travaillées : Expression écrite, avantages/inconvénients, connecteurs

  Le professeur de français de ton lycée a donné ce sujet de rédaction :
  "Le téléphone mobile : bienfait ou fléau pour les élèves ?"
  Ton professeur d'anglais te demande d'écrire le même texte EN ANGLAIS
  (12 à 15 phrases). Tu dois :
    · Introduire le sujet (1-2 phrases)
    · Présenter 3 avantages du téléphone pour les élèves
    · Présenter 3 inconvénients du téléphone pour les élèves
    · Utiliser les connecteurs : on the one hand / on the other hand /
      however / moreover / but / in conclusion
    · Conclure avec ton avis personnel et un conseil avec SHOULD

  GUIDE
  Étape 1 · Commence par : "The mobile phone is one of the most common tools
            among students in Côte d'Ivoire today."
  Étape 2 · Avantages : communication, accès à l'information, Mobile Money, urgences.
  Étape 3 · Inconvénients : distraction, dépendance, cyberharcèlement, arnaques.
  Étape 4 · Utilise les connecteurs pour structurer ton texte.
  Étape 5 · Conclus avec : "In conclusion, students should..."

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — A telephone conversation + a written opinion
  Durée conseillée : 35 minutes

  PARTIE A — Dialogue (8 répliques minimum)
  Écris un dialogue téléphonique entre deux élèves ivoiriens, Kouamé et Bintou.
  Kouamé appelle Bintou pour lui demander de l'aide pour ses devoirs d'anglais.
  Bintou n'est pas disponible au début — c'est sa mère qui répond.
  Le dialogue doit contenir obligatoirement :
    ✔ Une présentation au téléphone (this is... speaking)
    ✔ Une demande de parler à quelqu'un (Can I speak to...?)
    ✔ Une expression d'attente (Hold on...)
    ✔ Un échange d'informations sur les devoirs
    ✔ Une formule de fin d'appel (I'll call back / Goodbye)
    ✔ Traduction complète du dialogue en français

  PARTIE B — Opinion écrite (6 à 8 phrases)
  Donne ton avis personnel sur la question suivante :
  "Should students be allowed to use their phones in class?"
  (Les élèves devraient-ils être autorisés à utiliser leur téléphone en classe ?)
  Tu dois utiliser :
    ✔ SHOULD / SHOULD NOT au moins 2 fois
    ✔ Les connecteurs : however · moreover · in my opinion
    ✔ Au moins 6 mots de vocabulaire de la leçon
    ✔ Traduction complète de ton opinion en français

  GUIDE
  Partie A :
  Étape 1 · Écris les noms des personnages avant chaque réplique : KOUAMÉ / BINTOU / MÈRE DE BINTOU
  Étape 2 · Commence par la sonnerie et la réponse de la mère.
  Étape 3 · Kouamé se présente et demande à parler à Bintou.
  Étape 4 · La mère demande à Bintou de venir — attente.
  Étape 5 · Bintou et Kouamé échangent sur les devoirs.
  Étape 6 · Fin de l'appel avec formule de politesse.
  Étape 7 · Traduis tout le dialogue en français.

  Partie B :
  Étape 1 · Commence par : "In my opinion, students should / should not use their phones in class because..."
  Étape 2 · Donne 2-3 arguments avec "however" et "moreover".
  Étape 3 · Conclus avec une recommandation claire.
  Étape 4 · Traduis tout le texte en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON

  PARTIE A — Dialogue téléphonique

  [Le téléphone de Bintou sonne]

  MÈRE DE BINTOU : Hello?
                   (Allô ?)

  KOUAMÉ         : Hello, this is Kouamé speaking. Can I speak to Bintou, please?
                   (Allô, c'est Kouamé à l'appareil. Puis-je parler à Bintou, s.v.p. ?)

  MÈRE DE BINTOU : Hello Kouamé! Hold on, please. I'll call her.
                   (Bonjour Kouamé ! Patientez s.v.p. Je vais l'appeler.)
                   [pause]
                   Bintou! Kouamé is on the phone for you!
                   (Bintou ! Kouamé est au téléphone pour toi !)

  BINTOU         : Coming! [au téléphone] Speaking! Hi Kouamé, how are you?
                   (J'arrive ! [au téléphone] C'est moi ! Salut Kouamé, comment vas-tu ?)

  KOUAMÉ         : I'm fine, thank you. I'm calling because I don't understand
                   the English homework for tomorrow.
                   (Je vais bien, merci. J'appelle parce que je ne comprends pas
                   le devoir d'anglais pour demain.)

  BINTOU         : Oh, it's the dialogue exercise on page 52. You have to write
                   a telephone conversation between two students.
                   (Oh, c'est l'exercice de dialogue à la page 52. Tu dois écrire
                   une conversation téléphonique entre deux élèves.)

  KOUAMÉ         : I see. And what expressions should we use?
                   (Je vois. Et quelles expressions devons-nous utiliser ?)

  BINTOU         : You should use "Hello, this is... speaking", "Can I speak to...",
                   "Hold on" and "I'll call back later". Have you got your notebook?
                   (Tu devrais utiliser "Hello, this is... speaking", "Can I speak to...",
                   "Hold on" et "I'll call back later". Tu as ton cahier ?)

  KOUAMÉ         : Yes, I'm writing it down. Thank you so much, Bintou!
                   (Oui, je le note. Merci beaucoup, Bintou !)

  BINTOU         : No problem! If you have more questions, I'll be home all evening.
                   (Pas de problème ! Si tu as d'autres questions, je serai à la maison
                   toute la soirée.)

  KOUAMÉ         : Perfect. I'll call back later if I need help. Goodbye, Bintou!
                   (Parfait. Je rappellerai plus tard si j'ai besoin d'aide. Au revoir, Bintou !)

  BINTOU         : Goodbye, Kouamé! Good luck with your homework!
                   (Au revoir, Kouamé ! Bonne chance pour tes devoirs !)

  ─────────────────────────────────────────────────────

  PARTIE B — Opinion écrite

  "In my opinion, students should not be allowed to use their phones freely
  in class, but teachers should allow responsible use in certain situations.
  (À mon avis, les élèves ne devraient pas être autorisés à utiliser leur
  téléphone librement en classe, mais les enseignants devraient permettre
  un usage responsable dans certaines situations.)

  On the one hand, phones are a distraction : students who text or browse
  social media during lessons do not concentrate and get bad results.
  Moreover, the ringtone of a phone can disturb the whole class.
  (D'un côté, les téléphones sont une distraction : les élèves qui envoient
  des SMS ou naviguent sur les réseaux sociaux pendant les cours ne se
  concentrent pas et obtiennent de mauvais résultats. De plus, la sonnerie
  d'un téléphone peut déranger toute la classe.)

  However, phones can also be useful in class. Students should be allowed
  to use them to search for information, use educational apps or communicate
  with their teacher when necessary.
  (Cependant, les téléphones peuvent aussi être utiles en classe. Les élèves
  devraient être autorisés à les utiliser pour chercher des informations,
  utiliser des applications éducatives ou communiquer avec leur professeur
  si nécessaire.)

  In my opinion, schools should not ban phones completely. Instead, teachers
  should establish clear rules : phones must be silent during lessons and
  students should only use them when the teacher gives permission.
  This way, the phone becomes a learning tool, not a distraction.
  (À mon avis, les écoles ne devraient pas interdire totalement les téléphones.
  À la place, les enseignants devraient établir des règles claires : les téléphones
  doivent être silencieux pendant les cours et les élèves ne devraient les utiliser
  que lorsque le professeur le permet. Ainsi, le téléphone devient un outil
  d'apprentissage, pas une distraction.)"

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1

  Partie A :
    1. charger → chargeur
    2. ringtone → sonnerie
    3. voicemail → messagerie vocale
    4. to hang up → raccrocher
    5. landline → téléphone fixe
    6. SIM card → carte SIM
    7. to dial → composer un numéro
    8. network → réseau

  Partie B :
    1. décrocher → to pick up
    2. appel → call
    3. batterie → battery
    4. message texte → text message / SMS
    5. application → app / application

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2

  1 → A : "Can I speak to Awa, please?"
          (Puis-je parler à Awa, s.v.p. ?)
  2 → C : "I'll call back later."
          (Je rappellerai plus tard.)
  3 → B : "Hello? / Hello!"
          (Allô ?)
  4 → D : "Hold on, please."
          (Patientez s.v.p.)
  5 → E : "Can I leave a message?"
          (Puis-je laisser un message ?)
  6 → F : "I'll text you."
          (Je t'enverrai un SMS.)
  7 → G : "Hello, this is [name] speaking."
          (Allô, c'est [prénom] à l'appareil.)
  8 → H : "Speaking!"
          (C'est moi !)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3

  Q1 : What percentage of Ivorians own a mobile phone?
  R1 : More than 80% of Ivorians own a mobile phone.
       (Plus de 80% des Ivoiriens possèdent un téléphone mobile.)

  Q2 : Name three things people use their phones for in Côte d'Ivoire.
  R2 : People use their phones to call family members, send text messages
       and use Mobile Money services.
       (Les gens utilisent leur téléphone pour appeler des membres de leur famille,
       envoyer des messages texte et utiliser les services Mobile Money.)

  Q3 : Why have many schools banned phones in classrooms?
  R3 : Many schools have banned phones in classrooms because they distract students.
       (Beaucoup d'écoles ont interdit les téléphones en classe car ils distraient les élèves.)

  Q4 : How can phones be useful for students?
  R4 : Phones are useful for students because they allow them to receive course
       materials and communicate with their teachers.
       (Les téléphones sont utiles pour les élèves car ils leur permettent de recevoir
       des cours et de communiquer avec leurs professeurs.)

  Q5 : What does the government encourage young people to do?
  R5 : The government encourages young people to use their phones responsibly
       and to develop their digital skills for the future.
       (Le gouvernement encourage les jeunes à utiliser leur téléphone de façon
       responsable et à développer leurs compétences numériques pour l'avenir.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4

  Dialogue complété et traduit :

  BAMBA  : Hello?
           (Allô ?)
  AMINATA : Hello, THIS IS Aminata speaking. CAN I SPEAK TO Bamba, please?
            (Allô, c'est Aminata à l'appareil. Puis-je parler à Bamba, s.v.p. ?)
  BAMBA  : SPEAKING! Hi Aminata!
           (C'est moi ! Salut Aminata !)
  AMINATA : Hi Bamba! I tried to call you earlier but THE LINE IS BUSY.
            (Salut Bamba ! J'ai essayé de t'appeler plus tôt mais la ligne était occupée.)
  BAMBA  : Sorry, I was talking to my mum.
           (Désolé, je parlais avec ma maman.)
  AMINATA : No problem. HOLD ON please, I need to find my notebook.
            (Pas de problème. Patientez s.v.p., j'ai besoin de trouver mon cahier.)
  BAMBA  : Of course, take your time.
           (Bien sûr, prends ton temps.)
  AMINATA : Actually, I can't find it now. I'LL CALL BACK later, OK?
            (En fait, je ne le trouve pas maintenant. Je rappellerai plus tard, d'accord ?)
  BAMBA  : Sure! Goodbye!
           (Bien sûr ! Au revoir !)
  AMINATA : Goodbye!
            (Au revoir !)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5

  Exemple de texte :

  "The mobile phone is one of the most common tools among students in Côte d'Ivoire today.
  But is it a useful tool or a source of problems for students?
  (Le téléphone mobile est l'un des outils les plus répandus chez les élèves
  en Côte d'Ivoire aujourd'hui. Mais est-ce un outil utile ou une source de
  problèmes pour les élèves ?)

  On the one hand, the mobile phone has many advantages for students.
  First, it allows them to communicate quickly with their family and teachers.
  Moreover, students can use their phones to search for information online
  and access educational resources.
  In addition, Mobile Money services help students' families to send money easily
  in case of emergency.
  (D'un côté, le téléphone mobile a de nombreux avantages pour les élèves.
  Premièrement, il leur permet de communiquer rapidement avec leur famille
  et leurs professeurs. De plus, les élèves peuvent utiliser leur téléphone
  pour chercher des informations en ligne et accéder aux ressources éducatives.
  En outre, les services Mobile Money aident les familles des élèves à envoyer
  de l'argent facilement en cas d'urgence.)

  On the other hand, the mobile phone also has serious disadvantages.
  However, many students use their phones during class and do not concentrate.
  Moreover, excessive use of the phone can lead to addiction and sleep problems.
  Cyberbullying and online scams are also serious risks for young people.
  (D'un autre côté, le téléphone mobile a aussi de sérieux inconvénients.
  Cependant, beaucoup d'élèves utilisent leur téléphone en classe et ne se
  concentrent pas. De plus, l'utilisation excessive du téléphone peut entraîner
  une dépendance et des problèmes de sommeil. Le cyberharcèlement et les arnaques
  en ligne sont aussi des risques sérieux pour les jeunes.)

  In conclusion, the mobile phone is neither completely good nor completely bad.
  Students should use their phones responsibly — as a tool for learning and
  communication, not as a distraction. Schools should establish clear rules
  about phone use so that students can benefit from technology without suffering
  from its negative effects.
  (En conclusion, le téléphone mobile n'est ni complètement bon ni complètement mauvais.
  Les élèves devraient utiliser leur téléphone de façon responsable — comme un outil
  d'apprentissage et de communication, pas comme une distraction. Les écoles devraient
  établir des règles claires sur l'utilisation du téléphone pour que les élèves puissent
  bénéficier de la technologie sans souffrir de ses effets négatifs.)"

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "The telephone is the greatest nuisance among conveniences,
  the greatest convenience among nuisances."
  (Le téléphone est le plus grand inconvénient parmi les commodités,
  et la plus grande commodité parmi les inconvénients.)
  — Robert Lynd

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
