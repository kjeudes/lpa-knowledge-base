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

COURS = """✅ Français · Leçon 1 — Œuvre intégrale narrative · en cours de génération...

---

LPA
Lycée Personnel Autonome

Français · Leçon 1 · Classe de Seconde

L'ŒUVRE INTÉGRALE NARRATIVE
L'œuvre intégrale narrative

Compétence C1 Lecture · Construction du sens

Professeur Mariam KONATÉ · Coefficient 4
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  FRANÇAIS · Leçon 1 — L'ŒUVRE INTÉGRALE NARRATIVE
  Niveau : Seconde | Série : —
  Professeur : Mariam KONATÉ · Coefficient : 4
  Compétence : C1 Lecture · Construction du sens
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Ancrage ivoirien
    · Questions Harkness 1, 2 et 3
  Section 2 — La leçon
    · Phase 1 — Présentation & prérequis
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

  [Pas de capsule de reprise — première leçon de l'année]

  ANCRAGE IVOIRIEN
  ─────────────────
  Aminata, élève en Seconde au Lycée Moderne de Cocody, reçoit de sa tante un roman en cadeau :
  L'Enfant noir de Camara Laye. Elle l'ouvre et tombe sur ces mots : « J'étais enfant et je
  jouais près de la case de mon père. » Elle le repose aussitôt : « Ce n'est qu'une histoire. »
  Sa professeure lui répond le lendemain : « Toute grande histoire est une façon de comprendre
  le monde. »

  C'est exactement ce que tu vas apprendre dans cette leçon : lire une œuvre narrative, ce
  n'est pas simplement suivre une histoire — c'est décoder les choix de l'auteur, comprendre
  qui raconte, comment il raconte, et pourquoi.


▶ HARKNESS 1 — Qu'est-ce qu'un roman ?
  Q : Selon toi, quelle est la différence entre une histoire que l'on raconte oralement
      le soir au village et un roman écrit ?
  Guide :
    · Pense à la forme : oral vs écrit — quelles contraintes ?
    · Pense à l'auteur : il choisit chaque mot, chaque personnage
    · Pense au lecteur : il lit seul, en silence, à son rythme
  Corrigé :
    Le conte oral est improvisé, vivant, collectif : le griot s'adapte à son public en temps
    réel. Le roman, lui, est fixé par l'écriture. L'auteur a tout construit à l'avance : les
    personnages, l'ordre des événements, le point de vue. Le lecteur reçoit l'œuvre seul, dans
    sa tête. Mais les deux forment ont un point commun fondamental : raconter une histoire pour
    transmettre une vision du monde, une vérité humaine.

▶ HARKNESS 2 — Qui raconte ?
  Q : Dans L'Enfant noir, Camara Laye écrit « je ». Cela veut-il dire que l'auteur parle
      de lui-même ? L'auteur et le narrateur sont-ils toujours la même personne ?
  Guide :
    · Pense à un film où la voix off raconte l'histoire — cette voix est-elle celle du
      réalisateur ?
    · Un auteur peut inventer un "je" fictif
    · La notion d'autobiographie vs roman autobiographique
  Corrigé :
    Non, l'auteur et le narrateur ne sont pas toujours la même personne. L'auteur est la
    personne réelle qui écrit le livre. Le narrateur est la voix à l'intérieur du texte qui
    raconte l'histoire. Quand Camara Laye écrit « je », il crée un narrateur qui lui ressemble,
    mais le texte reste un roman. On parle alors de roman autobiographique : il y a une part de
    vérité vécue, mais aussi une construction littéraire. Dans un roman purement fictif, le
    « je » peut être un personnage entièrement inventé.

▶ HARKNESS 3 — À quoi sert la littérature ?
  Q : Pourquoi lire un roman ivoirien ou africain plutôt que de regarder simplement un film
      sur le même sujet ?
  Guide :
    · Le roman donne accès aux pensées intérieures des personnages
    · La langue littéraire crée des images et des émotions que le cinéma ne peut pas toujours
      reproduire
    · La littérature africaine parle d'un vécu, d'une mémoire, d'une identité
  Corrigé :
    Le roman offre quelque chose que le film ne peut pas donner : l'accès direct aux pensées,
    aux doutes, aux rêves d'un personnage. La langue littéraire, avec ses images et ses
    rythmes, crée une expérience intime et unique. De plus, lire un roman africain, c'est
    rencontrer une voix qui parle de nous, de notre histoire, de nos sociétés. C'est une façon
    de construire son identité et de comprendre le monde qui nous entoure.


────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis
───────────────────────────────────

Titre : L'œuvre intégrale narrative
Objectifs :
  · Identifier les composantes d'un texte narratif (narrateur, point de vue, structure)
  · Reconnaître les tonalités dominantes d'un récit
  · Lire une œuvre intégrale en appliquant une méthode de lecture méthodique

Prérequis nécessaires :
  · Savoir distinguer un texte narratif d'un texte descriptif ou argumentatif (3ème)
  · Connaître les notions de personnage, d'intrigue et de cadre spatio-temporel
  · Avoir une pratique minimale de la lecture de textes longs

┌─ À RETENIR ──────────────────────────────────────────┐
  En Seconde, lire une œuvre intégrale, c'est lire un
  texte EN ENTIER et l'analyser avec des outils précis.
  Tu ne lis plus pour « savoir ce qui se passe » :
  tu lis pour comprendre COMMENT c'est construit et
  POURQUOI l'auteur a fait ces choix.
  "En vue du BAC, retenez que la lecture méthodique
  est la compétence fondamentale de l'épreuve de Français."
└──────────────────────────────────────────────────────┘

Œuvre de référence pour cette leçon :
  L'Enfant noir · Camara Laye · 1953 · Éditions Plon
  Auteur : Guinéen · Langue : Français · Contexte : Afrique coloniale
  Thèmes : enfance · famille · tradition · modernité · exil
  Note : œuvre fondamentale de la littérature africaine francophone,
         régulièrement au programme DPFC CI


Phase 2 · Découverte guidée
────────────────────────────

  I. QU'EST-CE QU'UNE ŒUVRE NARRATIVE ?

  Un texte narratif est un texte qui raconte une histoire. L'œuvre narrative intégrale est
  un texte long — roman, récit, nouvelle — que l'on lit en entier. Elle met en scène des
  personnages dans un cadre précis (lieu, époque) et suit une intrigue qui évolue du début
  à la fin.

  ┌─ À RETENIR ──────────────────────────────────────────┐
    Une œuvre narrative se reconnaît à :
    · La présence de personnages nommés
    · Des événements qui s'enchaînent dans le temps
    · Un narrateur qui raconte (même si on ne le voit pas)
    · Des lieux et des temps précis ou suggérés
  └──────────────────────────────────────────────────────┘


  II. LE NARRATEUR ET LE POINT DE VUE

  Le narrateur est la voix qui raconte l'histoire à l'intérieur du texte. Il ne faut jamais
  le confondre avec l'auteur, qui est la personne réelle ayant écrit le livre.

  Il existe trois types de narrateurs et de points de vue :

  1 · Le narrateur à la PREMIÈRE PERSONNE (point de vue interne)
      → Il dit « je » · Il participe à l'histoire · Il raconte ce qu'il ressent
      → Exemple dans L'Enfant noir : « Je grandissais et peu à peu quittais mon enfance. »
      → Effet : le lecteur est au plus près des émotions du personnage

  2 · Le narrateur à la TROISIÈME PERSONNE — point de vue OMNISCIENT (point de vue externe total)
      → Il dit « il », « elle » · Il sait tout sur tous les personnages
      → Exemple fictif : « Kouadio ne savait pas encore que sa vie allait changer. »
      → Effet : le narrateur connaît même les pensées secrètes des personnages

  3 · Le narrateur à la TROISIÈME PERSONNE — point de vue EXTERNE
      → Il dit « il », « elle » · Il observe de l'extérieur, comme une caméra
      → Il ne rapporte que ce qui est visible, sans entrer dans les pensées
      → Effet : le lecteur doit deviner les émotions des personnages

  ⚠ ATTENTION
    Dans L'Enfant noir, le narrateur dit « je » — c'est un point de vue interne.
    Mais attention : Camara Laye (auteur réel) et le narrateur-enfant (voix du texte)
    ne sont PAS la même personne, même si l'œuvre est autobiographique.
    L'auteur a construit ce « je » avec des choix littéraires précis.


  III. LA STRUCTURE NARRATIVE

  Toute œuvre narrative suit une structure logique que l'on peut décomposer en cinq étapes.
  Cette structure s'appelle le SCHÉMA NARRATIF :

  Étape 1 · La situation initiale
    → Le monde est stable · Le lecteur découvre les personnages et le cadre
    → Exemple L'Enfant noir : la vie paisible dans la concession familiale à Kouroussa

  Étape 2 · L'élément perturbateur (ou déclencheur)
    → Un événement vient briser l'équilibre
    → Exemple : l'obligation scolaire qui éloigne l'enfant de sa famille

  Étape 3 · Les péripéties (ou actions)
    → Le personnage agit, surmonte des obstacles, évolue
    → Exemple : les années au lycée à Conakry, l'apprentissage, les amis

  Étape 4 · L'élément de résolution
    → La situation se stabilise, pour le meilleur ou pour le pire
    → Exemple : la décision de partir en France pour les études

  Étape 5 · La situation finale
    → Nouvel équilibre — souvent différent du début
    → Exemple : l'exil définitif, la nostalgie de l'Afrique

  ┌─ À RETENIR ──────────────────────────────────────────┐
    Schéma narratif = SI → EP → P → ER → SF
    Retiens cette formule : elle s'applique à tout récit,
    du conte de ta grand-mère au roman le plus complexe.
  └──────────────────────────────────────────────────────┘


  IV. LES TONALITÉS DOMINANTES DU RÉCIT NARRATIF

  Une tonalité est l'atmosphère générale qui se dégage d'un texte. Elle dépend des mots
  choisis, du rythme des phrases et de l'intention de l'auteur.

  Dans L'Enfant noir, on trouve principalement :

  Tonalité LYRIQUE
  → L'auteur exprime ses émotions avec intensité
  → Mots du sentiment, de la nostalgie, de l'amour filial
  → « Je n'aimais rien tant que de regarder mon père travailler l'or »

  Tonalité ÉPIQUE
  → Célébration des traditions, des ancêtres, des héros
  → Amplification · Vocabulaire fort · Scènes de rite et d'initiation

  Tonalité RÉALISTE
  → Description précise de la vie quotidienne africaine
  → Noms de lieux réels · Gestes concrets · Absence d'idéalisation excessive

  ⚠ ATTENTION
    Un même texte peut mêler plusieurs tonalités. C'est même la marque des grands romans.
    En commentaire de texte, tu devras identifier la tonalité DOMINANTE et la justifier
    avec des exemples précis tirés du texte.


Phase 3 · Schémas textuels
────────────────────────────

[SCHÉMA 1 — Le schéma narratif de L'Enfant noir]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Axe horizontal : ligne du temps de gauche à droite, nommée "Progression du récit"
  · 5 cases rectangulaires posées sur l'axe, reliées par des flèches →
  · Case 1 (gauche, fond vert clair)  : "Situation initiale — Vie à Kouroussa"
  · Case 2 (fond orange clair)        : "Élément perturbateur — La scolarité"
  · Case 3 (fond bleu clair, plus grande) : "Péripéties — Lycée de Conakry / Apprentissage"
  · Case 4 (fond violet clair)        : "Résolution — Départ pour la France"
  · Case 5 (droite, fond gris clair)  : "Situation finale — Exil et nostalgie"
  · Au-dessus de chaque flèche : une courbe montrant la tension dramatique (monte puis descend)
  · Légende en bas : SI = Situation Initiale · EP = Élément Perturbateur · P = Péripéties
                     ER = Élément de Résolution · SF = Situation Finale
Note générateur : draw.io ou Canva (forme "timeline horizontale")
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Carte mentale : Les outils de lecture narrative]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Nœud central (ovale, fond rouge) : "LECTURE NARRATIVE"
  · 3 branches principales rayonnant du centre :
    Branche 1 (haut gauche, fond bleu) : "LE NARRATEUR"
      → Sous-branches : "1ère personne / je" · "3ème personne / il" · "≠ auteur"
    Branche 2 (haut droite, fond vert) : "LE POINT DE VUE"
      → Sous-branches : "Interne" · "Omniscient" · "Externe"
    Branche 3 (bas centre, fond orange) : "LA STRUCTURE"
      → Sous-branches : "SI" · "EP" · "Péripéties" · "ER" · "SF"
  · Style : carte mentale avec lignes courbées, polices variées
Note générateur : Canva (modèle "Mind Map") ou XMind
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


Phase 4 · Définitions DPFC
────────────────────────────

◆ DÉFINITION DPFC — Narrateur
  La voix à l'intérieur du texte qui prend en charge le récit. Il peut participer à
  l'histoire (narrateur-personnage) ou la raconter de l'extérieur (narrateur-témoin
  ou narrateur omniscient). À ne pas confondre avec l'auteur.

◆ DÉFINITION DPFC — Point de vue narratif (ou focalisation)
  La perspective depuis laquelle l'histoire est racontée. On distingue :
  le point de vue interne (le narrateur est dans la tête d'un personnage),
  le point de vue omniscient (le narrateur sait tout sur tout le monde),
  le point de vue externe (le narrateur observe sans accès aux pensées).

◆ DÉFINITION DPFC — Schéma narratif
  Structure en cinq étapes organisant le déroulement d'un récit :
  situation initiale — élément perturbateur — péripéties — élément de résolution —
  situation finale.

◆ DÉFINITION DPFC — Tonalité littéraire
  L'atmosphère dominante qui se dégage d'un texte, produite par l'ensemble des
  choix stylistiques de l'auteur (lexique, rythme, figures de style). Exemples :
  tonalité lyrique, épique, réaliste, tragique, comique.

◆ DÉFINITION DPFC — Lecture méthodique
  Méthode d'analyse d'un texte littéraire consistant à dégager un sens global,
  à identifier les axes d'étude et à les justifier par des relevés précis dans le texte.

◆ DÉFINITION DPFC — Roman autobiographique
  Roman dans lequel l'auteur s'inspire de sa propre vie pour construire son récit,
  sans que le pacte de vérité soit totalement établi (contrairement à l'autobiographie).
  Exemple : L'Enfant noir de Camara Laye.


Phase 5 · Trace écrite + Analogies CI
───────────────────────────────────────

✎ TRACE ÉCRITE (à recopier)

  LEÇON 1 · L'ŒUVRE INTÉGRALE NARRATIVE
  Français · Seconde · Prof. Mariam KONATÉ

  I. LE NARRATEUR ET LE POINT DE VUE
  · Le narrateur = la voix qui raconte à l'intérieur du texte (≠ l'auteur)
  · Point de vue interne : narrateur dit « je » · accès aux émotions
  · Point de vue omniscient : narrateur dit « il/elle » · sait tout
  · Point de vue externe : narrateur dit « il/elle » · observe sans entrer dans les pensées

  II. LE SCHÉMA NARRATIF
  Situation initiale → Élément perturbateur → Péripéties → Élément de résolution → Situation finale
  Formule : SI → EP → P → ER → SF

  III. LES TONALITÉS DOMINANTES
  · Lyrique : émotions, sentiments, 1ère personne
  · Épique : exploits, célébration, amplification
  · Réaliste : description précise, ancrage réel, objectivité

  ŒUVRE DE RÉFÉRENCE : L'Enfant noir · Camara Laye · 1953
  → Narrateur à la 1ère personne · Point de vue interne
  → Tonalités dominantes : lyrique + réaliste + épique


~ ANALOGIE CI 1
  Le schéma narratif ressemble à la structure d'un match de football à l'épervain
  ou dans un quartier d'Abidjan :
  · Situation initiale = les deux équipes se mettent en place, tout est calme
  · Élément perturbateur = le coup de sifflet, un premier but contesté
  · Péripéties = les actions de jeu, les renversements de situation
  · Élément de résolution = le but décisif dans les dernières minutes
  · Situation finale = le résultat final et ses conséquences
  Comme dans un récit, chaque match a sa propre tension et sa propre histoire.

~ ANALOGIE CI 2
  Le point de vue narratif, c'est comme choisir d'où on regarde un griot raconter.
  · Si tu es assis juste à côté de lui et qu'il te parle directement → point de vue interne
  · Si tu es perché en haut d'un arbre et que tu vois tout le village en même temps → omniscient
  · Si tu regardes de loin, sans entendre ce qu'il dit → point de vue externe


Phase 5b · Synthèse
─────────────────────

5 POINTS ESSENTIELS À RETENIR :
  1 · Le narrateur est la voix du texte — il n'est jamais automatiquement l'auteur
  2 · Le point de vue (interne / omniscient / externe) conditionne ce que le lecteur sait
  3 · Le schéma narratif (SI → EP → P → ER → SF) s'applique à tout récit
  4 · La tonalité est l'atmosphère créée par les choix stylistiques de l'auteur
  5 · Lire une œuvre intégrale exige une méthode : observer, relever, interpréter

5 MOTS CLÉS + DÉFINITION COURTE :
  · Narrateur     : voix qui raconte à l'intérieur du texte
  · Focalisation  : angle de vision choisi par le narrateur
  · Schéma narratif : structure en 5 étapes de tout récit
  · Tonalité      : atmosphère dominante d'un texte
  · Lecture méthodique : analyse organisée et justifiée par des citations

3 ERREURS FRÉQUENTES + CORRECTION :
  Erreur 1 · « L'auteur dit que... »
    → Correction : dire « Le narrateur dit que... » ou « L'auteur montre que... »
      L'auteur écrit le livre ; le narrateur raconte l'histoire dans le texte.

  Erreur 2 · Confondre point de vue interne et omniscient
    → Correction : le point de vue interne entre dans la tête D'UN SEUL personnage.
      Le point de vue omniscient entre dans la tête de TOUS les personnages.

  Erreur 3 · Nommer une tonalité sans la justifier
    → Correction : en commentaire de texte, toute tonalité doit être prouvée par
      au moins un relevé précis du texte (un mot, une expression, une figure de style).


Phase 6 · Exercices guidés
────────────────────────────

◉ EXERCICE GUIDÉ 1 — Identifier le narrateur et le point de vue · Notion ciblée : narrateur / focalisation

  Énoncé :
  Lis l'extrait suivant (inspiré de L'Enfant noir · Camara Laye) puis réponds aux questions.

  « Mon père avait ce qu'on appelle le don : certains jours, quand il travaillait l'or, je
  voyais le serpent noir glisser vers l'enclume. Je savais alors que mon père allait réussir
  un chef-d'œuvre. Ma mère, elle, ne venait jamais dans cette partie de la case. »

  Question A · Qui raconte ? Relevez un indice dans le texte.
  Question B · Quel est le point de vue narratif ? Justifiez.
  Question C · Quelle est la tonalité dominante de cet extrait ? Justifiez avec un mot du texte.

  MÉTHODE
  Étape 1 · Repère les pronoms personnels dominants dans le texte
    → « Mon père », « je voyais », « je savais » → le narrateur dit « je » → il participe à l'histoire
  Étape 2 · Identifie le type de narrateur
    → Narrateur à la première personne → il est personnage de l'histoire
    → Indice textuel : « je savais alors » — accès à ses propres pensées
  Étape 3 · Détermine le point de vue
    → Le narrateur voit ce que voit son père, ressent ce qu'il ressent LUI
    → Il n'entre pas dans les pensées de sa mère (« Ma mère ne venait jamais ») — information externe
    → Point de vue : interne (centré sur le narrateur-enfant)
  Étape 4 · Identifie la tonalité
    → Champ lexical du merveilleux et de l'admiration : « don » · « chef-d'œuvre »
    → Tonalité lyrique (admiration filiale, émotion) mêlée de tonalité épique (le père comme héros)

  ✔ Vérification : ai-je cité un mot du texte pour chaque réponse ? Oui.
  ✔ Conclusion :
    A · C'est l'enfant-narrateur qui raconte. Indice : « je voyais »
    B · Point de vue interne — le narrateur partage ses propres perceptions et connaissances
    C · Tonalité lyrique dominante — justifiée par « chef-d'œuvre » et l'admiration exprimée
  ✔ Ce que cet exercice mobilise :
    · Identification du narrateur (1ère personne)
    · Distinction auteur / narrateur
    · Détermination du point de vue narratif (interne)
    · Identification et justification d'une tonalité littéraire
    · Méthode du relevé textuel


◉ EXERCICE GUIDÉ 2 — Reconnaître le schéma narratif · Notion ciblée : structure narrative

  Énoncé :
  Lis ce résumé d'un court récit ivoirien fictif, puis identifie les cinq étapes du schéma narratif.

  « Konan vivait paisiblement avec sa famille à Bouaké et apprenait le métier de son père,
  tisserand. Un matin, un incendie détruisit l'atelier familial. Konan dut alors partir à
  Abidjan pour chercher du travail. Il traversa de nombreuses difficultés : vol de ses
  économies, maladie, solitude. Après deux années d'efforts, il trouva un emploi stable dans
  une entreprise de textile. Il rentra au village avec assez d'argent pour reconstruire
  l'atelier et financer les études de sa sœur. »

  Question : Identifie et nomme les cinq étapes du schéma narratif. Cite un élément du texte pour chaque étape.

  MÉTHODE
  Étape 1 · Cherche la situation stable au début
    → « Konan vivait paisiblement... apprenait le métier de son père »
    → Situation initiale : vie calme, apprentissage familial à Bouaké

  Étape 2 · Cherche l'événement qui brise cet équilibre
    → « Un matin, un incendie détruisit l'atelier familial »
    → Élément perturbateur : l'incendie

  Étape 3 · Cherche les obstacles et actions du personnage
    → « vol de ses économies, maladie, solitude »
    → Péripéties : les épreuves subies à Abidjan

  Étape 4 · Cherche le moment où la situation se résout
    → « il trouva un emploi stable dans une entreprise de textile »
    → Élément de résolution : l'emploi obtenu

  Étape 5 · Cherche le nouvel équilibre à la fin
    → « Il rentra au village... reconstruire l'atelier et financer les études de sa sœur »
    → Situation finale : retour, reconstruction, projet familial

  ✔ Vérification : cinq étapes identifiées, chacune justifiée par une citation. Oui.
  ✔ Conclusion : Le schéma narratif est complet et cohérent. Ce récit suit la structure classique.
  ✔ Ce que cet exercice mobilise :
    · Identification des cinq étapes du schéma narratif
    · Justification par citation du texte
    · Compréhension de la progression dramatique d'un récit
    · Application de la méthode à un contexte ivoirien


────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Le narrateur dans un texte · Notions travaillées : narrateur, point de vue

  Lis l'extrait suivant :
  « Adjoua regardait la mer pour la première fois. Elle n'avait jamais quitté Yamoussoukro.
  Le bruit des vagues l'étourdit. Elle pensa à sa mère restée au village et sentit quelque
  chose se serrer dans sa poitrine. »

  Questions :
  A · À quelle personne grammaticale ce texte est-il écrit ?
  B · Qui raconte l'histoire ? Peut-on dire que c'est l'auteur ?
  C · Quel est le point de vue narratif ? Justifiez avec un élément du texte.
  D · Que ressent Adjoua ? Comment le savons-nous ?

  GUIDE
  Étape 1 · Relève tous les pronoms et identifie la personne grammaticale dominante
  Étape 2 · Distingue le narrateur de l'auteur
  Étape 3 · Vérifie si le narrateur entre dans les pensées du personnage ou non
  Étape 4 · Identifie ce que le narrateur révèle des émotions d'Adjoua et comment


◎ EXERCICE 2 — Schéma narratif complet · Notions travaillées : structure narrative, schéma en 5 étapes

  Voici cinq phrases qui résument un récit, mais dans le désordre.
  Remets-les dans l'ordre logique du schéma narratif et nomme chaque étape.

  Phrase A : « Après des mois d'entraînement épuisant, Mamadou remporta finalement la compétition nationale. »
  Phrase B : « Mamadou était un jeune footballeur talentueux de Grand-Bassam, admiré dans tout son quartier. »
  Phrase C : « Il décida alors de tout recommencer depuis le début avec un nouvel entraîneur. »
  Phrase D : « Une grave blessure au genou l'écarta du terrain pendant six mois. »
  Phrase E : « Il fut sélectionné pour représenter la Côte d'Ivoire dans un tournoi africain. »

  GUIDE
  Étape 1 · Identifie la phrase qui présente un monde stable (= SI)
  Étape 2 · Identifie la phrase qui brise cet équilibre (= EP)
  Étape 3 · Identifie la ou les phrases d'action et d'obstacle (= P)
  Étape 4 · Identifie la phrase où la situation se dénoue (= ER)
  Étape 5 · Identifie la phrase qui montre un nouvel équilibre (= SF)


◎ EXERCICE 3 — Les tonalités · Notions travaillées : tonalités littéraires, relevé lexical

  Lis ces trois courts extraits et identifie pour chacun la tonalité dominante.
  Justifie ta réponse en relevant un mot ou une expression du texte.

  Extrait 1 :
  « Le vieux Sékou leva son coupe-coupe au-dessus de la bête rugissante.
  D'un seul coup, majestueux et précis, il abattit l'ennemi. Le village entier cria
  sa victoire. »

  Extrait 2 :
  « Fatou écrivit en pleurant : "Je pense à toi chaque nuit. Les étoiles au-dessus
  du plateau me rappellent tes yeux. Mon cœur est vide depuis ton départ." »

  Extrait 3 :
  « La case mesurait environ quatre mètres carrés. Le sol de terre battue
  était propre. Un bol de riz, trois bouteilles et une natte constituaient
  tout le mobilier. »

  GUIDE
  Étape 1 · Lis chaque extrait et note le sentiment général qu'il provoque
  Étape 2 · Repère le champ lexical dominant (mots de l'action, des émotions, de la description)
  Étape 3 · Nomme la tonalité (épique / lyrique / réaliste / tragique / comique) et justifie


◎ EXERCICE 4 — Auteur vs narrateur · Notions travaillées : distinction auteur / narrateur, roman autobiographique

  Réponds aux questions suivantes en rédigeant des phrases complètes.

  A · Quelle est la différence entre un auteur et un narrateur ?
  B · Dans L'Enfant noir, Camara Laye écrit à la première personne. Cela fait-il de ce roman
      une autobiographie ? Explique la différence entre autobiographie et roman autobiographique.
  C · Donne un exemple de situation où le narrateur et l'auteur sont clairement différents.

  GUIDE
  Étape 1 · Définis auteur puis narrateur avec tes propres mots
  Étape 2 · Rappelle ce que tu sais de L'Enfant noir (qui parle ? est-ce une histoire vraie ?)
  Étape 3 · Réfléchis : qu'est-ce qui distingue une autobiographie d'un roman autobiographique ?
  Étape 4 · Invente un exemple simple pour illustrer la différence auteur / narrateur


◎ EXERCICE 5 — Lecture méthodique d'un extrait · Notions travaillées : lecture méthodique, ensemble des outils de la leçon

  Lis l'extrait suivant avec attention :

  « J'avais quinze ans quand je quittai Abengourou pour Abidjan. Ma mère m'avait préparé
  un repas d'attiéké et de poisson braisé — le dernier repas du pays. Dans le car bondé,
  les larmes coulaient sur mes joues sans que je les retienne. Je savais que cette route
  était un chemin sans retour. Abidjan m'attendait, grande et indifférente. »

  Questions :
  A · Qui raconte ? Quel est le point de vue narratif ?
  B · À quelle étape du schéma narratif cet extrait correspond-il ? Justifiez.
  C · Identifiez deux tonalités présentes. Justifiez chacune par un relevé.
  D · Relevez une expression qui montre que le narrateur a accès à ses propres pensées.

  GUIDE
  Étape 1 · Repère les pronoms et identifie le narrateur
  Étape 2 · Situe cet extrait dans une histoire possible : début / milieu / fin ?
  Étape 3 · Note les mots qui expriment des émotions ET les mots qui décrivent avec précision
  Étape 4 · Cherche une phrase qui révèle une pensée intérieure du narrateur


◈ DEVOIR MAISON — Le récit de ma propre « rupture »
  Durée conseillée : 45 minutes

  Énoncé :
  Rédige un court récit à la première personne (15 à 20 lignes) dans lequel tu racontes
  un moment de ta vie où quelque chose a changé (un déménagement, une rentrée scolaire,
  la perte d'un ami, un voyage, etc.).
  Ton récit doit respecter les contraintes suivantes :

  Contrainte 1 · Utiliser le point de vue interne (écrire à la 1ère personne : « je »)
  Contrainte 2 · Faire apparaître clairement au moins 3 étapes du schéma narratif
                 (situation initiale, élément perturbateur, péripétie ou résolution)
  Contrainte 3 · Créer une tonalité dominante (lyrique ou réaliste) et l'entretenir
                 par le choix des mots
  Contrainte 4 · Utiliser au moins un nom, un lieu ou un élément ivoirien concret
  Contrainte 5 · Rédiger en phrases complètes, avec introduction et conclusion

  GUIDE (sans corrigé)
  Étape 1 · Choisis le souvenir ou l'événement que tu vas raconter
  Étape 2 · Identifie la situation de départ (comment était ta vie avant ?)
  Étape 3 · Identifie l'événement qui a tout changé (quoi ? quand ? comment ?)
  Étape 4 · Rédige ce que tu as ressenti et vécu (les péripéties de ton émotion)
  Étape 5 · Conclus en montrant comment tu as changé ou ce que tu en retiens aujourd'hui
  Étape 6 · Relis et vérifie que les 5 contraintes sont respectées


────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON
  (Exemple de production conforme aux contraintes — à titre indicatif)

  Étape 1 · Situation initiale
  « Je vivais à Daloa depuis ma naissance. Notre maison était au bout d'une rue en terre
  rouge, près du marché. Chaque matin, l'odeur du pain chaud et des beignets m'accompagnait
  jusqu'à l'école du quartier. Je ne connaissais rien d'autre que cette ville et ces bruits. »
  → Point de vue interne respecté · Ancrage ivoirien (Daloa, beignets) · Tonalité réaliste

  Étape 2 · Élément perturbateur
  « Un soir de septembre, mon père nous annonça que nous partions pour Abidjan dans trois jours.
  Trois jours seulement. Je regardai ma chambre comme si je la voyais pour la première fois. »
  → Élément perturbateur clairement identifiable · Début de tonalité lyrique (nostalgie)

  Étape 3 · Péripétie / Résolution
  « Le jour du départ, ma voisine Aya m'apporta une petite calebasse peinte en souvenir.
  Je la glissai dans mon sac sans un mot. Dans le bus, je regardai défiler les cocotiers,
  les cases, les enfants qui couraient. Je savais que ce paysage appartenait désormais
  à mon passé. »
  → Péripétie émotionnelle · Tonalité lyrique maintenue · Accès aux pensées du narrateur

  Étape 4 · Situation finale (conclusion)
  « Abidjan m'a appris à grandir vite. Mais chaque fois que j'entends le bruit d'un marché
  le matin, je retourne, une seconde, dans ma rue rouge de Daloa. »
  → Clôture réflexive · Bilan personnel · Ouverture

  Résultat final : récit de 17 lignes · 5 contraintes respectées · Note estimée : 18/20


✔ CORRIGÉ — EXERCICE 1
  Étape 1 · Pronoms dominants : « elle », « sa », « la » → 3ème personne du singulier
  Étape 2 · Le narrateur n'est pas Adjoua — c'est une voix extérieure qui la raconte.
             Ce n'est pas l'auteur non plus : l'auteur écrit le texte, le narrateur
             choisit de raconter l'histoire d'Adjoua.
  Étape 3 · Le narrateur entre dans les pensées d'Adjoua : « Elle pensa à sa mère »
             → accès aux pensées intérieures → point de vue omniscient
             (ou interne centré sur Adjoua — les deux réponses sont acceptables)
  Étape 4 · Adjoua ressent de la nostalgie et de la tristesse. On le sait grâce à :
             « sentit quelque chose se serrer dans sa poitrine » (sensation physique intérieure)
             et « Elle pensa à sa mère » (accès aux pensées)

  Résultat final :
  A · 3ème personne
  B · Un narrateur extérieur à l'histoire (non, ce n'est pas l'auteur)
  C · Point de vue omniscient — justifié par « Elle pensa à sa mère »
  D · Nostalgie et tristesse — révélées par « sentit quelque chose se serrer dans sa poitrine »


✔ CORRIGÉ — EXERCICE 2
  Étape 1 · Phrase B → Situation initiale : monde stable, Mamadou est admiré
  Étape 2 · Phrase D → Élément perturbateur : blessure au genou
  Étape 3 · Phrase C → Péripétie : décision de recommencer avec un nouvel entraîneur
  Étape 4 · Phrase A → Élément de résolution : victoire à la compétition nationale
  Étape 5 · Phrase E → Situation finale : sélection nationale (nouvel équilibre, nouveau statut)

  Résultat final : Ordre correct → B · D · C · A · E
  SI = B · EP = D · P = C · ER = A · SF = E


✔ CORRIGÉ — EXERCICE 3
  Étape 1 · Extrait 1 → Tonalité ÉPIQUE
    Justification : « D'un seul coup, majestueux et précis » — amplification du geste héroïque
    + « Le village entier cria sa victoire » — célébration collective

  Étape 2 · Extrait 2 → Tonalité LYRIQUE
    Justification : « les larmes », « Mon cœur est vide » — expression intense des émotions
    + marques de la 1ère personne (« je pense », « mon cœur »)

  Étape 3 · Extrait 3 → Tonalité RÉALISTE
    Justification : « environ quatre mètres carrés », « sol de terre battue » —
    description précise, chiffrée, sans embellissement ni émotion

  Résultat final : Extrait 1 = épique · Extrait 2 = lyrique · Extrait 3 = réaliste


✔ CORRIGÉ — EXERCICE 4
  Étape 1 ·
    Auteur : personne réelle qui écrit l'œuvre, en dehors du texte.
    Narrateur : voix fictive à l'intérieur du texte qui prend en charge le récit.
    Un auteur peut inventer un narrateur très différent de lui-même.

  Étape 2 ·
    L'Enfant noir : Camara Laye écrit à la première personne, mais c'est un ROMAN.
    L'auteur a construit littérairement cette voix, même si elle s'inspire de sa vie.
    Ce n'est donc pas une autobiographie au sens strict.

  Étape 3 ·
    Autobiographie : l'auteur s'engage à raconter sa VIE RÉELLE · pacte avec le lecteur ·
    exemple : Si la lune pouvait parler de Marie NDiaye (autobiographie stricte)
    Roman autobiographique : s'inspire de la vie réelle MAIS avec liberté fictive ·
    pas de pacte de vérité total · exemple : L'Enfant noir de Camara Laye

  Étape 4 ·
    Exemple : Un auteur ivoirien écrit un roman où le narrateur s'appelle Koffi et vit à
    San Pedro. L'auteur, lui, s'appelle Jean-Pierre et vit à Paris. Narrateur ≠ auteur.

  Résultat final :
  A · L'auteur crée le texte depuis l'extérieur ; le narrateur existe dans le texte.
  B · C'est un roman autobiographique : Laye s'inspire de sa vie, mais avec une liberté
      fictionnelle. L'autobiographie exige un pacte de vérité total.
  C · Exemple au choix, du moment que auteur ≠ narrateur est clairement illustré.


✔ CORRIGÉ — EXERCICE 5
  Étape 1 ·
  A · Narrateur à la 1ère personne (« j'avais », « je quittai », « je savais »)
     Point de vue interne — accès aux émotions et pensées du narrateur

  Étape 2 ·
  B · Cet extrait correspond à l'ÉLÉMENT PERTURBATEUR / début des PÉRIPÉTIES
     Justification : le départ est une rupture (quitter Abengourou), la vie bascule

  Étape 3 ·
  C · Tonalité LYRIQUE : « les larmes coulaient sur mes joues », « Mon cœur »,
      expression des émotions intenses
     Tonalité RÉALISTE : « le car bondé », « Abidjan m'attendait, grande et indifférente »
      description précise et sans idéalisation

  Étape 4 ·
  D · « Je savais que cette route était un chemin sans retour »
     → Pensée intérieure du narrateur · accès direct à sa conscience · confirme le point de vue interne

  Résultat final :
  A · Narrateur 1ère personne · point de vue interne
  B · Élément perturbateur / début des péripéties (le départ de la ville natale)
  C · Lyrique (les larmes) + Réaliste (le car bondé, Abidjan indifférente)
  D · « Je savais que cette route était un chemin sans retour »


════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  « Un lecteur vit mille vies avant de mourir. Celui qui ne lit jamais n'en vit qu'une. »
  (George R.R. Martin)

  Prof. Mariam KONATÉ · Français · Seconde · DPFC/MENET-FP CI · 2025-2026

════════════════════════════════════════════════════════
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
