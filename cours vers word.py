# ============================================================
#  CONVERTISSEUR DE COURS → WORD  (LPA / DPFC Côte d'Ivoire)
#  Google Colab — python-docx
#
#  UTILISATION :
#  1. Installe : !pip install python-docx -q
#  2. Colle ton cours dans la variable COURS ci-dessous
#  3. Exécute la cellule → le fichier Word se télécharge
# ============================================================

# ↓↓↓  COLLE TON COURS ICI  ↓↓↓
COURS = """════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 14 — Village life (Vie au village)
  Niveau : 3ème | Compétence : C5 — Écoute simple
  Professeur : Evelyne ASSI
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  1. Avant la leçon
  2. La leçon
  3. Après la leçon
  4. Corrigés complets

────────────────────────────────────────────────────────
SECTION 1 — AVANT LA LEÇON
────────────────────────────────────────────────────────

[CAPSULE DE REPRISE — Leçon 13 : City life]

  5 points clés de la leçon précédente :
  Point 1 · La vie en ville (city life) a des avantages :
            emplois (jobs), services, modernité, divertissements.
  Point 2 · La vie en ville a des inconvénients :
            bruit (noise), pollution, embouteillages
            (traffic jams), loyers chers (expensive rent).
  Point 3 · "There is" s'utilise avec un nom SINGULIER ;
            "There are" avec un nom PLURIEL.
  Point 4 · Comparatif des adjectifs courts (1 syllabe) :
            adj + -ER + than → big → bigger than.
  Point 5 · Comparatif des adjectifs longs (2+ syllabes) :
            more + adj + than → expensive → more expensive than.
            ⚠ Jamais "more bigger" — erreur grave.

  3 questions flash :
  Q1 · Complète : "There ________ many schools in Abidjan."
       → There ARE many schools in Abidjan.
  Q2 · Forme le comparatif de "hot".
       → hotter than (CVC → double T)
  Q3 · Traduis : "The city is more polluted than the village."
       → La ville est plus polluée que le village.

  Lien avec la leçon du jour :
  En Leçon 13, tu as décrit la vie en VILLE.
  Aujourd'hui tu décris son opposé : la vie au VILLAGE.
  Tu vas comparer les deux milieux, apprendre à exprimer
  des habitudes rurales, et utiliser le superlatif
  pour désigner ce qui est le plus ou le moins important.

────────────────────────────────────────────────────────

[ANCRAGE IVOIRIEN]

  En Côte d'Ivoire, plus de 40 % de la population vit
  encore en zone rurale. Des villages comme Tiébissou,
  Daoukro, Katiola ou Boundiali concentrent des communautés
  agricoles qui cultivent le cacao, le café, l'igname
  et le riz. La vie au village est rythmée par les saisons,
  les cérémonies traditionnelles et le travail de la terre.
  Les femmes jouent un rôle central : elles cultivent,
  élèvent les enfants et maintiennent les traditions.
  Mais le village manque souvent d'hôpitaux, d'électricité
  permanente et d'accès à Internet. C'est l'une des raisons
  de l'exode rural vers Abidjan.

  Lien DPFC : Environnement · agriculture · civisme ·
  valorisation du patrimoine culturel ivoirien.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — Qu'est-ce qu'un village ?
  Q : "How is a village different from a city?"
      (En quoi un village est-il différent d'une ville ?)
  Guide :
    Piste 1 · Pense au nombre d'habitants.
    Piste 2 · Pense aux activités principales des habitants.
    Piste 3 · Pense aux services disponibles ou absents.
  Corrigé :
    A village is a small community where fewer people live,
    usually in a rural area. (Un village est une petite
    communauté où moins de personnes vivent, généralement
    en zone rurale.)
    People in villages usually work in agriculture :
    they grow crops like cocoa, yam and rice.
    (Les gens au village travaillent généralement dans
    l'agriculture : ils cultivent des produits comme
    le cacao, l'igname et le riz.)
    Unlike cities, villages often have fewer services :
    no hospitals, limited electricity, no Internet.
    (Contrairement aux villes, les villages ont souvent
    moins de services : pas d'hôpital, électricité limitée,
    pas d'Internet.)

▶ HARKNESS 2 — Les avantages de la vie au village
  Q : "Why do some people prefer village life to city life?"
      (Pourquoi certaines personnes préfèrent-elles
      la vie au village à la vie en ville ?)
  Guide :
    Piste 1 · Pense à la tranquillité et à la nature.
    Piste 2 · Pense aux liens communautaires et familiaux.
    Piste 3 · Pense au coût de la vie.
  Corrigé :
    Some people prefer village life because it is quieter
    and less stressful than city life.
    (Certaines personnes préfèrent la vie au village
    parce qu'elle est plus calme et moins stressante
    que la vie en ville.)
    In villages, people live close to nature and eat fresh
    food from their own farms.
    (Au village, les gens vivent près de la nature et
    mangent des aliments frais de leurs propres champs.)
    Community life is stronger : neighbours help each other
    and family ties are closer.
    (La vie communautaire est plus forte : les voisins
    s'entraident et les liens familiaux sont plus étroits.)
    Also, the cost of living is lower in the village.
    (De plus, le coût de la vie est plus bas au village.)

▶ HARKNESS 3 — Le superlatif
  Q : "How do you say that something is THE MOST
       or THE LEAST in English?"
      (Comment dit-on que quelque chose est LE PLUS
       ou LE MOINS en anglais ?)
  Guide :
    Piste 1 · Quelle est la différence entre
              "bigger" et "the biggest" ?
    Piste 2 · Comment forme-t-on le superlatif
              d'un adjectif long ?
    Piste 3 · Quel article utilise-t-on toujours
              avec le superlatif ?
  Corrigé :
    The superlative expresses the highest or lowest degree
    among three or more things. (Le superlatif exprime
    le degré le plus élevé ou le plus bas parmi
    trois éléments ou plus.)
    Short adjectives : the + adj + -EST
    → big → the biggest (le/la plus grand[e])
    → fast → the fastest (le/la plus rapide)
    Long adjectives : the most + adj
    → beautiful → the most beautiful (le/la plus beau/belle)
    → expensive → the most expensive (le/la plus cher/chère)
    We ALWAYS use "the" before a superlative.
    (On utilise TOUJOURS "the" avant un superlatif.)
    Example : "Abidjan is the biggest city in Côte d'Ivoire."
    (Abidjan est la plus grande ville de Côte d'Ivoire.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis
────────────────────────────────────

  Titre : Village life (La vie au village)
  Compétence : C5 — Écoute simple
  Professeur Evelyne ASSI : "La vie au village, c'est
  la vie de millions d'Ivoiriens. Parlons-en en anglais —
  avec fierté et précision !"

  Objectifs de la leçon :
  → Décrire la vie au village en anglais
  → Connaître le vocabulaire des activités et lieux ruraux
  → Exprimer les avantages et inconvénients de la vie rurale
  → Utiliser le superlatif correctement
  → Comparer la vie en ville et au village à l'écrit

  Prérequis nécessaires :
  → Vocabulaire de la ville (Leçon 13)
  → There is / There are (Leçon 13)
  → Le comparatif (Leçon 13)
  → Le Present Simple (habitudes)

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée
────────────────────────────

  LIS CE TEXTE :

  ~~~
  My name is Koné. I come from Katiola, a small village
  in the north of Côte d'Ivoire. Life in my village is
  very different from life in Abidjan.

  In Katiola, people wake up early. Farmers go to their
  fields to grow yam, maize and millet. Women fetch water
  from the river and cook traditional meals. Children walk
  to school every morning. In the evening, families sit
  together under the big mango tree and share stories.

  The village is the quietest place I know. The air is
  the freshest and the nights are the most peaceful.
  However, the nearest hospital is fifty kilometres away,
  and there is no Internet connection. For many young people,
  the village is the hardest place to find a job.
  ~~~

  TRADUCTION COMPLÈTE :
  Je m'appelle Koné. Je viens de Katiola, un petit village
  du nord de la Côte d'Ivoire. La vie dans mon village est
  très différente de la vie à Abidjan.
  À Katiola, les gens se lèvent tôt. Les agriculteurs vont
  dans leurs champs pour cultiver l'igname, le maïs et
  le mil. Les femmes vont chercher de l'eau à la rivière
  et cuisinent des repas traditionnels. Les enfants marchent
  jusqu'à l'école chaque matin. Le soir, les familles
  s'assoient ensemble sous le grand manguier et partagent
  des histoires.
  Le village est l'endroit le plus calme que je connaisse.
  L'air est le plus frais et les nuits sont les plus
  paisibles. Cependant, l'hôpital le plus proche est à
  cinquante kilomètres, et il n'y a pas de connexion
  Internet. Pour beaucoup de jeunes, le village est
  l'endroit le plus difficile pour trouver un emploi.

┌─ À RETENIR — LE SUPERLATIF ─────────────────────────┐
  Le superlatif exprime le DEGRÉ MAXIMUM ou MINIMUM
  parmi trois éléments ou plus.
  On utilise TOUJOURS "the" devant le superlatif.

  Adjectif COURT (1 syllabe) → the + adj + -EST
    quiet  → the quietest   (le/la plus calme)
    fresh  → the freshest   (le/la plus frais/fraîche)
    hard   → the hardest    (le/la plus difficile)
    ⚠ CVC : big → the biggest (double G)
    ⚠ hot → the hottest (double T)

  Adjectif LONG (2 syllabes ou plus) → the most + adj
    peaceful   → the most peaceful   (le/la plus paisible)
    expensive  → the most expensive  (le/la plus cher/chère)
    beautiful  → the most beautiful  (le/la plus beau/belle)

  Forme négative : the least + adj (le/la moins…)
    → the least expensive (le/la moins cher/chère)

  Exemples ivoiriens :
  "Katiola is the smallest town in the region."
  (Katiola est la plus petite ville de la région.)
  "The mango tree is the tallest tree in the village."
  (Le manguier est le plus grand arbre du village.)
└──────────────────────────────────────────────────────┘

┌─ À RETENIR — HABITUDES RURALES (Present Simple) ────┐
  On utilise le Present Simple pour décrire des habitudes
  régulières et des vérités générales.
  Structure : Subject + verb (base form) + complement.
  Attention : He/She/It → verbe + -S

  Exemples :
  · Farmers grow yam and maize.
    (Les agriculteurs cultivent l'igname et le maïs.)
  · Women fetch water from the river.
    (Les femmes vont chercher de l'eau à la rivière.)
  · Children walk to school every morning.
    (Les enfants marchent jusqu'à l'école chaque matin.)
  · The family sits together in the evening.
    (La famille s'assoit ensemble le soir.)
└──────────────────────────────────────────────────────┘

⚠ ATTENTION — Pièges et faux amis
  · "village" se prononce /ˈvɪl.ɪdʒ/ — VI-lidj
    PAS "vilaghe" ni "villaj".
  · "field" = champ (agricole) — PAS "field" au sens
    de terrain de sport uniquement. Ici : champ de culture.
  · "fetch" = aller chercher (et ramener)
    PAS juste "chercher". "Fetch water" = aller chercher
    de l'eau ET la ramener.
  · "crop" = culture agricole / récolte
    PAS "crope" ni "crop" au sens de photo recadrée.
  · "mango tree" = manguier (l'arbre entier)
    "mango" seul = la mangue (le fruit).
  · "peaceful" se prononce /ˈpiːs.fəl/ — PIISS-feul
    PAS "pi-si-ful".

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels
───────────────────────────

[SCHÉMA 1 — Carte mentale : Village Life Vocabulary]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  • Centre (grand ovale vert) : VILLAGE LIFE / VIE AU VILLAGE
  • Branche 1 (vers le haut, marron) : PLACES / LIEUX
    → field (champ)
    → river (rivière)
    → mango tree (manguier)
    → well (puits)
    → path (sentier)
  • Branche 2 (vers la droite, orange) : ACTIVITIES / ACTIVITÉS
    → grow crops (cultiver)
    → fetch water (chercher de l'eau)
    → cook (cuisiner)
    → walk to school (marcher jusqu'à l'école)
    → share stories (partager des histoires)
  • Branche 3 (vers le bas, rouge) : PROBLEMS / PROBLÈMES
    → no hospital (pas d'hôpital)
    → no Internet (pas d'Internet)
    → few jobs (peu d'emplois)
    → limited electricity (électricité limitée)
  • Branche 4 (vers la gauche, verte) : ADVANTAGES / AVANTAGES
    → quiet (calme)
    → fresh air (air frais)
    → peaceful (paisible)
    → cheap (bon marché)
    → strong community (communauté forte)
  • Traductions en français sous chaque terme anglais
Note générateur : Canva (mind map) ou MindMeister
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Tableau : Comparatif vs Superlatif]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  • Titre en haut (fond bleu foncé, texte blanc) :
    COMPARATIVE vs SUPERLATIVE

  • Tableau à 4 colonnes :
    Col 1 : ADJECTIF (fond gris clair)
    Col 2 : COMPARATIF (fond jaune)
    Col 3 : SUPERLATIF (fond orange)
    Col 4 : TRADUCTION (fond vert clair)

    Ligne 1 :
    quiet | quieter than | the quietest | plus calme / le plus calme

    Ligne 2 :
    big | bigger than | the biggest | plus grand / le plus grand
    ⚠ (CVC → double G)

    Ligne 3 :
    peaceful | more peaceful than | the most peaceful |
    plus paisible / le plus paisible

    Ligne 4 :
    expensive | more expensive than | the most expensive |
    plus cher / le plus cher

  • Encadré rouge en bas :
    ⚠ TOUJOURS "the" devant le superlatif
    ⚠ JAMAIS "the most quietest" — erreur grave

Note générateur : Canva (tableau) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC
───────────────────────────

◆ DÉFINITION DPFC — Village
  A village is a small community in a rural area where
  people live, usually working in agriculture and maintaining
  strong traditional and community ties.
  (Un village est une petite communauté en zone rurale
  où les gens vivent, travaillant généralement dans
  l'agriculture et maintenant de forts liens traditionnels
  et communautaires.)

◆ DÉFINITION DPFC — Village life (vie au village)
  Village life refers to the way of living in rural areas,
  characterised by agricultural activities, closeness to
  nature, strong community bonds, and limited access
  to modern services.
  (La vie au village désigne le mode de vie en zone rurale,
  caractérisé par les activités agricoles, la proximité
  avec la nature, de forts liens communautaires et
  un accès limité aux services modernes.)

◆ DÉFINITION DPFC — Superlative (superlatif)
  The superlative is a grammatical form used to express
  the highest or lowest degree of a quality among three
  or more elements. It is always used with "the".
  Short adjectives : the + adj + -est.
  Long adjectives : the most + adj.
  (Le superlatif est une forme grammaticale utilisée pour
  exprimer le degré le plus élevé ou le plus bas d'une
  qualité parmi trois éléments ou plus. Il s'utilise
  toujours avec "the".)

◆ DÉFINITION DPFC — Crop (culture agricole)
  A crop is a plant that is grown by farmers in fields
  for food or other uses, such as yam, maize, cocoa or rice.
  (Une culture agricole est une plante cultivée par les
  agriculteurs dans des champs pour se nourrir ou d'autres
  usages, comme l'igname, le maïs, le cacao ou le riz.)

◆ DÉFINITION DPFC — Community (communauté)
  A community is a group of people who live in the same
  area and share common values, traditions and activities.
  (Une communauté est un groupe de personnes qui vivent
  dans la même zone et partagent des valeurs, des traditions
  et des activités communes.)

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI
───────────────────────────────────────

✎ TRACE ÉCRITE (à recopier dans le cahier)

  LEÇON 14 — VILLAGE LIFE (La vie au village)
  Compétence C5 · Écoute simple · Anglais LV1 · 3ème

  I. VOCABULAIRE DE LA VIE AU VILLAGE

  Lieux :
  field        = champ
  river        = rivière
  well         = puits
  mango tree   = manguier
  path         = sentier
  farm         = ferme / exploitation agricole

  Activités :
  grow crops   = cultiver des récoltes
  fetch water  = aller chercher de l'eau
  cook         = cuisiner
  walk to school = marcher jusqu'à l'école
  share stories  = partager des histoires
  work the land  = travailler la terre

  Cultures :
  yam    = igname
  maize  = maïs
  rice   = riz
  cocoa  = cacao
  millet = mil

  Avantages :
  quiet     = calme
  peaceful  = paisible
  fresh air = air frais
  cheap     = bon marché
  safe      = sûr(e)

  Inconvénients :
  no hospital        = pas d'hôpital
  no Internet        = pas d'Internet
  few jobs           = peu d'emplois
  limited electricity = électricité limitée
  far from services  = loin des services

  II. LE SUPERLATIF

  Adjectif COURT → the + adj + -EST
    quiet    → the quietest   (le/la plus calme)
    fresh    → the freshest   (le/la plus frais/fraîche)
    hard     → the hardest    (le/la plus difficile)
    ⚠ CVC : big → the biggest (double G)
    ⚠ hot → the hottest (double T)

  Adjectif LONG → the most + adj
    peaceful   → the most peaceful   (le/la plus paisible)
    expensive  → the most expensive  (le/la plus cher/chère)
    beautiful  → the most beautiful  (le/la plus beau/belle)

  ⚠ RÈGLE ABSOLUE : TOUJOURS "the" devant le superlatif.

  Exemples :
  "The village is the quietest place I know."
  (Le village est l'endroit le plus calme que je connaisse.)
  "Abidjan is the biggest city in Côte d'Ivoire."
  (Abidjan est la plus grande ville de Côte d'Ivoire.)

  III. COMPARATIF vs SUPERLATIF — RÉCAPITULATIF

  Comparatif (2 éléments) :
  "The village is quieter THAN the city."
  (Le village est plus calme QUE la ville.)

  Superlatif (3 éléments ou plus, ou degré absolu) :
  "The village is THE quietest place."
  (Le village est l'endroit LE PLUS calme.)

~ ANALOGIE CI — 1
  Le superlatif, c'est comme proclamer le champion
  d'une compétition de lutte traditionnelle à Bouaké :
  il n'y en a qu'UN qui est "the strongest" (le plus fort).
  Le comparatif, c'est juste dire qui est plus fort
  que l'autre dans un match en 1-contre-1.

~ ANALOGIE CI — 2
  "Fetch water" au village, c'est exactement comme aller
  au marigot ou au puits avec une bassine et rentrer
  avec cette bassine pleine. "Fetch" = aller ET revenir
  avec quelque chose — c'est un aller-retour complet.

────────────────────────────────────────────────────────

Phase 5b · Synthèse
────────────────────

  5 points essentiels à retenir :
  Point 1 · La vie au village est caractérisée par
            l'agriculture, la nature, la communauté
            et la tranquillité.
  Point 2 · Le superlatif exprime le degré maximum
            parmi trois éléments ou plus.
  Point 3 · Superlatif adjectif court : the + adj + -EST.
  Point 4 · Superlatif adjectif long : the most + adj.
  Point 5 · "The" est TOUJOURS présent avant le superlatif.

  5 mots clés + définition courte :
  · Field = champ (terrain cultivé par les agriculteurs)
  · Crop = récolte / culture (plante cultivée pour
    se nourrir)
  · Peaceful = paisible (calme et tranquille)
  · Community = communauté (groupe de personnes
    partageant le même lieu et les mêmes valeurs)
  · Fetch = aller chercher et ramener

  3 erreurs fréquentes + correction :
  Erreur 1 · Oublier "the" devant le superlatif :
             "Village is quietest place."
  Correction · The village is THE quietest place.

  Erreur 2 · Mélanger comparatif et superlatif :
             "The village is more quiet than the city
             and the quieter place I know."
  Correction · Comparatif pour 2 choses : quieter than.
               Superlatif pour le degré absolu :
               the quietest place.

  Erreur 3 · Écrire "the most quietest" (double superlatif)
  Correction · Soit "the quietest" soit "the most peaceful"
               — jamais les deux formes ensemble.

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés
───────────────────────────

◉ EXERCICE GUIDÉ 1 — Le superlatif
  Notion ciblée : formation du superlatif

  Énoncé :
  Transforme ces phrases comparatives en phrases
  superlatives selon le modèle.

  Modèle :
  Katiola is smaller than Abidjan.
  → Katiola is the smallest village in the region.

  1. The mango tree is taller than the other trees.
     → The mango tree is ________ tree in the village.
  2. Village air is fresher than city air.
     → Village air is ________ air in the country.
  3. The river is more beautiful than the well.
     → The river is ________ place in the village.
  4. Farming is harder than cooking.
     → Farming is ________ work in the village.
  5. Life at the village is more peaceful than city life.
     → Village life is ________ life I know.

  MÉTHODE
  Étape 1 · Identifie l'adjectif dans la phrase comparative.
  Étape 2 · Détermine si c'est un adjectif court ou long.
            tall → 1 syllabe → the + -EST → the tallest
            fresh → 1 syllabe → the + -EST → the freshest
            beautiful → 3 syllabes → the most → the most beautiful
            hard → 1 syllabe → the + -EST → the hardest
            peaceful → 2 syllabes → the most → the most peaceful
  Étape 3 · Ajoute "the" obligatoirement devant le superlatif.
  ✔ Vérification : "the" toujours présent ·
    jamais de double superlatif
  ✔ Résultats :
     1. The mango tree is the tallest tree in the village.
        (Le manguier est le plus grand arbre du village.)
     2. Village air is the freshest air in the country.
        (L'air du village est le plus frais du pays.)
     3. The river is the most beautiful place in the village.
        (La rivière est le plus bel endroit du village.)
     4. Farming is the hardest work in the village.
        (L'agriculture est le travail le plus difficile
        du village.)
     5. Village life is the most peaceful life I know.
        (La vie au village est la vie la plus paisible
        que je connaisse.)

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Décrire la vie au village
  Notion ciblée : Present Simple + vocabulaire rural
                  + superlatif en contexte

  Énoncé :
  Lis ces informations sur la famille Diabaté de Boundiali
  et rédige un paragraphe de 4 phrases en anglais.

  Informations :
  · Famille : père, mère, 3 enfants
  · Activité : cultivent l'igname et le maïs
  · Matin : les enfants marchent jusqu'à l'école
  · Soir : famille réunie sous le manguier
  · Opinion de Koné : "le village est l'endroit
    le plus paisible du monde"

  MÉTHODE
  Étape 1 · Phrase 1 — présente la famille et son activité.
            Sujet : "The Diabaté family" → Present Simple
            → "The Diabaté family lives in Boundiali.
               They grow yam and maize."
  Étape 2 · Phrase 2 — décris la routine matinale.
            Sujet : "The children" → Present Simple
            → "Every morning, the children walk to school."
  Étape 3 · Phrase 3 — décris la soirée.
            Sujet : "In the evening, the family" → Present Simple
            → "In the evening, the family sits together
               under the mango tree."
  Étape 4 · Phrase 4 — utilise le superlatif pour l'opinion.
            → "For Koné, the village is the most peaceful
               place in the world."
  ✔ Vérification : Present Simple correct ·
    superlatif avec "the" · vocabulaire rural présent
  ✔ Paragraphe complet :
    "The Diabaté family lives in Boundiali.
    They grow yam and maize in their fields.
    Every morning, the children walk to school.
    In the evening, the family sits together under
    the mango tree. For Koné, the village is the most
    peaceful place in the world."
    (La famille Diabaté vit à Boundiali. Ils cultivent
    l'igname et le maïs dans leurs champs. Chaque matin,
    les enfants marchent jusqu'à l'école. Le soir, la
    famille se réunit sous le manguier. Pour Koné, le
    village est l'endroit le plus paisible du monde.)

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

  Exercices d'entraînement

◎ EXERCICE 1 — Vocabulaire de la vie au village
  Notions travaillées : village life vocabulary

  Associe chaque mot anglais à sa traduction française.

  Colonne A (anglais) :
  1. field
  2. fetch water
  3. crop
  4. mango tree
  5. peaceful

  Colonne B (français) :
  a. paisible
  b. manguier
  c. champ
  d. aller chercher de l'eau
  e. récolte / culture agricole

  GUIDE
  Étape 1 · Relis la trace écrite — vocabulaire de la vie
            au village.
  Étape 2 · Pour chaque mot, cherche sa traduction.
  Étape 3 · Note tes réponses : 1→c, etc.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Present Simple : habitudes rurales
  Notions travaillées : Present Simple + vocabulaire
                        des activités rurales

  Mets les verbes entre parenthèses au Present Simple.
  Attention à l'accord avec he/she/it (verbe + -S).
  Traduis ensuite chaque phrase.

  1. Farmers (grow) ________ yam and maize every season.
  2. The woman (fetch) ________ water from the river.
  3. Children (walk) ________ to school every morning.
  4. The family (sit) ________ under the mango tree at night.
  5. People (share) ________ stories in the evening.

  GUIDE
  Étape 1 · Identifie le sujet de chaque phrase.
  Étape 2 · Vérifie : sujet singulier he/she/it → verbe + -S.
  Étape 3 · Écris la phrase complète et traduis-la.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Formation du superlatif
  Notions travaillées : règles du superlatif

  Donne le superlatif de ces adjectifs.
  Précise la règle appliquée.

  1. small    → ________ (règle : ?)
  2. beautiful → ________ (règle : ?)
  3. hot      → ________ (règle : ?)
  4. safe     → ________ (règle : ?)
  5. difficult → ________ (règle : ?)

  GUIDE
  Étape 1 · Compte les syllabes de chaque adjectif.
  Étape 2 · Vérifie le schéma CVC si nécessaire.
  Étape 3 · Applique la règle et n'oublie pas "the".

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Comparer ville et village
  Notions travaillées : comparatif + superlatif en contexte

  Écris 4 phrases en anglais comparant ou décrivant
  au superlatif la ville et le village.
  Utilise un adjectif différent à chaque fois.

  Adjectifs imposés : noisy · cheap · modern · exciting

  Exemple fourni :
  quiet → "The village is quieter than the city.
           It is the quietest place I know."

  Pour chaque adjectif :
  · 1 phrase au comparatif (ville vs village)
  · 1 phrase au superlatif

  GUIDE
  Étape 1 · Identifie si l'adjectif est court ou long.
  Étape 2 · Forme le comparatif : adj + -er than /
            more + adj + than.
  Étape 3 · Forme le superlatif : the + adj + -est /
            the most + adj.
  Étape 4 · Construis deux phrases pour chaque adjectif.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Rédiger un texte sur la vie au village
  Notions travaillées : cohérence textuelle + Present Simple
                        + superlatif + vocabulaire complet

  Aminata vit dans le village de Daoukro.
  Elle doit écrire un paragraphe en anglais
  pour son cours d'anglais afin de décrire
  sa vie au village.

  Écris ce paragraphe pour elle.
  Longueur : 70 à 90 mots.

  Tu dois inclure :
  · Au moins 3 phrases au Present Simple
    (habitudes quotidiennes)
  · Au moins 2 superlatifs
  · Au moins 4 mots de vocabulaire de la vie au village
    (field, crop, fetch, mango tree, peaceful,
    community, etc.)
  · Le nom de Daoukro ou d'un lieu ivoirien réel
  · Un avantage ET un inconvénient de la vie au village

  GUIDE
  Étape 1 · Commence par présenter Aminata et son village.
  Étape 2 · Décris ses habitudes quotidiennes
            (Present Simple).
  Étape 3 · Donne un avantage au superlatif :
            "Daoukro is the most peaceful…"
  Étape 4 · Donne un inconvénient
            (there is no… / it is hard to…).
  Étape 5 · Conclus par le sentiment d'Aminata
            sur sa vie au village.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — Village Life vs City Life
  Durée conseillée : 35 minutes

  CONTEXTE :
  Ton professeur d'anglais t'a demandé d'écrire
  un texte comparatif pour le journal du lycée.
  Le sujet : "Village life and city life in Côte d'Ivoire".

  ÉNONCÉ :
  Rédige un texte de 100 à 120 mots en anglais comparant
  la vie au village et la vie en ville en Côte d'Ivoire.

  Ton texte doit contenir :
  1. Une introduction : présente les deux milieux.
  2. Un paragraphe sur la vie au village :
     avantages + inconvénients.
  3. Un paragraphe sur la vie en ville :
     avantages + inconvénients.
  4. Une conclusion : quel milieu préfères-tu et pourquoi ?

  Contraintes obligatoires :
  · Minimum 3 phrases au Present Simple
  · Minimum 2 phrases au comparatif (avec "than")
  · Minimum 2 phrases au superlatif (avec "the")
  · Minimum 5 mots de vocabulaire variés
    (ville + village confondus)
  · Au moins 2 noms de lieux ivoiriens réels

  GUIDE PAS À PAS (sans corrigé)
  Étape 1 · Introduction (2 phrases) :
            "In Côte d'Ivoire, some people live in cities
            like… while others live in villages like…"
  Étape 2 · Paragraphe village (3-4 phrases) :
            Habitudes rurales au Present Simple +
            un superlatif positif + un inconvénient.
  Étape 3 · Paragraphe ville (3-4 phrases) :
            Services et avantages + comparatif +
            un inconvénient.
  Étape 4 · Conclusion (2 phrases) :
            Ton avis personnel avec justification.
            "In my opinion, I prefer… because it is
            the most… for me."
  Étape 5 · Compte tes mots (100 à 120).
  Étape 6 · Relis et vérifie :
            "the" devant chaque superlatif ?
            "than" dans chaque comparatif ?
            Present Simple correct (verbe + -S
            pour he/she/it) ?

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON

  TITRE PROPOSÉ :
  "Village Life and City Life in Côte d'Ivoire"

  TEXTE PROPOSÉ :

  In Côte d'Ivoire, some people live in big cities
  like Abidjan, while others live in small villages
  like Daoukro or Katiola.

  Village life is the most peaceful way of living.
  Farmers grow yam and maize in their fields every day.
  Women fetch water from the river and cook traditional
  meals. The air is fresher and the community is stronger
  than in the city. However, there are no hospitals nearby
  and there is no Internet connection.

  City life is more exciting than village life. There are
  many jobs, good schools and modern hospitals in Abidjan.
  But the city is the most stressful place to live.
  Traffic jams and pollution make life harder every day.

  In my opinion, I prefer village life because it is
  the most peaceful and the cheapest way to live.

  (Nombre de mots : 117 — dans les limites requises)

  TRADUCTION COMPLÈTE :
  En Côte d'Ivoire, certaines personnes vivent dans
  de grandes villes comme Abidjan, tandis que d'autres
  vivent dans de petits villages comme Daoukro ou Katiola.
  La vie au village est le mode de vie le plus paisible.
  Les agriculteurs cultivent l'igname et le maïs dans leurs
  champs chaque jour. Les femmes vont chercher de l'eau
  à la rivière et cuisinent des repas traditionnels. L'air
  est plus frais et la communauté est plus forte qu'en ville.
  Cependant, il n'y a pas d'hôpitaux à proximité et pas
  de connexion Internet.
  La vie en ville est plus excitante que la vie au village.
  Il y a beaucoup d'emplois, de bonnes écoles et des
  hôpitaux modernes à Abidjan. Mais la ville est l'endroit
  le plus stressant où vivre. Les embouteillages et la
  pollution rendent la vie plus difficile chaque jour.
  À mon avis, je préfère la vie au village parce que c'est
  le mode de vie le plus paisible et le moins cher.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1

  1→c  (field = champ)
  2→d  (fetch water = aller chercher de l'eau)
  3→e  (crop = récolte / culture agricole)
  4→b  (mango tree = manguier)
  5→a  (peaceful = paisible)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2

  1. Farmers grow yam and maize every season.
     (Les agriculteurs cultivent l'igname et le maïs
     chaque saison.)
     → "Farmers" = they → pas de -S sur le verbe

  2. The woman fetches water from the river.
     (La femme va chercher de l'eau à la rivière.)
     → "The woman" = she → fetch + -ES (verbe en -ch)

  3. Children walk to school every morning.
     (Les enfants marchent jusqu'à l'école chaque matin.)
     → "Children" = they → pas de -S

  4. The family sits together under the mango tree at night.
     (La famille s'assoit ensemble sous le manguier le soir.)
     → "The family" = it/she → sit + -S → sits

  5. People share stories in the evening.
     (Les gens partagent des histoires le soir.)
     → "People" = they → pas de -S

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3

  1. small → the smallest
     (Règle : adjectif court 1 syllabe → the + adj + -EST)

  2. beautiful → the most beautiful
     (Règle : adjectif long 3 syllabes → the most + adj)

  3. hot → the hottest
     (Règle : adjectif court CVC [h-o-t] → double le T
     → the hottest)

  4. safe → the safest
     (Règle : adjectif court en -E → on ajoute simplement
     -ST → the safest)

  5. difficult → the most difficult
     (Règle : adjectif long 3 syllabes → the most + adj)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4

  Adjectif 1 : noisy (bruyant — 2 syllabes → more/most)
  Comparatif : "The city is noisier than the village."
  ⚠ noisy → 2 syllabes mais finit en -Y → noisier
  (règle : adj en consonne + Y → change Y en I + ER)
  (Abidjan est plus bruyante que le village.)
  Superlatif : "Abidjan is the noisiest city I know."
  (Abidjan est la ville la plus bruyante que je connaisse.)

  Adjectif 2 : cheap (bon marché — 1 syllabe → -ER/-EST)
  Comparatif : "The village is cheaper than the city."
  (Le village est moins cher que la ville.)
  Superlatif : "The village is the cheapest place to live."
  (Le village est l'endroit le moins cher où vivre.)

  Adjectif 3 : modern (moderne — 2 syllabes → more/most)
  Comparatif : "The city is more modern than the village."
  (La ville est plus moderne que le village.)
  Superlatif : "Abidjan is the most modern city
               in Côte d'Ivoire."
  (Abidjan est la ville la plus moderne de Côte d'Ivoire.)

  Adjectif 4 : exciting (excitant — 3 syllabes → more/most)
  Comparatif : "City life is more exciting than village life."
  (La vie en ville est plus excitante que la vie au village.)
  Superlatif : "For young people, the city is the most
               exciting place to live."
  (Pour les jeunes, la ville est l'endroit le plus
  excitant où vivre.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5

  MODÈLE DE PARAGRAPHE PROPOSÉ :

  "My name is Aminata and I live in Daoukro, a small
  village in Côte d'Ivoire. Every day, my mother fetches
  water from the river and cooks traditional meals.
  My father grows yam and maize in the fields.
  In the evening, we sit together under the mango tree.
  Daoukro is the most peaceful place I know — the air
  is the freshest and the community is the strongest.
  However, there is no hospital nearby and there is
  no Internet. For me, village life is the best."

  TRADUCTION :
  Je m'appelle Aminata et je vis à Daoukro, un petit
  village de Côte d'Ivoire. Chaque jour, ma mère va
  chercher de l'eau à la rivière et cuisine des repas
  traditionnels. Mon père cultive l'igname et le maïs
  dans les champs. Le soir, nous nous asseyons ensemble
  sous le manguier. Daoukro est l'endroit le plus paisible
  que je connaisse — l'air est le plus frais et la
  communauté est la plus forte. Cependant, il n'y a pas
  d'hôpital à proximité et pas d'Internet. Pour moi,
  la vie au village est la meilleure.

  Vérification :
  ✔ Present Simple : fetches, cooks, grows, sit ✔
  ✔ Superlatifs : the most peaceful, the freshest,
    the strongest ✔
  ✔ Vocabulaire rural : river, mango tree, fields, yam,
    maize, community, crop ✔
  ✔ Lieu ivoirien : Daoukro ✔
  ✔ Avantage : most peaceful + freshest air ✔
  ✔ Inconvénient : no hospital, no Internet ✔

════════════════════════════════════════════════════════
  Prof. Evelyne ASSI · Anglais LV1 · 3ème
  DPFC/MENET-FP CI · 2025-2026
  "Speak even if it's not perfect — silence teaches nothing."
  (Parle même si c'est imparfait — le silence n'apprend rien.)
════════════════════════════════════════════════════════

"""
# ↑↑↑  FIN DU COURS  ↑↑↑


# ============================================================
#  GÉNÉRATEUR — NE PAS MODIFIER CE QUI SUIT
# ============================================================

import re
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---- Couleurs ----
C_BLEU_MARINE  = RGBColor(0x1A, 0x3A, 0x6B)
C_BLEU_VIF     = RGBColor(0x0D, 0x47, 0xA1)
C_VERT_FONCE   = RGBColor(0x1B, 0x5E, 0x20)
C_ROUGE        = RGBColor(0xC6, 0x28, 0x28)
C_ORANGE_FONCE = RGBColor(0xE6, 0x51, 0x00)
C_VIOLET       = RGBColor(0x6A, 0x1B, 0x9A)
C_BLANC        = RGBColor(0xFF, 0xFF, 0xFF)
C_GRIS_TEXTE   = RGBColor(0x42, 0x42, 0x42)

BG_VERT_PALE   = RGBColor(0xE8, 0xF5, 0xE9)
BG_BLEU_PALE   = RGBColor(0xE3, 0xF2, 0xFD)
BG_JAUNE_PALE  = RGBColor(0xFF, 0xF9, 0xC4)
BG_ORANGE_PALE = RGBColor(0xFF, 0xF3, 0xE0)
BG_ROUGE_PALE  = RGBColor(0xFF, 0xEB, 0xEE)
BG_VIOLET_PALE = RGBColor(0xF3, 0xE5, 0xF5)
BG_GRIS_PALE   = RGBColor(0xF5, 0xF5, 0xF5)
BG_GRIS_MOYEN  = RGBColor(0xEE, 0xEE, 0xEE)


def _hex(c): return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def set_cell_bg(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), _hex(color))
    tcPr.append(shd)


doc = Document()
for sec in doc.sections:
    sec.top_margin    = Cm(1.8)
    sec.bottom_margin = Cm(1.8)
    sec.left_margin   = Cm(2.0)
    sec.right_margin  = Cm(2.0)


def add_para(text="", bold=False, italic=False, size=11,
             color=None, align=WD_ALIGN_PARAGRAPH.LEFT,
             sb=0, sa=4):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if text:
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = color
    return p


def add_box(lines, bg=BG_GRIS_PALE, title=None,
            title_color=C_BLEU_VIF, title_size=11):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, bg)
    if title:
        p = cell.paragraphs[0]
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(title_size)
        run.font.color.rgb = title_color
        p.paragraph_format.space_after = Pt(4)
    first = True
    for line in lines:
        if first and not title:
            p = cell.paragraphs[0]
            first = False
        else:
            p = cell.add_paragraph()
        p.add_run(line).font.size = Pt(10)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_banner(title, bg=C_BLEU_MARINE):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = C_BLANC
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


# ============================================================
#  PARSEUR
# ============================================================

lines = COURS.strip().split('\n')
i = 0
total = len(lines)

# --- Extraction de l'en-tête (entre les deux lignes ════) ---
header_lines = []
in_header = False
header_done = False
body_start = 0

for idx, ln in enumerate(lines):
    s = ln.strip()
    if re.match(r'^[═]{4,}', s) and not in_header and not header_done:
        in_header = True
        continue
    if re.match(r'^[═]{4,}', s) and in_header:
        in_header = False
        header_done = True
        body_start = idx + 1
        break
    if in_header and s:
        header_lines.append(s)

# ---- Construire l'en-tête ----
if header_lines:
    # Ligne 1 : matière · leçon
    add_para(header_lines[0], bold=True, size=14,
             color=C_BLEU_MARINE, align=WD_ALIGN_PARAGRAPH.CENTER, sa=2)
    for hl in header_lines[1:]:
        add_para(hl, size=10, color=C_GRIS_TEXTE,
                 align=WD_ALIGN_PARAGRAPH.CENTER, sa=2)
    add_para(sa=4)

# ============================================================
#  CORPS DU DOCUMENT
# ============================================================

i = body_start
current_box_lines = []
current_box_type  = None   # "RETENIR" | "ATTENTION" | "TRACE" | "HARKNESS"
                           # "CAPSULE" | "ANCRAGE"   | "DEF"   | "ANALOGIE"
                           # "SCHEMA"  | "EXERCICE_G"| "CORRIGE"| "SOMMAIRE"
current_box_title = ""
in_tilde_block = False
tilde_lines = []


def flush_box():
    global current_box_lines, current_box_type, current_box_title
    if not current_box_lines and not current_box_type:
        return
    t  = current_box_type
    ls = current_box_lines
    ti = current_box_title

    if t == "SOMMAIRE":
        add_box(ls, bg=BG_GRIS_PALE, title="Sommaire")
    elif t == "CAPSULE":
        add_box(ls, bg=BG_VERT_PALE, title=ti, title_color=C_VERT_FONCE)
    elif t == "ANCRAGE":
        add_box(ls, bg=BG_BLEU_PALE, title=ti, title_color=C_BLEU_VIF)
    elif t == "SCHEMA":
        add_box(ls, bg=BG_GRIS_MOYEN, title=ti, title_color=C_BLEU_VIF)
    elif t == "HARKNESS":
        add_box(ls, bg=BG_GRIS_PALE, title=ti, title_color=C_BLEU_MARINE)
    elif t == "RETENIR":
        add_box(ls, bg=BG_JAUNE_PALE, title=ti, title_color=C_BLEU_VIF)
    elif t == "ATTENTION":
        add_box(ls, bg=BG_ROUGE_PALE, title=ti, title_color=C_ROUGE)
    elif t == "TRACE":
        add_box(ls, bg=BG_VERT_PALE, title=ti, title_color=C_VERT_FONCE)
    elif t == "ANALOGIE":
        add_box(ls, bg=BG_ORANGE_PALE, title=ti, title_color=C_ORANGE_FONCE)
    elif t == "DEF":
        add_box(ls, bg=BG_VIOLET_PALE, title=ti, title_color=C_VIOLET)
    elif t == "EXERCICE_G":
        add_box(ls, bg=BG_BLEU_PALE, title=ti, title_color=C_BLEU_VIF)
    elif t == "CORRIGE":
        add_box(ls, bg=BG_VERT_PALE, title=ti, title_color=C_VERT_FONCE)
    elif t == "TEXTE":
        add_box(ls, bg=BG_GRIS_PALE)
    else:
        for ln in ls:
            if ln.strip():
                add_para(ln.strip(), size=10, sa=2)

    current_box_lines = []
    current_box_type  = None
    current_box_title = ""


while i < total:
    raw = lines[i]
    s   = raw.strip()
    i  += 1

    # Lignes vides
    if not s:
        continue

    # ── Séparateurs ────────────────────────────────────────
    if re.match(r'^[═]{4,}', s):
        flush_box()
        continue

    if re.match(r'^[─]{4,}', s):
        flush_box()
        continue

    # ── SECTION banner ─────────────────────────────────────
    if re.match(r'^SECTION \d+\s*[—–-]', s):
        flush_box()
        add_banner(s)
        continue

    # ── Phase ──────────────────────────────────────────────
    if re.match(r'^Phase \d+\s*[·.]', s):
        flush_box()
        add_para(s, bold=True, size=12, color=C_BLEU_MARINE, sb=6, sa=3)
        continue

    # ── SOMMAIRE ───────────────────────────────────────────
    if s == "SOMMAIRE":
        flush_box()
        current_box_type  = "SOMMAIRE"
        current_box_title = "Sommaire"
        continue

    # ── [BLOC TITRE] ───────────────────────────────────────
    m = re.match(r'^\[(.+)\]$', s)
    if m:
        flush_box()
        titre_bloc = m.group(1)
        if "CAPSULE" in titre_bloc.upper():
            current_box_type  = "CAPSULE"
            current_box_title = "📌 " + titre_bloc
        elif "ANCRAGE" in titre_bloc.upper():
            current_box_type  = "ANCRAGE"
            current_box_title = "🌍 " + titre_bloc
        elif "SCHÉMA" in titre_bloc.upper() or "SCHEMA" in titre_bloc.upper():
            current_box_type  = "SCHEMA"
            current_box_title = "📊 " + titre_bloc
        else:
            current_box_type  = "TEXTE"
            current_box_title = titre_bloc
        continue

    # ── HARKNESS ───────────────────────────────────────────
    if re.match(r'^[▶►]\s*HARKNESS', s):
        flush_box()
        current_box_type  = "HARKNESS"
        current_box_title = "▶ " + s.lstrip("▶► ").strip()
        continue

    # ── À RETENIR ──────────────────────────────────────────
    if re.match(r'^[┌].*À RETENIR', s):
        flush_box()
        title_clean = re.sub(r'^[┌─\s]+', '', s).rstrip('─┐').strip()
        current_box_type  = "RETENIR"
        current_box_title = "📘 " + title_clean
        continue

    if re.match(r'^[└]', s):
        # Ferme l'encadré À RETENIR
        flush_box()
        continue

    # ── ATTENTION ──────────────────────────────────────────
    if re.match(r'^⚠', s):
        flush_box()
        current_box_type  = "ATTENTION"
        current_box_title = s
        continue

    # ── TRACE ÉCRITE ───────────────────────────────────────
    if re.match(r'^✎', s):
        flush_box()
        current_box_type  = "TRACE"
        current_box_title = "✎ " + s.lstrip("✎ ").strip()
        continue

    # ── ANALOGIE ───────────────────────────────────────────
    if re.match(r'^~\s*ANALOGIE', s):
        flush_box()
        current_box_type  = "ANALOGIE"
        current_box_title = "💡 " + s.lstrip("~ ").strip()
        continue

    # ── DÉFINITION DPFC ────────────────────────────────────
    if re.match(r'^◆\s*DÉFINITION', s):
        flush_box()
        current_box_type  = "DEF"
        current_box_title = s.lstrip("◆ ").strip()
        continue

    # ── EXERCICE GUIDÉ ─────────────────────────────────────
    if re.match(r'^◉', s):
        flush_box()
        current_box_type  = "EXERCICE_G"
        current_box_title = "◉ " + s.lstrip("◉ ").strip()
        continue

    # ── EXERCICE MAISON ────────────────────────────────────
    if re.match(r'^◎', s):
        flush_box()
        add_para("◎ " + s.lstrip("◎ ").strip(),
                 bold=True, size=11, color=C_BLEU_VIF, sb=4, sa=2)
        current_box_type  = "TEXTE"
        current_box_title = ""
        continue

    # ── DEVOIR MAISON ──────────────────────────────────────
    if re.match(r'^◈', s):
        flush_box()
        add_para("◈ " + s.lstrip("◈ ").strip(),
                 bold=True, size=12, color=C_ROUGE, sb=6, sa=3)
        continue

    # ── CORRIGÉ ────────────────────────────────────────────
    if re.match(r'^✔\s*CORRIGÉ', s):
        flush_box()
        current_box_type  = "CORRIGE"
        current_box_title = "✔ " + s.lstrip("✔ ").strip()
        continue

    # ── Bloc ~~~ (texte anglais) ───────────────────────────
    if s == "~~~":
        if not in_tilde_block:
            in_tilde_block = True
            tilde_lines = []
        else:
            in_tilde_block = False
            flush_box()
            add_box(tilde_lines, bg=BG_GRIS_MOYEN,
                    title="📖 Texte", title_color=C_BLEU_MARINE)
            tilde_lines = []
        continue

    if in_tilde_block:
        tilde_lines.append(s)
        continue

    # ── Texte courant → ajouter au bloc en cours ──────────
    if current_box_type is not None:
        current_box_lines.append(s)
    else:
        # Pas de bloc ouvert → paragraphe normal
        if re.match(r'^(Objectifs|Prérequis|MÉTHODE|GUIDE|TRADUCTION)', s):
            add_para(s, bold=True, size=11, color=C_BLEU_MARINE, sb=3, sa=2)
        elif re.match(r'^(→|·|\d+\.)\s', s):
            p = doc.add_paragraph(style="List Bullet")
            p.add_run(s.lstrip("→·").strip()).font.size = Pt(10)
        else:
            add_para(s, size=10, sa=2)

flush_box()

# ============================================================
#  SAUVEGARDE ET TÉLÉCHARGEMENT
# ============================================================

# Nom automatique depuis l'en-tête
nom = "cours_lpa.docx"
if header_lines:
    nom_raw = header_lines[0]
    nom = re.sub(r'[^\w\s·—-]', '', nom_raw)
    nom = re.sub(r'[\s·—]+', '_', nom).strip('_') + ".docx"

doc.save(nom)
print(f"✅ Document généré : {nom}")

try:
    from google.colab import files
    files.download(nom)
except ImportError:
    print("(Hors Colab — fichier sauvegardé localement)")
