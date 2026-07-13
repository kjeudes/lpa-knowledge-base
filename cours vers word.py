# ============================================================
#  CONVERTISSEUR DE COURS → WORD  (LPA / DPFC Côte d'Ivoire)
#  Google Colab — python-docx  — VERSION 4 (style modèle v2)
#
#  UTILISATION :
#  1. Installe : !pip install python-docx -q
#  2. Colle ton cours dans la variable COURS ci-dessous
#  3. Exécute → téléchargement automatique
#
#  FORMAT DU COURS (inchangé) :
#  ════...════  en-tête  ════...════
#  TITRE_LECON: ...
#  COMPETENCE: ...
#  PROFESSEUR: ...
#  ORGANISME: ...
#  SOMMAIRE
#  ────...────
#  SECTION N — ...
#  [CAPSULE DE REPRISE — ...]
#  [ANCRAGE IVOIRIEN]
#  ▶ HARKNESS N — ...
#  Phase N · ...
#  ~~~  texte  ---TRADUCTION---  ~~~
#  ┌─ À RETENIR — ... ─┐ ... └─┘
#  ⚠ ATTENTION — ...
#  ◆ DÉFINITION DPFC — ...
#  ✎ TRACE ÉCRITE
#  ~ ANALOGIE CI — N
#  SYNTHESE_DEBUT ... SYNTHESE_FIN
#  ◉ EXERCICE GUIDÉ N — ...
#  ◎ EXERCICE N — ...
#  ◈ DEVOIR MAISON — ...
#  ✔ CORRIGÉ — ...
#  [SCHÉMA N — ...]
# ============================================================

COURS = """════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 12 — Fashion show (Défilé de mode)
  Niveau : 3ème | Compétence : C4 — Écrit élaboré
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

[CAPSULE DE REPRISE — Leçon 11 : Modern and traditional clothes]

  5 points clés de la leçon précédente :
  Point 1 · On distingue deux grands types de vêtements :
            modern clothes (vêtements modernes) et
            traditional clothes (vêtements traditionnels).
  Point 2 · Le vocabulaire clé : dress (robe), shirt (chemise),
            trousers (pantalon), kente (tissu traditionnel ghanéen
            et ivoirien), boubou (grand vêtement traditionnel
            d'Afrique de l'Ouest).
  Point 3 · On utilise le Present Simple pour décrire des habitudes
            vestimentaires : "She wears a boubou every Friday."
            (Elle porte un boubou chaque vendredi.)
  Point 4 · L'adjectif se place AVANT le nom en anglais :
            "a beautiful dress" — jamais "a dress beautiful".
  Point 5 · Verbes utiles : wear (porter), dress (s'habiller),
            choose (choisir), design (créer/concevoir).

  3 questions flash :
  Q1 · Comment dit-on "tissu traditionnel" en anglais ?
       → traditional fabric
  Q2 · Traduis : "He is wearing a white shirt."
       → Il porte une chemise blanche.
  Q3 · Place l'adjectif correctement : "shoes / leather / brown"
       → brown leather shoes (chaussures en cuir marron)

  Lien avec la leçon du jour :
  Aujourd'hui, on va plus loin : on ne parle plus seulement
  des vêtements, on décrit un ÉVÉNEMENT — le défilé de mode
  (fashion show). Tu vas apprendre à rédiger une description
  d'événement en anglais, à parler de l'apparence des gens
  et à utiliser le Present Continuous pour décrire ce qui
  se passe en ce moment.

────────────────────────────────────────────────────────

[ANCRAGE IVOIRIEN]

  Imagine que tu es à Abidjan, au Palais de la Culture
  de Treichville. Ce soir a lieu le Grand Défilé Wax & Fashion,
  organisé par le styliste ivoirien Adama TRAORÉ.
  Des mannequins défilent sur le podium (runway) vêtus de
  tenues en wax coloré, en bogolan et en bazin riche.
  Des centaines de spectateurs applaudissent. Des journalistes
  de RTI (Radiodiffusion Télévision Ivoirienne) prennent des notes.
  Toi, tu dois écrire un article en anglais pour le journal
  de ton école sur cet événement.

  Lien DPFC : Entrepreneuriat et valorisation du patrimoine
  culturel ivoirien à travers la mode.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — La mode et l'identité culturelle
  Q : "Can fashion show who you are and where you come from?"
      (La mode peut-elle montrer qui tu es et d'où tu viens ?)
  Guide :
    Piste 1 · Pense aux vêtements traditionnels de ta région
              (wax, kente, boubou…).
    Piste 2 · Pense aux vêtements que tu portes à l'école
              vs ceux du week-end.
    Piste 3 · Est-ce qu'une tenue peut raconter une histoire ?
  Corrigé :
    Yes, fashion can show who you are and where you come from.
    (Oui, la mode peut montrer qui tu es et d'où tu viens.)
    Traditional clothes like the wax fabric or the boubou show
    African identity and pride. (Les vêtements traditionnels
    comme le wax ou le boubou montrent l'identité et la fierté
    africaines.) At the same time, modern clothes like jeans
    and T-shirts show that young people follow global trends.
    (En même temps, les vêtements modernes comme les jeans
    montrent que les jeunes suivent les tendances mondiales.)
    In Côte d'Ivoire, many people mix both styles — for example,
    wearing a wax top with jeans. (En Côte d'Ivoire, beaucoup
    de personnes mélangent les deux styles.)

▶ HARKNESS 2 — Décrire ce que l'on voit
  Q : "How do you describe what is happening right now
       in English?" (Comment décrit-on ce qui se passe
       en ce moment en anglais ?)
  Guide :
    Piste 1 · Quelle est la différence entre "she walks"
              et "she is walking" ?
    Piste 2 · Quand utilise-t-on le Present Continuous ?
    Piste 3 · Comment forme-t-on le Present Continuous ?
  Corrigé :
    We use the Present Continuous to describe actions happening
    right now. (On utilise le Present Continu pour décrire
    des actions qui se passent en ce moment.)
    Form : Subject + am/is/are + verb-ing.
    (Forme : Sujet + am/is/are + verbe-ing.)
    Example : "The model is walking on the runway."
    (Le mannequin marche sur le podium — en ce moment.)
    We use the Present Simple for habits :
    "She walks every day." (Elle marche tous les jours.)

▶ HARKNESS 3 — Vocabulaire de la mode
  Q : "What words do you need to describe a fashion show?"
      (Quels mots faut-il pour décrire un défilé de mode ?)
  Guide :
    Piste 1 · Qui participe à un défilé de mode ?
    Piste 2 · Qu'est-ce qu'on voit, qu'est-ce qu'on entend ?
    Piste 3 · Comment décrire une tenue de façon précise ?
  Corrigé :
    People at a fashion show : model/mannequin (mannequin),
    designer (styliste), audience (public/spectateurs),
    photographer (photographe), journalist (journaliste).
    (Personnes présentes : mannequin, styliste, public,
    photographe, journaliste.)
    What you see : runway (podium/passerelle), outfit (tenue),
    collection (collection), spotlight (projecteur).
    (Ce qu'on voit : podium, tenue, collection, projecteur.)
    Describing an outfit : elegant (élégant), colourful
    (coloré), fitted (ajusté), loose (ample),
    embroidered (brodé), patterned (à motifs).

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis
────────────────────────────────────

  Titre : Fashion show (Le défilé de mode)
  Compétence : C4 — Écrit élaboré
  Professeur Evelyne ASSI : "Aujourd'hui on ne se contente
  pas de nommer les vêtements — on raconte ce qui se passe.
  Parle même si c'est imparfait — le silence n'apprend rien."

  Objectifs de la leçon :
  → Décrire un événement de mode en anglais (à l'écrit)
  → Utiliser le Present Continuous correctement
  → Maîtriser le vocabulaire du défilé de mode
  → Rédiger un paragraphe descriptif structuré

  Prérequis nécessaires :
  → Vocabulaire des vêtements (Leçon 11)
  → Les adjectifs de couleur, de taille, de matière
  → La structure Subject + Verb + Complement
  → Le verbe "to be" au présent (am / is / are)

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée
────────────────────────────

  LIS CE TEXTE — il se passe en ce moment à Abidjan :

  ~~~
  It is Friday evening in Abidjan. The Grand Wax Fashion Show
  is taking place at the Palais de la Culture de Treichville.
  The audience is clapping and cheering. On the runway, a
  model named Mariama is walking confidently. She is wearing
  a long, colourful wax dress with golden embroidery. Another
  model, Kouamé, is showing a modern boubou in deep blue
  bogolan fabric. The designer, Adama Traoré, is smiling
  backstage. Photographers are taking pictures. Everyone
  is enjoying the show.
  ~~~

  TRADUCTION COMPLÈTE :
  C'est vendredi soir à Abidjan. Le Grand Défilé Wax Fashion
  se déroule au Palais de la Culture de Treichville.
  Le public applaudit et acclame. Sur le podium, un mannequin
  nommé Mariama marche avec assurance. Elle porte une longue
  robe en wax colorée avec une broderie dorée. Un autre
  mannequin, Kouamé, présente un boubou moderne en tissu
  bogolan bleu profond. Le styliste, Adama Traoré, sourit
  dans les coulisses. Des photographes prennent des photos.
  Tout le monde apprécie le spectacle.

┌─ À RETENIR ──────────────────────────────────────────┐
  Le PRESENT CONTINUOUS (présent continu) décrit une action
  qui se passe EN CE MOMENT, au moment où l'on parle.
  FORME : Subject + am / is / are + verb + -ING
  Exemples :
  · I am watching the show. (Je regarde le spectacle.)
  · She is wearing a wax dress. (Elle porte une robe en wax.)
  · They are taking pictures. (Ils prennent des photos.)
  CONTRAIRE : le Present Simple décrit une HABITUDE.
  · She wears wax every Friday. (Elle porte du wax chaque vendredi.)
└──────────────────────────────────────────────────────┘

  FORMATION DU PRESENT CONTINUOUS — Règle complète

  Verbe se terminant par -E muet → on retire le -E puis on ajoute -ING
    make → making (faire)
    write → writing (écrire)
    smile → smiling (sourire)

  Verbe court avec une voyelle + une consonne → on double la consonne
    run → running (courir)
    sit → sitting (s'asseoir)
    clap → clapping (applaudir)

  Tous les autres verbes → on ajoute simplement -ING
    walk → walking (marcher)
    wear → wearing (porter)
    show → showing (montrer)

⚠ ATTENTION — Faux amis et pièges
  · "He is wearing" = Il porte (en ce moment) — PAS "il est portant"
    Ne traduis JAMAIS mot à mot avec "être + participe présent".
  · "Fashion" se prononce : /ˈfæʃ.ən/ — FA-cheune
    PAS "fachione" ni "fashione".
  · "Designer" = styliste/créateur de mode — PAS "designeur"
    Prononciation : /dɪˈzaɪ.nər/ — di-ZAI-neur
  · "Runway" = podium/passerelle — en FR "podium" est aussi correct.
  · "Outfit" = une tenue complète — PAS seulement un vêtement.

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels
───────────────────────────

[SCHÉMA 1 — Le Present Continuous : formation]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  • Zone centrale (rectangle horizontal, fond vert clair) :
    PRESENT CONTINUOUS = AM / IS / ARE + VERBE + -ING
  • Colonne gauche (3 cercles empilés, couleur bleue) :
    Cercle 1 : I → am
    Cercle 2 : He / She / It → is
    Cercle 3 : You / We / They → are
  • Colonne droite (3 exemples en texte noir) :
    I am walking. (Je marche.)
    She is wearing. (Elle porte.)
    They are clapping. (Ils applaudissent.)
  • Flèche verte du bas vers le haut : "Action EN CE MOMENT"
  • Encadré rouge séparé à droite :
    ⚠ PIÈGE : "I am wear" → FAUX
    ✔ CORRECT : "I am wearing"
  • Légende : vert = correct · rouge = erreur à éviter
Note générateur : Canva (tableau comparatif) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Carte mentale : Vocabulary of a Fashion Show]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  • Centre (ovale jaune) : FASHION SHOW / DÉFILÉ DE MODE
  • Branche 1 (vers le haut, couleur orange) : PEOPLE
    → model (mannequin)
    → designer (styliste)
    → audience (public)
    → photographer (photographe)
  • Branche 2 (vers la droite, couleur bleue) : PLACE
    → runway (podium)
    → backstage (coulisses)
    → spotlight (projecteur)
    → stage (scène)
  • Branche 3 (vers le bas, couleur rose) : CLOTHES
    → outfit (tenue)
    → collection (collection)
    → embroidery (broderie)
    → fabric (tissu)
  • Branche 4 (vers la gauche, couleur verte) : ADJECTIVES
    → elegant (élégant)
    → colourful (coloré)
    → fitted (ajusté)
    → patterned (à motifs)
  • Traductions en français sous chaque mot anglais
Note générateur : Canva (mind map) ou MindMeister
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC
───────────────────────────

◆ DÉFINITION DPFC — Fashion show
  A fashion show is an event where designers present their
  new collection of clothes through models walking on a runway.
  (Un défilé de mode est un événement où les stylistes
  présentent leur nouvelle collection de vêtements à travers
  des mannequins défilant sur un podium.)

◆ DÉFINITION DPFC — Present Continuous
  The Present Continuous is a verb tense used to describe
  an action happening at the moment of speaking.
  It is formed with : Subject + am/is/are + verb + -ing.
  (Le Present Continu est un temps verbal utilisé pour
  décrire une action qui se passe au moment où l'on parle.
  Il se forme avec : Sujet + am/is/are + verbe + -ing.)

◆ DÉFINITION DPFC — Model (mannequin)
  A model is a person who wears clothes during a fashion show
  to present a designer's collection to the public.
  (Un mannequin est une personne qui porte des vêtements
  lors d'un défilé de mode pour présenter la collection
  d'un styliste au public.)

◆ DÉFINITION DPFC — Designer (styliste)
  A designer is a person who creates and makes clothes,
  bags, shoes or accessories.
  (Un styliste est une personne qui crée et fabrique
  des vêtements, sacs, chaussures ou accessoires.)

◆ DÉFINITION DPFC — Outfit (tenue)
  An outfit is a set of clothes worn together on a
  specific occasion.
  (Une tenue est un ensemble de vêtements portés ensemble
  pour une occasion précise.)

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI
───────────────────────────────────────

✎ TRACE ÉCRITE (à recopier dans le cahier)

  LEÇON 12 — FASHION SHOW (Le défilé de mode)
  Compétence C4 · Écrit élaboré · Anglais LV1 · 3ème

  I. VOCABULAIRE DU DÉFILÉ DE MODE

  Personnes :
  model = mannequin
  designer = styliste
  audience = public / spectateurs
  photographer = photographe

  Lieux et objets :
  runway = podium / passerelle
  backstage = coulisses
  outfit = tenue
  collection = collection
  fabric = tissu

  Adjectifs de description :
  elegant = élégant(e)
  colourful = coloré(e)
  fitted = ajusté(e)
  loose = ample
  embroidered = brodé(e)
  patterned = à motifs

  II. LE PRESENT CONTINUOUS

  Utilisation : action qui se passe EN CE MOMENT.
  Formation : Subject + am / is / are + verb + -ING

  I am walking.      → Je marche. (en ce moment)
  She is wearing.    → Elle porte. (en ce moment)
  They are clapping. → Ils applaudissent. (en ce moment)

  Règles de formation du -ING :
  · Verbe en -E muet → retire le E, ajoute -ING
    (make → making / smile → smiling)
  · Verbe court CVC → double la consonne finale, ajoute -ING
    (run → running / clap → clapping)
  · Autres verbes → ajoute simplement -ING
    (walk → walking / show → showing)

  III. STRUCTURE POUR DÉCRIRE UNE TENUE

  [Name] is wearing + a/an + [adj. couleur] + [adj. style]
  + [nom du vêtement] + with + [détail].

  Exemple :
  "Mariama is wearing a long, colourful wax dress with
  golden embroidery."
  (Mariama porte une longue robe en wax colorée avec
  une broderie dorée.)

~ ANALOGIE CI — 1
  Le Present Continuous, c'est comme un commentateur
  sportif sur Radio Côte d'Ivoire pendant un match des
  Éléphants : il décrit CE QUI SE PASSE EN CE MOMENT,
  pas ce qui se passe d'habitude.
  "Gradel is running!" = Gradel court ! (maintenant, sur le terrain)
  PAS "Gradel runs" qui voudrait dire qu'il court tous les jours.

~ ANALOGIE CI — 2
  Décrire une tenue à un défilé, c'est comme décrire
  l'habit d'un chef traditionnel à une cérémonie de
  chefferie à Yamoussoukro : tu donnes la couleur,
  le tissu, les ornements — dans cet ordre.

────────────────────────────────────────────────────────

Phase 5b · Synthèse
────────────────────

  5 points essentiels à retenir :
  Point 1 · Un fashion show (défilé de mode) est un événement
            où des mannequins présentent des collections de vêtements.
  Point 2 · On utilise le Present Continuous pour décrire
            une action en cours : Subject + am/is/are + verb-ing.
  Point 3 · Pour décrire une tenue, on suit l'ordre :
            couleur + style + nom du vêtement + détail.
  Point 4 · Les adjectifs se placent TOUJOURS avant le nom
            en anglais.
  Point 5 · Un texte descriptif d'événement en anglais
            suit l'ordre : Qui → Quoi → Comment.

  5 mots clés + définition courte :
  · Fashion show = défilé de mode (événement de présentation
    de vêtements)
  · Model = mannequin (personne qui présente les tenues)
  · Designer = styliste (créateur de vêtements)
  · Outfit = tenue (ensemble de vêtements)
  · Runway = podium (passerelle où défilent les mannequins)

  3 erreurs fréquentes + correction :
  Erreur 1 · Écrire "She is wear a dress."
  Correction · She is WEARing a dress. (toujours le -ING !)

  Erreur 2 · Mettre l'adjectif après le nom :
             "a dress colourful"
  Correction · a COLOURFUL dress (adjectif avant le nom)

  Erreur 3 · Oublier le verbe "to be" dans le Present Continuous :
             "She wearing a dress."
  Correction · She IS wearing a dress. (am/is/are obligatoire)

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés
───────────────────────────

◉ EXERCICE GUIDÉ 1 — Conjugaison au Present Continuous
  Notion ciblée : Formation du Present Continuous

  Énoncé :
  Mets les verbes entre parenthèses au Present Continuous.
  1. The model ________ (walk) on the runway.
  2. The designer ________ (smile) backstage.
  3. The photographers ________ (take) pictures.
  4. Aïcha ________ (wear) a beautiful wax dress.
  5. We ________ (watch) the fashion show.

  MÉTHODE
  Étape 1 · Identifie le sujet de chaque phrase.
            → Sujet 1 : "The model" = he/she → is
            → Sujet 2 : "The designer" = he/she → is
            → Sujet 3 : "The photographers" = they → are
            → Sujet 4 : "Aïcha" = she → is
            → Sujet 5 : "We" = we → are
  Étape 2 · Applique la règle du -ING à chaque verbe.
            → walk → walking (verbe ordinaire, on ajoute -ING)
            → smile → smiling (verbe en -E muet, on retire E)
            → take → taking (verbe en -E muet, on retire E)
            → wear → wearing (verbe ordinaire, on ajoute -ING)
            → watch → watching (verbe ordinaire, on ajoute -ING)
  Étape 3 · Assemble : Subject + is/are + verb-ING
  ✔ Vérification : chaque phrase doit contenir am/is/are + verbe-ING
  ✔ Résultats :
     1. The model is walking on the runway.
        (Le mannequin marche sur le podium.)
     2. The designer is smiling backstage.
        (Le styliste sourit dans les coulisses.)
     3. The photographers are taking pictures.
        (Les photographes prennent des photos.)
     4. Aïcha is wearing a beautiful wax dress.
        (Aïcha porte une belle robe en wax.)
     5. We are watching the fashion show.
        (Nous regardons le défilé de mode.)

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Rédiger une description de tenue
  Notion ciblée : Structure de description d'une tenue
                  + vocabulaire du défilé

  Énoncé :
  Regarde ces informations sur le mannequin Kouamé et écris
  une description complète de 3 phrases en anglais.

  Informations :
  · Prénom : Kouamé
  · Il marche sur le podium (en ce moment)
  · Il porte : boubou / bleu foncé / tissu bogolan / broderies argentées
  · Le public applaudit (en ce moment)

  MÉTHODE
  Étape 1 · Rédige la première phrase : qui + que fait-il en ce moment
            → Sujet : Kouamé → is (Present Continuous)
            → Action : walk → walking
            → Lieu : on the runway
            → Phrase 1 : "Kouamé is walking on the runway."
  Étape 2 · Décris sa tenue :
            ordre → couleur + type de tissu + nom du vêtement + détail
            → "He is wearing a dark blue bogolan boubou
               with silver embroidery."
            (Il porte un boubou bogolan bleu foncé avec
            des broderies argentées.)
  Étape 3 · Décris la réaction du public :
            → Sujet : "The audience" = they → are
            → Action : clap → clapping (verbe court CVC : double P)
            → "The audience is clapping."
            (Le public applaudit.)
  ✔ Vérification : 3 phrases au Present Continuous · vocabulaire
    du défilé présent · adjectifs avant les noms
  ✔ Réponse complète :
    "Kouamé is walking on the runway. He is wearing a dark blue
    bogolan boubou with silver embroidery. The audience is clapping."

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

  Exercices d'entraînement

◎ EXERCICE 1 — Vocabulaire du défilé de mode
  Notions travaillées : fashion show vocabulary

  Relie chaque mot anglais à sa traduction française.
  (Écris le numéro correspondant.)

  Colonne A (anglais) :
  1. runway
  2. designer
  3. audience
  4. outfit
  5. backstage

  Colonne B (français) :
  a. styliste
  b. coulisses
  c. tenue
  d. podium / passerelle
  e. public / spectateurs

  GUIDE
  Étape 1 · Relis la trace écrite — vocabulaire du défilé de mode.
  Étape 2 · Pour chaque mot anglais, cherche sa traduction
            dans la Colonne B.
  Étape 3 · Note tes réponses sous la forme : 1→d, 2→a, etc.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Formation du Present Continuous
  Notions travaillées : règles du -ING / conjugaison

  Transforme ces verbes à l'infinitif en forme -ING.
  Applique la bonne règle à chaque fois.

  1. dance   →  _______  (règle : ?)
  2. sit     →  _______  (règle : ?)
  3. show    →  _______  (règle : ?)
  4. write   →  _______  (règle : ?)
  5. run     →  _______  (règle : ?)

  GUIDE
  Étape 1 · Pour chaque verbe, identifie sa terminaison :
            se termine par -E muet ? par CVC ? ordinaire ?
  Étape 2 · Applique la règle correspondante.
  Étape 3 · Écris la forme -ING correcte.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Construire des phrases au Present Continuous
  Notions travaillées : structure Subject + am/is/are + verb-ING

  Construis des phrases complètes au Present Continuous
  avec les éléments donnés.

  1. Fatou / wear / a long red dress
  2. The photographer / take / pictures
  3. I / look / at the runway
  4. The models / walk / on the stage
  5. The designer / show / his new collection

  GUIDE
  Étape 1 · Identifie le sujet et choisis am / is / are.
  Étape 2 · Mets le verbe à la forme -ING.
  Étape 3 · Ajoute le complément.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Décrire une tenue
  Notions travaillées : ordre des adjectifs + vocabulaire tenue

  Lis ces informations et écris une phrase de description
  de tenue pour chaque mannequin.
  Utilise le Present Continuous.

  Mannequin A — Awa :
  tenue : robe / wax / verte / longue / avec ceinture dorée

  Mannequin B — Issiaka :
  tenue : boubou / blanc / en bazin / avec broderies bleues

  Mannequin C — Nadia :
  tenue : jupe / courte / noire / avec motifs géométriques

  GUIDE
  Étape 1 · Structure : [Name] is wearing a/an [adj.] [adj.]
            [vêtement] with [détail].
  Étape 2 · Place les adjectifs AVANT le nom.
  Étape 3 · Ajoute le détail (with + accessoire ou ornement).

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Rédiger un paragraphe descriptif
  Notions travaillées : cohérence textuelle + Present Continuous
                        + vocabulaire complet du défilé

  Tu assistes au Grand Défilé Wax Fashion au Palais de la
  Culture d'Abidjan. Rédige un paragraphe de 5 phrases
  en anglais pour décrire ce que tu vois en ce moment.

  Tu dois inclure :
  · Au moins 3 verbes au Present Continuous
  · Au moins 2 descriptions de tenues
  · Au moins 3 mots de vocabulaire du défilé
    (runway, designer, audience, outfit, backstage, etc.)
  · Le nom d'un lieu ivoirien et un prénom ivoirien

  GUIDE
  Étape 1 · Commence par situer l'action (où ? quand ?).
  Étape 2 · Décris ce qui se passe sur le podium.
  Étape 3 · Décris les tenues de 2 mannequins.
  Étape 4 · Décris la réaction du public.
  Étape 5 · Relis et vérifie tes formes au Present Continuous.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — A Fashion Show in Abidjan
  Durée conseillée : 30 minutes

  CONTEXTE :
  Tu es journaliste pour le journal de ton lycée à Abidjan.
  Tu viens d'assister au premier défilé de mode organisé
  par des élèves de Terminale de ton lycée. Tu dois écrire
  un article court en anglais pour raconter cet événement.

  ÉNONCÉ :
  Rédige un article en anglais de 80 à 100 mots.
  Ton article doit contenir :
  1. Un titre en anglais.
  2. Une introduction : où et quand l'événement a-t-il eu lieu ?
  3. Un développement : décris 2 mannequins et leurs tenues
     (utilise le Present Continuous).
  4. Une conclusion : comment a réagi le public ?

  Contraintes obligatoires :
  · Minimum 4 phrases au Present Continuous
  · Minimum 4 mots de vocabulaire du défilé de mode
  · Au moins 1 prénom ivoirien et 1 lieu d'Abidjan
  · Les adjectifs AVANT les noms

  GUIDE PAS À PAS (sans corrigé)
  Étape 1 · Trouve un titre accrocheur.
            Ex : "A Wonderful Fashion Show at [Nom de ton lycée]"
  Étape 2 · Rédige l'introduction (1-2 phrases) :
            Quand ? Où ? Qui organise ?
            → Ce soir / au lycée / les élèves de Terminale
  Étape 3 · Décris le premier mannequin (2 phrases) :
            Prénom + Present Continuous + description de tenue
  Étape 4 · Décris le deuxième mannequin (2 phrases) :
            Même structure — tenue différente
  Étape 5 · Conclus en décrivant la réaction du public (1 phrase)
            → The audience is... / The students are...
  Étape 6 · Compte tes mots (80 à 100 mots) et relis.
  Étape 7 · Vérifie : am/is/are + verb-ING toujours présents ?
            Adjectifs avant les noms ?

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON

  TITRE PROPOSÉ :
  "A Brilliant Fashion Show at Lycée Moderne de Cocody"

  ARTICLE PROPOSÉ :

  This evening, a wonderful fashion show is taking place
  at Lycée Moderne de Cocody in Abidjan. The students of
  Terminale are organising the event.

  On the runway, Adjoua is walking confidently. She is wearing
  a long, colourful wax dress with golden embroidery. Another
  model, Konan, is showing a fitted blue bazin boubou with
  silver buttons.

  The audience is clapping and cheering loudly. Everyone is
  enjoying the show. The designers are smiling backstage.
  It is a great moment for the school!

  (Nombre de mots : 88 — dans les limites requises)

  TRADUCTION COMPLÈTE :
  Ce soir, un magnifique défilé de mode se déroule au Lycée
  Moderne de Cocody à Abidjan. Les élèves de Terminale
  organisent l'événement.
  Sur le podium, Adjoua marche avec assurance. Elle porte
  une longue robe en wax colorée avec une broderie dorée.
  Un autre mannequin, Konan, présente un boubou en bazin
  bleu ajusté avec des boutons argentés.
  Le public applaudit et acclame bruyamment. Tout le monde
  apprécie le spectacle. Les stylistes sourient dans les
  coulisses. C'est un grand moment pour l'école !

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1

  1→d  (runway = podium / passerelle)
  2→a  (designer = styliste)
  3→e  (audience = public / spectateurs)
  4→c  (outfit = tenue)
  5→b  (backstage = coulisses)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2

  1. dance  → dancing
     (Règle : verbe en -E muet → on retire le E, on ajoute -ING)

  2. sit    → sitting
     (Règle : verbe court CVC [consonne-voyelle-consonne]
     → on double la consonne finale T, on ajoute -ING)

  3. show   → showing
     (Règle : verbe ordinaire → on ajoute simplement -ING)

  4. write  → writing
     (Règle : verbe en -E muet → on retire le E, on ajoute -ING)

  5. run    → running
     (Règle : verbe court CVC → on double la consonne finale N,
     on ajoute -ING)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3

  1. Fatou is wearing a long red dress.
     (Fatou porte une longue robe rouge.)

  2. The photographer is taking pictures.
     (Le photographe prend des photos.)
     ⚠ Rappel : take → taking (verbe en -E muet)

  3. I am looking at the runway.
     (Je regarde le podium.)

  4. The models are walking on the stage.
     (Les mannequins marchent sur la scène.)

  5. The designer is showing his new collection.
     (Le styliste présente sa nouvelle collection.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4

  Mannequin A — Awa :
  "Awa is wearing a long green wax dress with a golden belt."
  (Awa porte une longue robe en wax verte avec une ceinture dorée.)

  Mannequin B — Issiaka :
  "Issiaka is wearing a white bazin boubou with blue embroidery."
  (Issiaka porte un boubou en bazin blanc avec des broderies bleues.)

  Mannequin C — Nadia :
  "Nadia is wearing a short black skirt with geometric patterns."
  (Nadia porte une jupe courte noire avec des motifs géométriques.)

  ✔ Vérification : adjectifs avant le nom ✔ · Present Continuous ✔
    · vocabulaire de la tenue ✔

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5

  MODÈLE DE PARAGRAPHE PROPOSÉ (non unique — d'autres
  réponses sont possibles si les contraintes sont respectées) :

  "Tonight, a wonderful fashion show is taking place at the
  Palais de la Culture in Treichville, Abidjan. On the runway,
  Mariama is walking gracefully. She is wearing a fitted
  colourful wax dress with golden embroidery. Another model,
  Sékou, is showing a long white boubou with blue patterns.
  The audience is cheering and clapping loudly. The designer
  is smiling backstage. It is a great night for Ivorian fashion!"

  TRADUCTION :
  Ce soir, un magnifique défilé de mode se déroule au Palais
  de la Culture à Treichville, Abidjan. Sur le podium, Mariama
  marche avec grâce. Elle porte une robe en wax ajustée et
  colorée avec une broderie dorée. Un autre mannequin, Sékou,
  présente un long boubou blanc avec des motifs bleus. Le public
  acclame et applaudit bruyamment. Le styliste sourit dans les
  coulisses. C'est une grande soirée pour la mode ivoirienne !

  Vérification :
  ✔ Verbes au Present Continuous (is taking place, is walking,
    is wearing, is showing, is cheering, is clapping, is smiling)
  ✔ Descriptions de tenues (Mariama + Sékou) : 2 minimum
  ✔ Vocabulaire du défilé : runway, audience, designer,
    backstage, fashion show
  ✔ Lieu ivoirien : Palais de la Culture, Treichville, Abidjan
  ✔ Prénoms ivoiriens : Mariama, Sékou
  ✔ Adjectifs avant les noms

════════════════════════════════════════════════════════
  Prof. Evelyne ASSI · Anglais LV1 · 3ème
  DPFC/MENET-FP CI · 2025-2026
  "Speak even if it's not perfect — silence teaches nothing."
  (Parle même si c'est imparfait — le silence n'apprend rien.)
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
    # Grand titre
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(0)
    p4.paragraph_format.space_after  = Pt(6)
    r4 = p4.add_run(titre)
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

    for line in sommaire_lines:
        line = line.strip()
        if not line:
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
        p_sep = cell.add_paragraph()
        p_sep.paragraph_format.space_before = Pt(5)
        p_sep.paragraph_format.space_after  = Pt(2)
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

SOUS_TITRES = re.compile(
    r'^(Objectifs|Prérequis|MÉTHODE|GUIDE|CONTRAINTES|CONTEXTE|'
    r'CONSIGNE|STRUCTURE|GUIDE PAS|Énoncé|Titre\s*:|Compétence\s*:|Professeur\s+\w)')


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
            add_decouverte_box(eng_lines, trad_lines)
        continue
    if in_tilde:
        tilde_lines.append(s); continue

    # SYNTHESE
    if s == "SYNTHESE_DEBUT":
        flush(); in_syn = True; syn_lines = []; continue
    if s == "SYNTHESE_FIN":
        in_syn = False; add_synthese_box(syn_lines); syn_lines = []; continue
    if in_syn:
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

    # [BLOC TITRE]
    m = re.match(r'^\[(.+)\]$', s)
    if m:
        flush(); bl = m.group(1); tu = bl.upper()
        if "CAPSULE" in tu:
            cur_type = "CAPSULE";  cur_title = bl
        elif "ANCRAGE" in tu:
            cur_type = "ANCRAGE";  cur_title = bl
        elif "SCHÉMA" in tu or "SCHEMA" in tu:
            cur_type = "SCHEMA";   cur_title = bl
        else:
            cur_type = "SCHEMA";   cur_title = bl
        continue

    # HARKNESS
    if re.match(r'^[▶►]\s*HARKNESS', s):
        flush()
        cur_type  = "HARKNESS"
        cur_title = s.lstrip("▶► ").strip()
        continue

    # À RETENIR
    if re.match(r'^[┌].*À RETENIR', s):
        flush()
        tc = re.sub(r'^[┌─\s]+', '', s).rstrip('─┐').strip()
        cur_type = "RETENIR"; cur_title = tc; continue
    if re.match(r'^[└]', s):
        flush(); continue

    # ATTENTION
    if re.match(r'^⚠', s):
        flush()
        cur_type  = "ATTENTION"
        cur_title = s.lstrip("⚠ ").strip()
        continue

    # TRACE ÉCRITE
    if re.match(r'^✎', s):
        flush(); cur_type = "TRACE"; cur_lines = [s]; continue

    # ANALOGIE
    if re.match(r'^~\s*ANALOGIE', s):
        if cur_type == "TRACE":
            flush()
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

    # Citation finale
    if "silence teaches nothing" in s:
        flush()
        # Extraire la citation anglaise
        m_en = re.search(r'"([^"]+)"', s)
        if m_en:
            citation_en = m_en.group(1)
        continue
    if "silence n'apprend rien" in s:
        m_fr = re.search(r'\(([^)]+)\)', s)
        if m_fr:
            citation_fr = m_fr.group(1)
        continue
    # Ligne signature finale
    if re.match(r'^Prof\.?\s+\w+.*DPFC', s):
        citation_sig = s
        continue

    # Contenu d'un bloc ouvert
    if cur_type is not None:
        cur_lines.append(s)
    else:
        # Hors bloc
        if SOUS_TITRES.match(s):
            add_sous_titre(s)
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
