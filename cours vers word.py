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
  ANGLAIS LV1 · Leçon 12 — COSMETICS (Les cosmétiques)
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

  [CAPSULE DE REPRISE — Leçon 11 : Fashion Show]

  5 points clés de la leçon précédente :
  Point 1 · Le vocabulaire d'un défilé : model (mannequin),
            runway (podium), designer (créateur), collection,
            outfit (tenue)
  Point 2 · Décrire une tenue : "She is wearing a long red dress
            with golden accessories."
            (Elle porte une longue robe rouge avec des accessoires dorés.)
  Point 3 · Le Present Continuous pour décrire une action en cours :
            "The model is walking down the runway."
            (Le mannequin descend le podium.)
  Point 4 · Exprimer une opinion : "I think this outfit is
            elegant." (Je pense que cette tenue est élégante.)
  Point 5 · Les adjectifs de mode : stylish (élégant), trendy
            (à la mode), gorgeous (magnifique), casual (décontracté)

  3 questions flash :
  Q1 : How do you say "mannequin" in English?
       → A model
  Q2 : Translate : "Les créateurs ivoiriens sont très talentueux."
       → Ivorian designers are very talented.
  Q3 : Put in the Present Continuous :
       "She (wear) a traditional wax outfit."
       → She is wearing a traditional wax outfit.

  Lien avec la leçon du jour :
  En L11, on a parlé des vêtements lors d'un défilé.
  Aujourd'hui on va plus loin : les cosmétiques complètent
  une tenue — maquillage, parfums, soins de beauté font
  partie du même univers de la mode (C4).

────────────────────────────────────────────────────────
  [ANCRAGE IVOIRIEN]

  À Abidjan, le marché de Cocody et les boutiques du
  Plateau regorgent de produits cosmétiques : crèmes,
  huiles de coco, beurre de karité, henné pour les
  cheveux. Des marques ivoiriennes comme "Bloubana"
  ou des vendeuses au marché d'Adjamé proposent des
  cosmétiques naturels très prisés.
  Le karité (shea butter) vient du Nord de la CI
  et s'exporte dans le monde entier — c'est un
  cosmétique ivoirien de renommée mondiale !
  Lien DPFC : les cosmétiques non contrôlés peuvent
  contenir des produits dangereux pour la peau —
  thème santé/civisme.

────────────────────────────────────────────────────────
  [HARKNESS]

▶ HARKNESS 1 — Beauty standards / Les critères de beauté
  Q : "Do you think cosmetics are necessary to feel
      beautiful? Why or why not?"
      (Penses-tu que les cosmétiques sont nécessaires
      pour se sentir beau/belle ? Pourquoi ?)
  Guide :
    Piste 1 · La beauté naturelle vs la beauté aidée
              par les produits
    Piste 2 · L'influence des réseaux sociaux et
              de la publicité sur nos choix
    Piste 3 · Les risques des produits dépigmentants
              très utilisés en Afrique
  Corrigé :
    Cosmetics are not necessary to feel beautiful,
    but they can boost confidence. (Les cosmétiques ne
    sont pas nécessaires pour se sentir beau, mais ils
    peuvent renforcer la confiance en soi.)
    However, some products like skin-lightening creams
    are dangerous for health. (Cependant, certains
    produits comme les crèmes éclaircissantes sont
    dangereux pour la santé.)
    True beauty comes from within, but cosmetics can
    enhance natural features. (La vraie beauté vient
    de l'intérieur, mais les cosmétiques peuvent
    mettre en valeur les traits naturels.)

▶ HARKNESS 2 — Natural vs chemical cosmetics
  Q : "What is the difference between natural cosmetics
      and chemical cosmetics?"
      (Quelle est la différence entre les cosmétiques
      naturels et les cosmétiques chimiques ?)
  Guide :
    Piste 1 · Les ingrédients : plantes vs laboratoire
    Piste 2 · Les effets sur la peau à long terme
    Piste 3 · Le prix et l'accessibilité en CI
  Corrigé :
    Natural cosmetics are made from plants and minerals.
    (Les cosmétiques naturels sont fabriqués à partir
    de plantes et de minéraux.)
    Examples in Côte d'Ivoire : shea butter, coconut
    oil, black soap. (Exemples en CI : beurre de karité,
    huile de coco, savon noir.)
    Chemical cosmetics are produced in laboratories and
    may contain harmful substances. (Les cosmétiques
    chimiques sont produits en laboratoire et peuvent
    contenir des substances nocives.)

▶ HARKNESS 3 — Cosmetics and identity
  Q : "Can cosmetics reflect a person's culture
      or identity?"
      (Les cosmétiques peuvent-ils refléter la culture
      ou l'identité d'une personne ?)
  Guide :
    Piste 1 · Le henné utilisé dans les cérémonies
    Piste 2 · Les peintures corporelles traditionnelles
    Piste 3 · Les habits de fête et le maquillage
              dans les cultures ivoiriennes
  Corrigé :
    Yes, cosmetics can reflect culture. (Oui, les
    cosmétiques peuvent refléter la culture.)
    For example, Ivorian women use henna for weddings
    and ceremonies. (Par exemple, les femmes ivoiriennes
    utilisent le henné pour les mariages et cérémonies.)
    Body painting is also used in traditional festivals.
    (La peinture corporelle est aussi utilisée lors des
    fêtes traditionnelles.)
    So cosmetics are not only about beauty — they
    express identity and belonging. (Donc les cosmétiques
    ne sont pas seulement une question de beauté —
    ils expriment l'identité et l'appartenance.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

  Phase 1 · Présentation / Prérequis
  ───────────────────────────────────

  Titre : COSMETICS (Les cosmétiques)
  Compétence C4 · Écrit élaboré · Thème : LA MODE / FASHION

  Objectifs de la leçon :
  → Apprendre le vocabulaire des cosmétiques en anglais
  → Utiliser le Present Perfect pour parler d'expériences
  → Décrire des produits de beauté à l'écrit
  → Exprimer une opinion sur les cosmétiques

  Prérequis nécessaires :
  → Connaître le vocabulaire de base des vêtements (L10-L11)
  → Savoir utiliser le verbe "to be" et "to have"
  → Connaître les adjectifs courants (beautiful, natural,
    expensive, cheap...)
  → Savoir former une phrase simple au présent

────────────────────────────────────────────────────────

  Phase 2 · Découverte guidée
  ────────────────────────────

  VOCABULAIRE ESSENTIEL — LES COSMÉTIQUES

  ◆ Les produits de beauté / Beauty products :

  lipstick        → le rouge à lèvres    [prononciation : LIP-stik]
  foundation      → le fond de teint     [prononciation : fawn-DAY-shun]
  mascara         → le mascara           [prononciation : mas-KA-ra]
  eyeliner        → l'eye-liner          [prononciation : AÏ-laï-neur]
  perfume         → le parfum            [prononciation : pur-FYOUM]
  nail polish     → le vernis à ongles   [prononciation : NAÏL po-lish]
  moisturizer     → la crème hydratante  [prononciation : MOYS-chu-raï-zer]
  sunscreen       → la crème solaire     [prononciation : SUN-skrine]
  shampoo         → le shampoing         [prononciation : sham-POU]
  deodorant       → le déodorant         [prononciation : di-O-do-rant]

  ◆ Les cosmétiques naturels / Natural cosmetics :

  shea butter     → le beurre de karité  [prononciation : SHIA bu-teur]
  coconut oil     → l'huile de coco      [prononciation : KO-ko-nut oïl]
  black soap      → le savon noir        [prononciation : blak sop]
  henna           → le henné             [prononciation : HE-na]
  aloe vera       → l'aloe vera          [prononciation : A-lo VI-ra]

  ◆ Les parties du corps concernées / Body parts :

  skin            → la peau              [prononciation : skin]
  hair            → les cheveux          [prononciation : hèr]
  lips            → les lèvres           [prononciation : lips]
  nails           → les ongles           [prononciation : naïlz]
  face            → le visage            [prononciation : féïs]
  eyes            → les yeux             [prononciation : aïz]

┌─ À RETENIR ──────────────────────────────────────────┐
  FAUX AMI IMPORTANT :
  "Cosmetic" en anglais = produit de beauté en général
  Ce n'est PAS uniquement le maquillage !
  Un shampooing, une crème solaire, un déodorant
  sont tous des "cosmetics".
  En français, on dirait "produits cosmétiques"
  ou "produits de beauté".
└──────────────────────────────────────────────────────┘

⚠ ATTENTION
  "Perfume" se prononce [pur-FYOUM] en anglais,
  PAS [per-fum] comme en français.
  "Foundation" ne veut PAS dire "fondation"
  (un bâtiment) — ici c'est le fond de teint !
  → Faux ami à retenir absolument.

────────────────────────────────────────────────────────

  GRAMMAIRE — LE PRESENT PERFECT
  (pour parler d'expériences passées)

  Formation :
  Sujet + HAVE / HAS + participe passé du verbe

  → I / You / We / They → HAVE + participe passé
  → He / She / It       → HAS  + participe passé

  Exemples avec les cosmétiques :

  I have tried shea butter before.
  (J'ai déjà essayé le beurre de karité.)

  She has bought a new lipstick.
  (Elle a acheté un nouveau rouge à lèvres.)

  We have never used chemical products.
  (Nous n'avons jamais utilisé de produits chimiques.)

  Have you ever used coconut oil?
  (As-tu déjà utilisé de l'huile de coco ?)
  → Yes, I have. / No, I have never.
    (Oui. / Non, jamais.)

┌─ À RETENIR ──────────────────────────────────────────┐
  MOTS CLÉS DU PRESENT PERFECT :
  → already (déjà)  : I have already bought perfume.
  → never  (jamais) : She has never used mascara.
  → ever   (déjà — en question) : Have you ever tried...?
  → just   (venir de) : He has just bought shampoo.
  → yet    (encore/déjà en négatif/question) :
    I haven't bought it yet. (Je ne l'ai pas encore acheté.)
└──────────────────────────────────────────────────────┘

⚠ ATTENTION
  Simple Past vs Present Perfect :
  Simple Past → action passée à un moment PRÉCIS
  "I bought lipstick yesterday." (hier = moment précis)

  Present Perfect → expérience sans moment précis,
  ou action récente avec résultat présent.
  "I have bought lipstick." (sans dire quand exactement)

────────────────────────────────────────────────────────

  Phase 3 · Schémas textuels
  ───────────────────────────

[SCHÉMA 1 — Carte mentale : COSMETICS]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  • Centre (cercle, couleur rose) : COSMETICS
  • Branche 1 (gauche, violet) : MAKEUP / MAQUILLAGE
    → lipstick · mascara · foundation · eyeliner
  • Branche 2 (droite, vert) : NATURAL / NATURELS
    → shea butter · coconut oil · black soap · henna
  • Branche 3 (haut, bleu) : HAIR CARE / CHEVEUX
    → shampoo · conditioner · hair oil
  • Branche 4 (bas, orange) : SKIN CARE / PEAU
    → moisturizer · sunscreen · deodorant · perfume
  • Chaque terme anglais accompagné de sa traduction FR
  • Légende : vert = naturel · rose = maquillage
Note générateur : Canva, draw.io ou MindMeister
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Formation du Present Perfect]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  • 3 colonnes :
    Colonne 1 (bleu) : SUJET
      → I / You / We / They
      → He / She / It
    Colonne 2 (vert) : AUXILIAIRE
      → HAVE
      → HAS
    Colonne 3 (orange) : PARTICIPE PASSÉ
      → tried / bought / used / seen / never used
  • Flèches reliant les colonnes de gauche à droite
  • En bas : exemple complet coloré
    [I] + [have] + [tried] = "I have tried shea butter."
  • Zone rouge "ATTENTION" : He/She/It → HAS (pas HAVE)
Note générateur : Canva, PowerPoint, draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

  Phase 4 · Définitions DPFC
  ───────────────────────────

◆ DÉFINITION DPFC — Cosmetics (Les cosmétiques)
  Cosmetics are products used to clean, protect
  or enhance the appearance of the face, hair,
  skin or body.
  (Les cosmétiques sont des produits utilisés pour
  nettoyer, protéger ou améliorer l'apparence du
  visage, des cheveux, de la peau ou du corps.)

◆ DÉFINITION DPFC — Natural cosmetics
  Natural cosmetics are beauty products made from
  plants, minerals or animal products without
  artificial chemicals.
  (Les cosmétiques naturels sont des produits de
  beauté fabriqués à partir de plantes, de minéraux
  ou de produits animaux sans produits chimiques
  artificiels.)

◆ DÉFINITION DPFC — Present Perfect
  The Present Perfect is a verb tense used to talk
  about past experiences or recent actions that
  have a connection to the present.
  (Le Present Perfect est un temps verbal utilisé
  pour parler d'expériences passées ou d'actions
  récentes qui ont un lien avec le présent.)
  Formation : Subject + have/has + past participle.

────────────────────────────────────────────────────────

  Phase 5 · Trace écrite + Analogies CI
  ──────────────────────────────────────

✎ TRACE ÉCRITE (à recopier)

  LEÇON 12 — COSMETICS (Les cosmétiques)

  I. VOCABULAIRE ESSENTIEL
  lipstick = le rouge à lèvres
  foundation = le fond de teint
  mascara = le mascara
  perfume = le parfum
  moisturizer = la crème hydratante
  shea butter = le beurre de karité
  coconut oil = l'huile de coco
  henna = le henné

  II. GRAMMAIRE — PRESENT PERFECT
  Formation : Sujet + HAVE/HAS + participe passé

  I/You/We/They + HAVE + participe passé
  He/She/It     + HAS  + participe passé

  Exemples :
  (+) I have tried coconut oil. (J'ai essayé l'huile de coco.)
  (-) She has never used mascara. (Elle n'a jamais utilisé de mascara.)
  (?) Have you ever bought perfume? (As-tu déjà acheté du parfum ?)

  Mots clés : already · never · ever · just · yet

  III. DÉCRIRE UN PRODUIT COSMÉTIQUE
  Structure : This product is + adjectif.
              It is used for + nom/verbe-ing.
              It is made from + matière.

  Exemple :
  Shea butter is a natural cosmetic.
  It is used for moisturizing the skin.
  It is made from the shea tree.
  (Le beurre de karité est un cosmétique naturel.
   Il est utilisé pour hydrater la peau.
   Il est fabriqué à partir de l'arbre à karité.)

~ ANALOGIE CI
  Le Present Perfect, c'est comme raconter à ta
  maman que tu as DÉJÀ fait quelque chose, sans
  préciser l'heure exacte :
  "Maman, j'ai déjà fait mes devoirs !"
  → "Mum, I have already done my homework!"
  Tu ne dis pas QUAND exactement — juste que c'est
  fait et que ça a un résultat maintenant.

~ ANALOGIE CI
  Pense au marché d'Adjamé : la vendeuse dit
  "J'ai déjà vendu 10 pots de karité aujourd'hui"
  sans dire exactement à quelle heure.
  En anglais : "She has already sold 10 pots of
  shea butter today."
  → Present Perfect car le résultat (les ventes)
  compte plus que le moment exact.

────────────────────────────────────────────────────────

  Phase 5b · Synthèse
  ─────────────────────

  5 points essentiels à retenir :
  Point 1 · Les cosmétiques incluent le maquillage,
            les soins de la peau, des cheveux et les parfums
  Point 2 · La CI produit des cosmétiques naturels
            mondialement reconnus : karité, huile de coco
  Point 3 · Le Present Perfect = have/has + participe passé
  Point 4 · Mots-clés du Present Perfect :
            already, never, ever, just, yet
  Point 5 · Pour décrire un produit : It is used for... /
            It is made from... / It is good for...

  5 mots clés + définition courte :
  Cosmetics   → produits de beauté
  Lipstick    → rouge à lèvres (maquillage des lèvres)
  Shea butter → beurre de karité (cosmétique naturel CI)
  Moisturizer → crème hydratante (soin de la peau)
  Present Perfect → temps pour les expériences passées
                    (have/has + participe passé)

  3 erreurs fréquentes + correction :
  Erreur 1 · "She have tried..." ✗
             → "She HAS tried..." ✓
             (She = he/she/it → toujours HAS)
  Erreur 2 · "Foundation" = fondation (bâtiment) ✗
             → "Foundation" = fond de teint ✓
             (Faux ami !)
  Erreur 3 · "I have bought it yesterday." ✗
             → "I bought it yesterday." ✓
             (Avec "yesterday" = moment précis
             → Simple Past, PAS Present Perfect)

────────────────────────────────────────────────────────

  Phase 6 · Exercices guidés
  ───────────────────────────

◉ EXERCICE GUIDÉ 1 — Describe a cosmetic product
  Notion ciblée : Vocabulaire + structure descriptive

  Énoncé :
  Décris le beurre de karité (shea butter) en anglais
  en utilisant la structure suivante :
  — Ce que c'est (what it is)
  — À quoi ça sert (what it is used for)
  — D'où ça vient (where it comes from)
  — Pourquoi c'est bon (why it is good)

  MÉTHODE
  Étape 1 · Nomme le produit et sa catégorie
    → "Shea butter is a natural cosmetic."
  Étape 2 · Dis à quoi il sert avec "used for + V-ing"
    → "It is used for moisturizing the skin and hair."
  Étape 3 · Indique l'origine avec "made from" ou
    "comes from"
    → "It comes from the shea tree in northern
       Côte d'Ivoire."
  Étape 4 · Donne un avantage avec "good for"
    → "It is good for dry skin and natural hair."
  Étape 5 · Ajoute une phrase d'opinion
    → "I think shea butter is the best natural
       cosmetic in West Africa."
  ✔ Vérification : vérifie que tu as utilisé
    "is used for + V-ing" et "is made from / comes from"
  ✔ Conclusion attendue :
    Shea butter is a natural cosmetic.
    It is used for moisturizing the skin and hair.
    It comes from the shea tree in northern
    Côte d'Ivoire.
    It is good for dry skin and natural hair.
    I think shea butter is the best natural cosmetic
    in West Africa.

◉ EXERCICE GUIDÉ 2 — Present Perfect
  Notion ciblée : Formation et emploi du Present Perfect

  Énoncé :
  Amoin, une lycéenne d'Abidjan, parle de ses
  expériences avec les cosmétiques.
  Mets les verbes entre parenthèses au Present Perfect.

  1. Amoin ___ (try) shea butter for the first time.
  2. She ___ (buy) a new lipstick at the Plateau.
  3. Her mother ___ (never / use) chemical creams.
  4. ___ you ever ___ (see) a cosmetics fair in Abidjan?
  5. We ___ (just / discover) coconut oil for our hair.

  MÉTHODE
  Étape 1 · Identifie le sujet de chaque phrase
    → Amoin = She → HAS | you → HAVE | We → HAVE
  Étape 2 · Forme : sujet + have/has + participe passé
    → try → tried | buy → bought | use → used
    → see → seen | discover → discovered
  Étape 3 · Place "never" et "just" entre have/has
    et le participe passé
    → has never used | have just discovered
  Étape 4 · Pour la question, inverse sujet et have/has
    → Have you ever seen...?
  ✔ Vérification : relis chaque phrase — sujet correct ?
    have ou has ? participe passé correct ?
  ✔ Corrigé :
    1. Amoin has tried shea butter for the first time.
    2. She has bought a new lipstick at the Plateau.
    3. Her mother has never used chemical creams.
    4. Have you ever seen a cosmetics fair in Abidjan?
    5. We have just discovered coconut oil for our hair.

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulaire des cosmétiques
  Notions travaillées : Vocabulaire produits de beauté

  Relie chaque mot anglais à sa traduction française.
  (Écris le numéro et la lettre correspondante.)

  1. Lipstick          a. La crème hydratante
  2. Shampoo           b. Le vernis à ongles
  3. Moisturizer       c. Le rouge à lèvres
  4. Nail polish       d. Le parfum
  5. Perfume           e. Le shampooing

  GUIDE
  Étape 1 · Lis chaque mot anglais à voix haute
  Étape 2 · Cherche dans ta trace écrite le mot
            que tu ne connais pas
  Étape 3 · Associe par élimination si tu bloques

◎ EXERCICE 2 — Present Perfect : conjugaison
  Notions travaillées : Formation have/has + participe passé

  Conjugue les verbes au Present Perfect.

  1. I ___ (use) coconut oil before.
  2. Koffi ___ (buy) a perfume for his mother.
  3. They ___ (never / try) henna tattoos.
  4. ___ she ___ (ever / wear) lipstick?
  5. We ___ (just / find) a natural soap at the market.

  GUIDE
  Étape 1 · Repère le sujet (I, Koffi, they, she, we)
  Étape 2 · Choisis HAVE ou HAS selon le sujet
  Étape 3 · Forme le participe passé du verbe
  Étape 4 · Place never/just/ever au bon endroit

◎ EXERCICE 3 — Simple Past ou Present Perfect ?
  Notions travaillées : Distinguer Simple Past et Present Perfect

  Choisis le bon temps pour chaque phrase.

  1. Aya ___ (buy) mascara yesterday.
  2. I ___ (never / see) such a beautiful perfume.
  3. They ___ (visit) the cosmetics fair last week.
  4. She ___ (just / finish) her makeup.
  5. We ___ (use) shea butter when we were children.

  GUIDE
  Étape 1 · Cherche un indicateur de temps dans la phrase
            (yesterday, last week = Simple Past)
            (never, just, already = Present Perfect)
  Étape 2 · Si pas d'indicateur : est-ce une expérience
            générale ou un moment précis ?
  Étape 3 · Applique la règle et conjugue

◎ EXERCICE 4 — Décrire un produit
  Notions travaillées : Structure descriptive écrite (C4)

  Décris le savon noir ivoirien (black soap) en
  5 phrases en anglais.
  Utilise : is used for · is made from · is good for ·
  I think · It comes from

  GUIDE
  Étape 1 · Commence par : "Black soap is a natural
            cosmetic from Côte d'Ivoire."
  Étape 2 · Dis à quoi il sert (used for + V-ing)
  Étape 3 · Explique d'où il vient et comment
            il est fait (made from)
  Étape 4 · Donne un bénéfice (good for)
  Étape 5 · Termine par ton opinion personnelle

◎ EXERCICE 5 — Expression écrite intégratrice
  Notions travaillées : Vocabulaire + Present Perfect +
                        Structure descriptive + Opinion

  Sujet :
  Ta camarade Adjoua t'écrit un message pour te demander
  quel cosmétique naturel ivoirien tu lui recommandes.
  Écris-lui une réponse de 8 à 10 phrases en anglais.

  Dans ta réponse, tu dois :
  → Nommer le produit et le décrire
  → Dire pourquoi tu l'as essayé (Present Perfect)
  → Expliquer ses avantages
  → Lui conseiller comment l'utiliser
  → Donner ton opinion finale

  GUIDE
  Étape 1 · Commence par saluer Adjoua et annoncer
            ton choix de produit
  Étape 2 · Décris le produit (what it is, where it
            comes from, what it is made from)
  Étape 3 · Parle de ton expérience au Present Perfect :
            "I have tried..." / "I have used..."
  Étape 4 · Explique les bénéfices avec "good for"
            et "used for"
  Étape 5 · Donne un conseil pratique
  Étape 6 · Termine par une opinion : "I think..." /
            "In my opinion..."

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — My beauty routine with natural cosmetics
  Durée conseillée : 30 minutes

  Sujet :
  Décris ta routine beauté idéale en utilisant
  uniquement des cosmétiques naturels ivoiriens.
  Rédige un texte de 10 à 12 phrases en anglais.

  Contraintes obligatoires :
  → Utilise au moins 6 mots de vocabulaire de la leçon
  → Utilise le Present Perfect au moins 3 fois
  → Utilise les structures : used for · made from ·
    good for · I think / In my opinion
  → Mentionne au moins 2 produits naturels ivoiriens
  → Respecte l'organisation : introduction (1 phrase) ·
    développement (8-10 phrases) · conclusion (1 phrase)

  GUIDE PAS À PAS
  Étape 1 · Introduction : présente le sujet en 1 phrase
    → "I prefer to use natural cosmetics in my
       beauty routine."
  Étape 2 · Présente ton premier produit :
    nomme-le, décris-le, dis d'où il vient
  Étape 3 · Parle de ton expérience avec ce produit
    au Present Perfect : "I have used... I have tried..."
  Étape 4 · Présente ton deuxième produit de la même façon
  Étape 5 · Explique les avantages des cosmétiques
    naturels vs chimiques (good for / bad for)
  Étape 6 · Donne ton opinion personnelle : I think...
  Étape 7 · Conclusion : résume en 1 phrase pourquoi
    tu recommandes les cosmétiques naturels

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON

  Modèle de production attendue :

  I prefer to use natural cosmetics in my beauty routine.
  (Je préfère utiliser des cosmétiques naturels dans
  ma routine beauté.)

  My first favourite product is shea butter.
  (Mon premier produit préféré est le beurre de karité.)
  It is a natural cosmetic made from the shea tree.
  (C'est un cosmétique naturel fabriqué à partir de
  l'arbre à karité.)
  It comes from the north of Côte d'Ivoire.
  (Il vient du nord de la Côte d'Ivoire.)
  I have used shea butter for two years now.
  (J'utilise le beurre de karité depuis deux ans maintenant.)
  It is used for moisturizing the skin and the hair.
  (Il est utilisé pour hydrater la peau et les cheveux.)
  It is good for dry skin and damaged hair.
  (Il est bon pour la peau sèche et les cheveux abîmés.)

  My second product is black soap.
  (Mon deuxième produit est le savon noir.)
  I have tried black soap for the first time last month.
  (J'ai essayé le savon noir pour la première fois
  le mois dernier.)
  It is made from palm oil and ash.
  (Il est fabriqué à partir d'huile de palme et de cendre.)
  It is used for cleaning the face and body deeply.
  (Il est utilisé pour nettoyer le visage et le corps
  en profondeur.)

  I think natural cosmetics are better than chemical
  products because they are safe for our health.
  (Je pense que les cosmétiques naturels sont meilleurs
  que les produits chimiques car ils sont sûrs pour
  notre santé.)

  In conclusion, I strongly recommend natural Ivorian
  cosmetics to everyone.
  (En conclusion, je recommande vivement les cosmétiques
  naturels ivoiriens à tout le monde.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1 (Vocabulaire)

  1 → c : Lipstick = le rouge à lèvres
  2 → e : Shampoo = le shampooing
  3 → a : Moisturizer = la crème hydratante
  4 → b : Nail polish = le vernis à ongles
  5 → d : Perfume = le parfum

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2 (Present Perfect : conjugaison)

  Étape 1 · Identifier sujets et auxiliaires :
    I → HAVE | Koffi → HAS | They → HAVE
    she → HAS | We → HAVE

  Étape 2 · Participes passés :
    use → used | buy → bought | try → tried
    wear → worn | find → found

  Étape 3 · Résultats :
  1. I have used coconut oil before.
     (J'ai déjà utilisé l'huile de coco avant.)
  2. Koffi has bought a perfume for his mother.
     (Koffi a acheté un parfum pour sa mère.)
  3. They have never tried henna tattoos.
     (Ils n'ont jamais essayé les tatouages au henné.)
  4. Has she ever worn lipstick?
     (A-t-elle jamais porté du rouge à lèvres ?)
  5. We have just found a natural soap at the market.
     (Nous venons juste de trouver un savon naturel
     au marché.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3 (Simple Past ou Present Perfect ?)

  Étape 1 · Analyse des indicateurs :
  Phrase 1 : "yesterday" → indicateur de temps précis
             → Simple Past
  Phrase 2 : "never" → expérience sans moment précis
             → Present Perfect
  Phrase 3 : "last week" → moment précis
             → Simple Past
  Phrase 4 : "just" → action récente
             → Present Perfect
  Phrase 5 : "when we were children" → moment passé précis
             → Simple Past

  Étape 2 · Corrigé :
  1. Aya bought mascara yesterday.
     (Aya a acheté du mascara hier.)
  2. I have never seen such a beautiful perfume.
     (Je n'ai jamais vu un parfum aussi beau.)
  3. They visited the cosmetics fair last week.
     (Ils ont visité le salon des cosmétiques la semaine
     dernière.)
  4. She has just finished her makeup.
     (Elle vient juste de finir son maquillage.)
  5. We used shea butter when we were children.
     (Nous utilisions le beurre de karité quand nous
     étions enfants.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4 (Décrire un produit)

  Modèle attendu :

  Black soap is a natural cosmetic from Côte d'Ivoire.
  (Le savon noir est un cosmétique naturel de CI.)
  It is used for cleaning the face and removing
  impurities from the skin.
  (Il est utilisé pour nettoyer le visage et enlever
  les impuretés de la peau.)
  It is made from palm oil, shea butter and ash.
  (Il est fabriqué à partir d'huile de palme, de
  beurre de karité et de cendre.)
  It comes from traditional Ivorian villages.
  (Il vient des villages traditionnels ivoiriens.)
  It is good for oily skin and skin problems like acne.
  (Il est bon pour la peau grasse et les problèmes
  de peau comme l'acné.)
  I think black soap is one of the best natural
  cosmetics in West Africa.
  (Je pense que le savon noir est l'un des meilleurs
  cosmétiques naturels en Afrique de l'Ouest.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5 (Expression écrite)

  Modèle de réponse à Adjoua :

  Hello Adjoua!
  (Bonjour Adjoua !)
  I am happy to recommend a natural Ivorian cosmetic
  to you.
  (Je suis heureuse de te recommander un cosmétique
  naturel ivoirien.)

  I recommend coconut oil (l'huile de coco).
  (Je te recommande l'huile de coco.)
  It is a natural cosmetic made from coconuts.
  (C'est un cosmétique naturel fabriqué à partir
  de noix de coco.)
  It comes from the south of Côte d'Ivoire.
  (Il vient du sud de la Côte d'Ivoire.)

  I have tried coconut oil for my hair and it is
  amazing!
  (J'ai essayé l'huile de coco pour mes cheveux
  et c'est incroyable !)
  I have used it every week for six months now.
  (Je l'utilise chaque semaine depuis six mois maintenant.)
  I have also recommended it to my sister Mariame.
  (Je l'ai aussi recommandé à ma sœur Mariame.)

  It is used for moisturizing the hair and the skin.
  (Elle est utilisée pour hydrater les cheveux et
  la peau.)
  It is good for dry hair and for cooking too!
  (Elle est bonne pour les cheveux secs et aussi
  pour cuisiner !)

  I think coconut oil is the best natural cosmetic
  because it is cheap, natural and very effective.
  (Je pense que l'huile de coco est le meilleur
  cosmétique naturel car elle est bon marché,
  naturelle et très efficace.)

  Try it, you will not be disappointed!
  (Essaie-la, tu ne seras pas déçue !)
  Your friend, [ton prénom].

════════════════════════════════════════════════════════
  FIN DE LA LEÇON 12 — COSMETICS
  Prof. Evelyne ASSI · Anglais LV1 · 3ème
  DPFC/MENET-FP CI 2025-2026
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

    # Citation finale — bloc ════...════ final
    if re.match(r'^[═]{4,}', s) and i > body_start + 10:
        # Second bloc ════ = pied de cours → on ignore les suivants
        # (déjà géré par le séparateur, mais on protège la citation)
        flush(); continue
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
    # Ligne signature finale Prof. … · DPFC/… ou Prof. … · Anglais…
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
