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

COURS = """✅ Français · Leçon 13 — Texte argumentatif II — réfuter · en cours de génération...

---

LPA
Lycée Personnel Autonome

Français · Leçon 13 · Classe de 3ème

TEXTE ARGUMENTATIF II — RÉFUTER
Texte argumentatif II — Réfuter une thèse

Compétence C3 · Traiter des situations relatives à la rédaction d'écrits divers

Professeur Adjoua KONAN · Coefficient 4
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  FRANÇAIS · Leçon 13 — TEXTE ARGUMENTATIF II — RÉFUTER
  Niveau : 3ème | Série : Programme commun
  Professeur : Adjoua KONAN · Coefficient : 4
  Compétence : C3 · Expression écrite — Rédaction d'écrits divers
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 12 : Texte argumentatif I — étayer)
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

  CAPSULE DE REPRISE — Leçon 12 : Texte argumentatif I — Étayer

  5 points clés à retenir :
  Point 1 · Étayer signifie soutenir une thèse en apportant des arguments et des exemples qui la renforcent.
  Point 2 · La thèse est la position que l'auteur défend : elle doit être clairement énoncée dès l'introduction.
  Point 3 · Chaque argument est illustré par un exemple concret, précis et pertinent.
  Point 4 · Le schéma argumentatif d'un texte qui étaye : Thèse → Argument 1 + Exemple → Argument 2 + Exemple → Conclusion qui confirme la thèse.
  Point 5 · Le plan en 3 parties (Introduction / Développement / Conclusion) est obligatoire au BEPC.

  3 questions flash :
  Q1 · Qu'est-ce qu'étayer une thèse ?
  R1 · C'est la soutenir en apportant des arguments et des exemples qui prouvent qu'elle est vraie.

  Q2 · Quelle est la différence entre un argument et un exemple ?
  R2 · L'argument est une raison logique ; l'exemple est un fait concret qui illustre l'argument.

  Q3 · Comment se termine un texte argumentatif qui étaye ?
  R3 · Par une conclusion qui confirme la thèse de départ.

  Lien avec la leçon du jour :
  Aujourd'hui, on fait le chemin inverse : au lieu de soutenir une thèse, on va l'attaquer et la démolir. C'est la réfutation.

────────────────────────────────────────────────────────

  ANCRAGE IVOIRIEN

  Aminata et Kouadio sont deux élèves de 3ème au Lycée Classique d'Abidjan.
  Aminata affirme : "Les élèves qui regardent la télévision le soir travaillent mieux en classe."
  Kouadio n'est pas du tout d'accord. Il veut réfuter cette affirmation.
  Comment va-t-il s'y prendre ? Quels arguments peut-il utiliser pour montrer qu'Aminata a tort ?
  C'est exactement ce que vous allez apprendre aujourd'hui : réfuter une thèse de manière organisée et convaincante.

────────────────────────────────────────────────────────

▶ HARKNESS — La réfutation dans la vie quotidienne
  Q : Dans la vie de tous les jours en Côte d'Ivoire, cite une situation où quelqu'un réfute l'opinion d'une autre personne. Comment fait-il pour convaincre ?
  Guide :
    Piste 1 · Pense à un marché d'Abidjan où un acheteur refuse le prix du vendeur.
    Piste 2 · Pense à un parent qui refuse les arguments d'un enfant.
    Piste 3 · Réfuter, c'est dire "tu as tort" ET expliquer pourquoi.
  Corrigé : Au marché de Cocody, une cliente dit que les tomates sont trop chères. Le vendeur dit "elles coûtent 500 FCFA". La cliente réfute : "Non, c'est trop cher, car hier elles coûtaient 300 FCFA au marché de Yopougon, et la qualité est la même." Elle utilise un contre-argument (le prix d'hier) et un exemple (marché de Yopougon) pour démontrer que la thèse du vendeur est fausse.

▶ HARKNESS — Étayer ou réfuter ?
  Q : Quelle est la différence fondamentale entre un texte qui étaye et un texte qui réfute ?
  Guide :
    Piste 1 · Dans étayer, on est pour la thèse. Dans réfuter, on est contre.
    Piste 2 · Réfuter nécessite de citer la thèse adverse avant de l'attaquer.
    Piste 3 · Pense au verbe "démolir" : réfuter, c'est démolir les arguments de l'autre.
  Corrigé : Étayer, c'est construire une démonstration pour prouver qu'une thèse est vraie. Réfuter, c'est démontrer qu'une thèse adverse est fausse ou insuffisante. Dans un texte qui réfute, l'auteur présente d'abord la thèse qu'il combat, puis la détruit argument par argument grâce à des contre-arguments et des exemples contraires.

▶ HARKNESS — La force d'un contre-argument
  Q : Pour réfuter efficacement, suffit-il de dire "je ne suis pas d'accord" ? Pourquoi ?
  Guide :
    Piste 1 · Dire "je ne suis pas d'accord" sans explication, c'est une opinion, pas un argument.
    Piste 2 · Un bon contre-argument apporte une preuve ou un fait contraire.
    Piste 3 · Pense à ce que ferait un avocat face à un accusateur.
  Corrigé : Non, il ne suffit pas de dire "je ne suis pas d'accord". Pour réfuter, il faut : premièrement, nommer précisément la thèse adverse ; deuxièmement, produire un contre-argument logique qui démontre l'erreur ou la faiblesse de cette thèse ; troisièmement, appuyer ce contre-argument par un exemple concret ou un fait vérifiable. Sans ces trois éléments, la réfutation n'est pas convaincante.

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

  Titre : Texte argumentatif II — Réfuter une thèse
  Compétence DPFC : C3 — Expression écrite

  Objectifs de la leçon :
  Objectif 1 · Comprendre ce que signifie réfuter une thèse.
  Objectif 2 · Identifier la thèse adverse et les outils de la réfutation.
  Objectif 3 · Construire un schéma argumentatif de réfutation.
  Objectif 4 · Rédiger un texte argumentatif complet qui réfute une thèse.

  Prérequis nécessaires :
  · Savoir ce qu'est une thèse, un argument, un exemple (Leçon 12).
  · Connaître le plan en 3 parties : Introduction / Développement / Conclusion.
  · Savoir utiliser des connecteurs logiques (car, parce que, donc, ainsi…).
  · Savoir distinguer fait et opinion.

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

  Lisez attentivement ce court texte :

  TEXTE DE RÉFÉRENCE
  "Certains affirment que les jeunes Ivoiriens ne lisent plus parce qu'ils passent trop de temps sur les réseaux sociaux. Cette affirmation mérite d'être nuancée. En effet, de nombreux jeunes utilisent justement les réseaux sociaux pour accéder à des articles, des livres en ligne et des informations culturelles. À Abidjan, plusieurs groupes WhatsApp d'étudiants partagent chaque semaine des textes littéraires. De plus, les plateformes numériques ont rendu la lecture plus accessible que jamais, notamment pour les élèves qui n'ont pas les moyens d'acheter des livres papier. Prétendre que les réseaux sociaux nuisent à la lecture, c'est ignorer cette réalité nouvelle."

  ANALYSE DU TEXTE

  Étape 1 · Identifier la thèse adverse (la thèse que l'auteur combat) :
  → "Les jeunes Ivoiriens ne lisent plus à cause des réseaux sociaux."
  C'est la thèse adverse : l'auteur ne la soutient pas, il va la réfuter.

  Étape 2 · Identifier les contre-arguments :
  Contre-argument 1 : Les réseaux sociaux donnent accès à des contenus culturels et littéraires.
  Contre-argument 2 : Les plateformes numériques rendent la lecture plus accessible.

  Étape 3 · Identifier les exemples :
  Exemple 1 : Les groupes WhatsApp d'étudiants à Abidjan qui partagent des textes littéraires.
  Exemple 2 : Les élèves qui n'ont pas les moyens d'acheter des livres papier.

  Étape 4 · Identifier la conclusion :
  → L'auteur conclut que prétendre que les réseaux sociaux nuisent à la lecture, c'est ignorer la réalité.

┌─ À RETENIR ──────────────────────────────────────────┐
  Un texte argumentatif qui réfute suit toujours ce
  schéma :
  1. Présenter la thèse adverse
  2. La combattre avec des contre-arguments
  3. Illustrer chaque contre-argument par un exemple
  4. Conclure en rejetant définitivement la thèse adverse
└──────────────────────────────────────────────────────┘

  LES OUTILS LINGUISTIQUES DE LA RÉFUTATION

  Pour introduire la thèse adverse :
  · "Certains affirment que…"
  · "On prétend souvent que…"
  · "Il est courant de penser que…"
  · "D'aucuns soutiennent que…"

  Pour marquer le rejet / la réfutation :
  · "Cependant…" / "Pourtant…" / "Or…"
  · "Cette affirmation est fausse car…"
  · "Rien ne prouve que…"
  · "Au contraire…" / "En réalité…"
  · "Il convient de nuancer cette idée…"
  · "Cette thèse mérite d'être réfutée…"

  Pour introduire un contre-argument :
  · "En effet…" / "Car…" / "Parce que…"
  · "La réalité montre que…"
  · "Il est prouvé que…"

  Pour illustrer :
  · "Ainsi…" / "Par exemple…" / "C'est le cas de…"
  · "À titre d'exemple…"

  Pour conclure la réfutation :
  · "Ainsi, il apparaît clairement que…"
  · "Il est donc inexact de prétendre que…"
  · "En définitive, cette thèse ne résiste pas à l'analyse…"

⚠ ATTENTION
  Réfuter ne signifie pas insulter ou attaquer la personne.
  On attaque les IDÉES, pas les personnes.
  Au BEPC, un texte qui réfute sans contre-arguments
  ni exemples n'obtient aucun point de développement.

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Le schéma argumentatif de la réfutation]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Boîte centrale en haut (couleur rouge) : "THÈSE ADVERSE — à réfuter"
    Exemple : "Les jeunes ne lisent plus à cause des réseaux sociaux."
  · Flèche vers le bas portant la mention "RÉFUTATION"
  · Deux boîtes en dessous (couleur verte) côte à côte :
    Boîte gauche : "Contre-argument 1 + Exemple 1"
    Boîte droite : "Contre-argument 2 + Exemple 2"
  · Flèche vers le bas des deux boîtes convergeant vers une boîte finale (couleur bleue) :
    "CONCLUSION — La thèse adverse est rejetée"
  · Légende :
    Rouge = ce qu'on combat
    Vert = nos armes (contre-arguments + exemples)
    Bleu = notre victoire (conclusion)
Note générateur : Canva, draw.io ou GeoGebra (mode texte)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Carte mentale : La boîte à outils de la réfutation]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Nœud central (couleur orange) : "RÉFUTER"
  · 4 branches principales :
    Branche 1 (rouge) : "Thèse adverse" → sous-nœuds : "Identifier" / "Citer" / "Annoncer qu'on la rejette"
    Branche 2 (verte) : "Contre-arguments" → sous-nœuds : "Logiques" / "Factuels" / "Introduits par : Or / Cependant / En réalité"
    Branche 3 (bleue) : "Exemples" → sous-nœuds : "Concrets" / "Ivoiriens si possible" / "Introduits par : Par exemple / Ainsi"
    Branche 4 (violette) : "Conclusion" → sous-nœuds : "Rejeter définitivement" / "Affirmer sa propre position"
Note générateur : MindMeister, Canva, draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Réfuter
  Réfuter une thèse, c'est démontrer, à l'aide de contre-arguments et d'exemples, qu'une affirmation est fausse, insuffisante ou inacceptable.

◆ DÉFINITION DPFC — Thèse adverse
  La thèse adverse est l'opinion ou l'affirmation que l'auteur du texte argumentatif choisit de combattre et de rejeter.

◆ DÉFINITION DPFC — Contre-argument
  Le contre-argument est une raison logique qui s'oppose à la thèse adverse et qui démontre son erreur ou sa faiblesse.

◆ DÉFINITION DPFC — Outils de la langue (réfutation)
  Les outils de la langue propres à la réfutation sont les connecteurs d'opposition et de concession : cependant, pourtant, or, en réalité, au contraire, néanmoins, toutefois.

◆ DÉFINITION DPFC — Schéma argumentatif de réfutation
  Le schéma argumentatif de réfutation est l'organisation logique d'un texte qui combat une thèse : Présentation de la thèse adverse → Contre-arguments + Exemples → Conclusion de rejet.

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier)

  I. QU'EST-CE QUE RÉFUTER ?
  Réfuter une thèse, c'est démontrer qu'elle est fausse ou inacceptable
  grâce à des contre-arguments et des exemples.
  On ne soutient pas la thèse : on la combat.

  II. LE SCHÉMA ARGUMENTATIF DE LA RÉFUTATION
  Introduction  → Présenter la thèse adverse + annoncer qu'on va la réfuter
  Développement → Contre-argument 1 + Exemple 1
                  Contre-argument 2 + Exemple 2
  Conclusion    → Rejeter définitivement la thèse adverse

  III. LES OUTILS DE LA RÉFUTATION
  · Introduire la thèse adverse : "Certains affirment que… / On prétend souvent que…"
  · Marquer le rejet : "Cependant / Pourtant / Or / En réalité / Au contraire"
  · Introduire un contre-argument : "En effet / Car / La réalité montre que…"
  · Illustrer : "Par exemple / Ainsi / C'est le cas de…"
  · Conclure : "Il est donc inexact de prétendre que… / En définitive…"

  IV. RÈGLES À RETENIR POUR LE BEPC
  Règle 1 · Toujours citer la thèse adverse avant de la réfuter.
  Règle 2 · Chaque contre-argument doit être accompagné d'un exemple.
  Règle 3 · La conclusion doit rejeter clairement la thèse adverse.
  Règle 4 · Utiliser les connecteurs d'opposition (cependant, or, pourtant…).
  Règle 5 · Le texte doit comporter au minimum 2 contre-arguments.

~ ANALOGIE CI — Le grin d'Abidjan
  Dans un grin (groupe d'amis qui débattent autour d'un thé à Abidjan), quand quelqu'un dit "Didier Drogba est le meilleur footballeur africain de tous les temps", un ami peut le réfuter ainsi : "Cette affirmation est discutable. En réalité, Roger Milla a joué et marqué à 42 ans en Coupe du Monde 1994, ce qu'aucun autre Africain n'a fait. De plus, il a remporté le titre de meilleur footballeur africain à deux reprises."
  C'est exactement le schéma de la réfutation : thèse adverse + contre-arguments + exemples concrets.

~ ANALOGIE CI — Le tribunal d'Abidjan
  Imaginez un avocat au Tribunal de Première Instance d'Abidjan. Le procureur affirme : "L'accusé est coupable car il était sur les lieux du crime." L'avocat réfute : "Cette affirmation est inexacte. En réalité, mon client était à Bouaké ce jour-là, comme le prouvent son billet de bus et le témoignage de trois personnes." La réfutation, c'est le rôle de l'avocat de la défense.

────────────────────────────────────────────────────────

Phase 5b · Synthèse

  5 points essentiels à retenir :
  Point 1 · Réfuter = combattre une thèse adverse en prouvant qu'elle est fausse.
  Point 2 · Le schéma : Thèse adverse → Contre-arguments + Exemples → Conclusion de rejet.
  Point 3 · Les connecteurs d'opposition sont les armes de la réfutation : or, cependant, pourtant, en réalité.
  Point 4 · Chaque contre-argument doit être accompagné d'un exemple concret.
  Point 5 · La conclusion doit rejeter clairement et définitivement la thèse adverse.

  5 mots clés + définition courte :
  Mot 1 · Réfuter → démontrer qu'une thèse est fausse.
  Mot 2 · Thèse adverse → l'affirmation qu'on combat.
  Mot 3 · Contre-argument → raison logique contre la thèse adverse.
  Mot 4 · Connecteur d'opposition → mot qui introduit le rejet (or, cependant…).
  Mot 5 · Schéma argumentatif → organisation logique d'un texte qui réfute.

  3 erreurs fréquentes + correction :
  Erreur 1 · Oublier de citer la thèse adverse avant de la réfuter.
  Correction → Toujours commencer par "Certains affirment que…" ou équivalent.

  Erreur 2 · Donner des contre-arguments sans exemples.
  Correction → Après chaque contre-argument, ajouter "Par exemple…" ou "Ainsi…"

  Erreur 3 · Confondre réfuter et étayer (soutenir au lieu de combattre).
  Correction → Vérifier que le texte s'oppose à la thèse de départ, ne la soutient pas.

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Analyser un texte qui réfute · Notion ciblée : Identification du schéma argumentatif de réfutation

  Énoncé :
  Lisez ce texte et répondez aux questions.

  "Il est fréquent d'entendre que les femmes sont moins douées que les hommes pour les sciences. Cette idée reçue ne résiste pas à l'analyse. En réalité, de nombreuses femmes ont contribué aux plus grandes découvertes scientifiques. Marie Curie, par exemple, a reçu deux prix Nobel, en physique et en chimie. En Côte d'Ivoire, la professeure Aïcha Kone dirige un laboratoire de recherche à l'université FHB de Cocody et a publié plusieurs travaux reconnus internationalement. Affirmer que les femmes sont moins douées pour les sciences, c'est donc nier des faits historiques et actuels."

  Questions :
  1. Quelle est la thèse adverse (la thèse que l'auteur réfute) ?
  2. Identifiez deux contre-arguments.
  3. Identifiez deux exemples.
  4. Relevez deux connecteurs d'opposition ou de réfutation.
  5. Quelle est la conclusion ?

  MÉTHODE
  Étape 1 · Lire le texte une première fois en entier pour le comprendre globalement.
  Étape 2 · Chercher la phrase qui commence par "Il est fréquent…" ou "On dit souvent…" → c'est la thèse adverse.
  Étape 3 · Chercher les phrases introduites par "En réalité", "En effet", "Car" → ce sont les contre-arguments.
  Étape 4 · Chercher les phrases introduites par "Par exemple", "Ainsi", les noms propres → ce sont les exemples.
  Étape 5 · Chercher les mots qui marquent le rejet → connecteurs d'opposition.
  Étape 6 · Lire la dernière phrase → c'est la conclusion.
  ✔ Vérification : Chaque contre-argument doit être associé à un exemple.
  ✔ Conclusion : "Il est donc inexact de prétendre que les femmes sont moins douées pour les sciences."
  ✔ Ce que cet exercice mobilise : identification de la thèse adverse · identification des contre-arguments · identification des exemples · repérage des connecteurs d'opposition · lecture analytique d'un texte argumentatif.

  CORRIGÉ DE L'EXERCICE GUIDÉ 1 :
  1. Thèse adverse : "Les femmes sont moins douées que les hommes pour les sciences."
  2. Contre-argument 1 : De nombreuses femmes ont contribué aux plus grandes découvertes scientifiques.
     Contre-argument 2 : En Côte d'Ivoire, des femmes dirigent des laboratoires et publient des travaux reconnus.
  3. Exemple 1 : Marie Curie — deux prix Nobel (physique et chimie).
     Exemple 2 : La professeure Aïcha Kone — laboratoire à l'université FHB de Cocody.
  4. Connecteurs relevés : "En réalité" / "ne résiste pas à l'analyse" / "c'est donc nier…"
  5. Conclusion : "Affirmer que les femmes sont moins douées pour les sciences, c'est nier des faits historiques et actuels."

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Rédiger un texte qui réfute · Notion ciblée : Production écrite — construction du schéma argumentatif de réfutation

  Énoncé :
  Réfutez la thèse suivante en 15 à 20 lignes :
  "Les élèves qui regardent la télévision le soir travaillent moins bien à l'école."
  Vous utiliserez au moins 2 contre-arguments, 2 exemples et les connecteurs appropriés.

  MÉTHODE
  Étape 1 · Rédiger l'introduction : présenter la thèse adverse et annoncer qu'on va la réfuter.
  Étape 2 · Trouver le contre-argument 1 : par exemple, certaines émissions télévisées sont éducatives.
  Étape 3 · Illustrer le contre-argument 1 par un exemple ivoirien ou concret.
  Étape 4 · Trouver le contre-argument 2 : par exemple, regarder la télévision peut être une détente qui améliore la concentration.
  Étape 5 · Illustrer le contre-argument 2 par un exemple.
  Étape 6 · Rédiger la conclusion : rejeter définitivement la thèse adverse.
  ✔ Vérification : Ai-je cité la thèse adverse ? Ai-je 2 contre-arguments avec exemples ? Ai-je utilisé des connecteurs d'opposition ?
  ✔ Conclusion attendue : La thèse adverse est clairement rejetée dans la dernière phrase.
  ✔ Ce que cet exercice mobilise : construction du schéma argumentatif de réfutation · utilisation des connecteurs d'opposition · rédaction d'une introduction et d'une conclusion · production d'un texte argumentatif complet · mobilisation d'exemples concrets.

  MODÈLE DE RÉDACTION COMPLET :

  "Il est courant d'affirmer que les élèves qui regardent la télévision le soir travaillent moins bien à l'école. Cette affirmation mérite d'être réfutée.

  En réalité, toutes les émissions télévisées ne nuisent pas aux études. Certaines chaînes, comme TV5 Monde, diffusent des documentaires éducatifs sur la science, l'histoire ou la géographie. Ainsi, un élève de 3ème qui regarde un documentaire sur la Seconde Guerre mondiale enrichit directement ses connaissances en Histoire-Géographie.

  De plus, regarder la télévision en soirée peut constituer une détente nécessaire après une longue journée de travail scolaire. Or, il est prouvé que le cerveau a besoin de pauses pour mieux mémoriser. À Abidjan, de nombreux élèves des classes de Terminale témoignent que regarder une émission le soir leur permet de décompresser et de se concentrer plus efficacement le lendemain.

  Il est donc inexact de prétendre que la télévision nuit systématiquement aux résultats scolaires. C'est l'usage qu'on en fait qui détermine son impact, non la télévision elle-même."

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Identifier la thèse adverse · Notions travaillées : Thèse adverse, lecture analytique

  Lisez les trois extraits suivants et, pour chacun, identifiez la thèse adverse que l'auteur réfute.

  Extrait A :
  "Beaucoup pensent que le téléphone portable est nuisible aux élèves. Cependant, force est de constater que de nombreux élèves ivoiriens utilisent leur téléphone pour accéder à des ressources pédagogiques en ligne, consulter des dictionnaires ou rejoindre des groupes d'étude."

  Extrait B :
  "On entend souvent dire que les marchés traditionnels africains sont moins efficaces que les supermarchés modernes. Pourtant, les marchés comme celui d'Adjamé à Abidjan offrent des produits frais, des prix accessibles et créent des milliers d'emplois locaux."

  Extrait C :
  "Certains affirment que le sport prend du temps aux élèves et nuit à leurs études. Or, de nombreuses études montrent que la pratique sportive régulière améliore la concentration et la mémoire."

  GUIDE
  Étape 1 · Lire chaque extrait en cherchant la phrase qui présente l'opinion que l'auteur va combattre.
  Étape 2 · Repérer les expressions introductives : "Beaucoup pensent que…", "On entend souvent dire que…", "Certains affirment que…"
  Étape 3 · Recopier la thèse adverse de chaque extrait en une phrase complète.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Connecteurs d'opposition · Notions travaillées : Outils de la langue, connecteurs de réfutation

  Complétez les phrases suivantes avec le connecteur d'opposition approprié parmi : cependant / or / pourtant / en réalité / au contraire / néanmoins.

  1. On prétend que les jeunes Ivoiriens ne s'intéressent pas à la politique. __________, les élections présidentielles de 2020 ont montré une forte mobilisation des jeunes voters.
  2. Certains pensent que les filles réussissent moins bien en Mathématiques. __________, les résultats du BEPC montrent que les filles obtiennent souvent de meilleures notes que les garçons dans cette matière.
  3. Il est courant de penser que manger dans la rue est dangereux pour la santé. __________, les maquis et les vendeurs de rue d'Abidjan respectent souvent des normes d'hygiène strictes.
  4. On dit que les transports en commun ivoiriens sont inefficaces. __________, les gbaka et les woro-woro transportent chaque jour des millions de personnes à travers Abidjan.

  GUIDE
  Étape 1 · Lire chaque phrase en entier pour comprendre le sens général.
  Étape 2 · Identifier si la deuxième partie de la phrase contredit ou nuance la première.
  Étape 3 · Choisir le connecteur le plus précis selon le degré d'opposition.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Construire un contre-argument · Notions travaillées : Contre-argument, exemple, schéma argumentatif

  Pour chacune des thèses adverses suivantes, rédigez un contre-argument accompagné d'un exemple concret ivoirien ou africain.

  Thèse adverse 1 : "Le cacao ivoirien ne profite pas aux populations locales."
  Thèse adverse 2 : "Les langues africaines ne peuvent pas être utilisées dans l'enseignement."
  Thèse adverse 3 : "Internet rend les élèves paresseux."

  GUIDE
  Étape 1 · Lire la thèse adverse et identifier son erreur ou sa faiblesse.
  Étape 2 · Formuler un contre-argument qui commence par "En réalité…" ou "Or…"
  Étape 3 · Ajouter un exemple introduit par "Par exemple…" ou "Ainsi…" ou "C'est le cas de…"
  Étape 4 · Vérifier que l'exemple est concret, précis et lié au contexte ivoirien ou africain si possible.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Remettre dans l'ordre · Notions travaillées : Schéma argumentatif de réfutation, organisation du texte

  Les phrases suivantes forment un texte argumentatif qui réfute, mais elles ont été mélangées. Remettez-les dans le bon ordre (Introduction → Contre-argument 1 + Exemple 1 → Contre-argument 2 + Exemple 2 → Conclusion).

  Phrase A : "Il est donc inexact de dire que le football est une perte de temps pour les jeunes Ivoiriens."
  Phrase B : "Certains affirment que pratiquer le football prend trop de temps aux jeunes et les éloigne de leurs études."
  Phrase C : "Par exemple, le joueur ivoirien Franck Kessié a commencé à jouer au football dès l'enfance tout en continuant ses études et est devenu un professionnel reconnu en Europe."
  Phrase D : "De plus, le football développe des valeurs essentielles comme la discipline, le travail d'équipe et la persévérance, qui sont également utiles dans les études."
  Phrase E : "En réalité, le football peut être une école de vie qui enseigne la rigueur et l'organisation du temps."
  Phrase F : "Ainsi, de nombreux clubs de football scolaires à Abidjan permettent aux élèves de concilier sport et excellence académique."

  GUIDE
  Étape 1 · Identifier la phrase qui présente la thèse adverse → c'est l'introduction.
  Étape 2 · Identifier les phrases qui commencent par "En réalité" et "De plus" → ce sont les contre-arguments.
  Étape 3 · Identifier les phrases qui commencent par "Par exemple" et "Ainsi" → ce sont les exemples.
  Étape 4 · Identifier la phrase de conclusion → elle rejette définitivement la thèse adverse.
  Étape 5 · Recopier les phrases dans le bon ordre.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Rédaction complète · Notions travaillées : Production intégrale d'un texte argumentatif de réfutation

  Réfutez la thèse suivante en un texte de 15 à 20 lignes :
  "Les femmes ne devraient pas travailler et rester à la maison pour s'occuper de leur famille."

  Contraintes obligatoires :
  · Citer la thèse adverse dans l'introduction.
  · Utiliser au minimum 2 contre-arguments.
  · Illustrer chaque contre-argument par un exemple concret (ivoirien ou africain de préférence).
  · Utiliser au moins 3 connecteurs d'opposition parmi : cependant / or / pourtant / en réalité / au contraire / néanmoins.
  · Conclure en rejetant clairement la thèse adverse.

  GUIDE
  Étape 1 · Rédiger l'introduction : présenter la thèse adverse et annoncer la réfutation.
  Étape 2 · Rédiger le contre-argument 1 et son exemple.
  Étape 3 · Rédiger le contre-argument 2 et son exemple.
  Étape 4 · Rédiger la conclusion : rejeter la thèse adverse.
  Étape 5 · Vérifier les connecteurs, l'orthographe, la ponctuation.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — Réfutation : La place de la culture traditionnelle ivoirienne
  Durée conseillée : 45 minutes

  Énoncé :
  Réfutez la thèse suivante en un texte argumentatif complet de 20 à 25 lignes :
  "Les traditions culturelles ivoiriennes sont un frein au développement économique du pays."

  Contraintes obligatoires :
  · Citer la thèse adverse dans l'introduction.
  · Développer au minimum 2 contre-arguments solides.
  · Illustrer chaque contre-argument par un exemple précis tiré de la réalité ivoirienne.
  · Utiliser au moins 4 connecteurs d'opposition ou de réfutation.
  · Rédiger une conclusion qui rejette clairement et définitivement la thèse adverse.
  · Respecter le plan : Introduction / Développement (CA1 + Ex1 / CA2 + Ex2) / Conclusion.
  · Soigner l'orthographe, la ponctuation et la cohérence du texte.

  GUIDE (sans corrigé)
  Étape 1 · Réfléchir : quelles traditions ivoiriennes pourraient au contraire favoriser le développement ? (tourisme culturel, artisanat, cohésion sociale…)
  Étape 2 · Rédiger l'introduction : "Certains affirment que… Cependant, cette thèse mérite d'être réfutée…"
  Étape 3 · Contre-argument 1 : par exemple, les traditions ivoiriennes attirent le tourisme et génèrent des revenus.
  Étape 4 · Exemple 1 : les masques Goli des Baoulé, le Festival des Masques de Man, la Fête des Ignames…
  Étape 5 · Contre-argument 2 : les traditions renforcent la cohésion sociale, condition du développement économique.
  Étape 6 · Exemple 2 : les tontines traditionnelles ivoiriennes financent des projets économiques locaux.
  Étape 7 · Conclure : "Il est donc inexact de prétendre que les traditions ivoiriennes freinent le développement…"
  Étape 8 · Relire et corriger.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON

  Introduction :
  "Certains affirment que les traditions culturelles ivoiriennes constituent un obstacle au développement économique de la Côte d'Ivoire. Cette thèse, bien que répandue, ne résiste pas à une analyse sérieuse. En réalité, les traditions ivoiriennes sont une richesse qui peut, au contraire, contribuer au développement du pays."

  Contre-argument 1 :
  Les traditions culturelles ivoiriennes constituent un véritable moteur pour le secteur touristique et artisanal, source de revenus considérables.
  Exemple 1 :
  Ainsi, le Festival des Masques de Man attire chaque année des milliers de touristes nationaux et internationaux. Les masques Goli des Baoulé et les pagnes tissés de Korhogo sont vendus sur les marchés africains et européens, générant des revenus pour des milliers d'artisans ivoiriens.

  Contre-argument 2 :
  De plus, les traditions ivoiriennes renforcent la cohésion sociale et la solidarité communautaire, conditions indispensables à tout développement économique durable.
  Exemple 2 :
  Or, les tontines traditionnelles, pratiquées dans tout le pays, permettent à des groupes de femmes et d'hommes de cotiser ensemble pour financer des projets économiques : commerce, agriculture, construction. Ces mécanismes de solidarité traditionnelle suppléent parfois au manque d'accès au crédit bancaire.

  Conclusion :
  "Il est donc inexact de prétendre que les traditions culturelles ivoiriennes freinent le développement économique. Au contraire, bien valorisées, elles constituent un patrimoine vivant et une ressource économique réelle. En définitive, c'est l'absence de politique culturelle cohérente, et non les traditions elles-mêmes, qui peut constituer un frein au développement."

────────────────────────────────────────────────────────

✔ CORRIGÉ — Exercice 1

  Extrait A — Thèse adverse : "Le téléphone portable est nuisible aux élèves."
  Extrait B — Thèse adverse : "Les marchés traditionnels africains sont moins efficaces que les supermarchés modernes."
  Extrait C — Thèse adverse : "Le sport prend du temps aux élèves et nuit à leurs études."

────────────────────────────────────────────────────────

✔ CORRIGÉ — Exercice 2

  1. On prétend que les jeunes Ivoiriens ne s'intéressent pas à la politique. CEPENDANT / OR, les élections présidentielles de 2020 ont montré une forte mobilisation des jeunes.
  2. Certains pensent que les filles réussissent moins bien en Mathématiques. POURTANT / OR, les résultats du BEPC montrent que les filles obtiennent souvent de meilleures notes.
  3. Il est courant de penser que manger dans la rue est dangereux. EN RÉALITÉ / CEPENDANT, les maquis et vendeurs de rue d'Abidjan respectent souvent des normes d'hygiène strictes.
  4. On dit que les transports en commun ivoiriens sont inefficaces. OR / CEPENDANT, les gbaka et les woro-woro transportent chaque jour des millions de personnes.

  Note : plusieurs réponses sont acceptables pour chaque phrase, à condition que le connecteur marque bien une opposition.

────────────────────────────────────────────────────────

✔ CORRIGÉ — Exercice 3

  Thèse adverse 1 : "Le cacao ivoirien ne profite pas aux populations locales."
  Contre-argument : En réalité, la filière cacao emploie directement plus de 6 millions de personnes en Côte d'Ivoire, dont une majorité de petits producteurs ruraux qui tirent leur revenu principal de cette culture.
  Exemple : Par exemple, dans la région de Soubré, des coopératives de producteurs comme la COOPAC redistribuent une partie de leurs bénéfices à leurs membres sous forme de primes et d'équipements agricoles.

  Thèse adverse 2 : "Les langues africaines ne peuvent pas être utilisées dans l'enseignement."
  Contre-argument : Or, plusieurs pays africains ont déjà intégré avec succès des langues locales dans leur système éducatif, améliorant ainsi la compréhension des élèves.
  Exemple : Ainsi, au Mali, le bambara est utilisé comme langue d'enseignement dans les premières années du primaire, ce qui a amélioré les taux de réussite des élèves ruraux.

  Thèse adverse 3 : "Internet rend les élèves paresseux."
  Contre-argument : En réalité, Internet offre un accès à des ressources pédagogiques immenses qui stimulent la curiosité intellectuelle et l'autonomie des élèves.
  Exemple : Par exemple, de nombreux élèves ivoiriens utilisent des plateformes comme YouTube ou des applications éducatives pour réviser leurs cours de Mathématiques ou d'Anglais, souvent avec de meilleurs résultats que par les méthodes traditionnelles seules.

────────────────────────────────────────────────────────

✔ CORRIGÉ — Exercice 4

  Ordre correct :
  Phrase B → Phrase E → Phrase C → Phrase D → Phrase F → Phrase A

  Reconstruction du texte dans l'ordre :
  "Certains affirment que pratiquer le football prend trop de temps aux jeunes et les éloigne de leurs études. [B]
  En réalité, le football peut être une école de vie qui enseigne la rigueur et l'organisation du temps. [E]
  Par exemple, le joueur ivoirien Franck Kessié a commencé à jouer au football dès l'enfance tout en continuant ses études et est devenu un professionnel reconnu en Europe. [C]
  De plus, le football développe des valeurs essentielles comme la discipline, le travail d'équipe et la persévérance, qui sont également utiles dans les études. [D]
  Ainsi, de nombreux clubs de football scolaires à Abidjan permettent aux élèves de concilier sport et excellence académique. [F]
  Il est donc inexact de dire que le football est une perte de temps pour les jeunes Ivoiriens." [A]

────────────────────────────────────────────────────────

✔ CORRIGÉ — Exercice 5

  Modèle de corrigé :

  "Certains affirment que les femmes ne devraient pas travailler et rester à la maison pour s'occuper de leur famille. Cette thèse, pourtant encore répandue, ne résiste pas à l'analyse des réalités contemporaines.

  En réalité, les femmes constituent une force économique indispensable pour leur famille et pour la société tout entière. Or, en Côte d'Ivoire, de nombreuses femmes sont cheffes d'entreprise, enseignantes, médecins ou fonctionnaires et assurent à la fois leur rôle professionnel et familial. Par exemple, la Présidente du Conseil Économique, Social, Environnemental et Culturel de Côte d'Ivoire est une femme, preuve que les femmes ivoiriennes excellent dans les sphères publiques les plus exigeantes.

  De plus, priver les femmes du droit de travailler, c'est appauvrir les familles et ralentir le développement économique du pays. Au contraire, les études montrent que chaque femme qui accède au marché du travail améliore le niveau de vie de son foyer et contribue à l'éducation de ses enfants. Ainsi, les vendeuses des marchés d'Adjamé et de Treichville financent la scolarité de leurs enfants grâce à leurs revenus.

  Il est donc inexact de prétendre que les femmes devraient rester confinées au foyer. Néanmoins, concilier vie professionnelle et familiale demande une organisation que la société doit soutenir par des politiques adaptées. En définitive, le travail des femmes est une richesse pour la famille, pour la Côte d'Ivoire et pour l'Afrique tout entière."

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "L'éducation est l'arme la plus puissante que vous puissiez utiliser pour changer le monde."
  (Nelson Mandela)

  Prof. Adjoua KONAN · Français · 3ème · DPFC/MENET-FP CI · 2025-2026
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
