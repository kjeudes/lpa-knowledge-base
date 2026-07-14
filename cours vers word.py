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

COURS = """---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 15 · Classe de 3ème

RURAL EXODUS
L'exode rural

Compétence C5 · Écoute simple — Ville ou Village / City or Village

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 15 — RURAL EXODUS
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C5 · Écoute simple — Ville ou Village
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 14 : Village life)
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

CAPSULE DE REPRISE — Leçon 14 : Village life (La vie au village)

5 points clés à se rappeler :
  1. Un village est une petite communauté rurale éloignée des grandes villes.
  2. La vie au village est simple : agriculture, élevage, pêche, artisanat.
  3. Les infrastructures y sont limitées : peu d'écoles, d'hôpitaux, d'électricité.
  4. La solidarité et les valeurs traditionnelles sont très fortes au village.
  5. Le vocabulaire clé : farmer (agriculteur), crops (récoltes), well (puits), community (communauté), peaceful (paisible).

3 questions flash :
  Q1 : How do you say "agriculteur" in English?
  R1 : Farmer.

  Q2 : Give one advantage of village life. (Donne un avantage de la vie au village.)
  R2 : It is peaceful / There is a strong community spirit.

  Q3 : Complete the sentence: "In a village, people often grow their own ___."
  R3 : Food / crops.

Lien avec la leçon du jour :
  Nous avons vu pourquoi la vie au village est belle mais difficile.
  Aujourd'hui, nous allons comprendre pourquoi de nombreuses personnes QUITTENT le village pour aller en ville — c'est ce qu'on appelle l'exode rural (rural exodus).

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

Amadou a grandi à Bouaké. Son père est agriculteur. Mais depuis quelques années, les récoltes sont mauvaises et l'école du village manque de professeurs. À 17 ans, Amadou décide de partir à Abidjan pour chercher un emploi et une meilleure vie. Il rejoint des millions d'Ivoiriens qui font le même choix chaque année.

Ce phénomène — quitter le village pour s'installer en ville — s'appelle l'exode rural.
En Côte d'Ivoire, Abidjan concentre aujourd'hui plus de 20% de la population nationale, en grande partie à cause de l'exode rural.

Lien DPFC : Ce thème est directement lié aux problèmes de développement économique et d'urbanisation abordés en Histoire-Géographie.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — Causes de l'exode rural
  Q : Why do people leave their villages to go to the city?
      (Pourquoi les gens quittent-ils leur village pour aller en ville ?)
  Guide :
    · Pense aux conditions de vie au village (travail, école, santé).
    · Pense à ce que la ville offre que le village n'a pas.
    · Utilise les mots : jobs (emplois), schools (écoles), hospitals (hôpitaux).
  Corrigé :
    People leave their villages because life is hard there. There are few jobs, schools are far away, and there are not enough hospitals. They hope to find better opportunities in the city, such as work, education and modern services.
    (Les gens quittent leurs villages parce que la vie y est difficile. Il y a peu d'emplois, les écoles sont loin et les hôpitaux insuffisants. Ils espèrent trouver de meilleures opportunités en ville : travail, éducation et services modernes.)

▶ HARKNESS 2 — Conséquences de l'exode rural
  Q : What are the consequences of rural exodus for the city and the village?
      (Quelles sont les conséquences de l'exode rural pour la ville et le village ?)
  Guide :
    · Pour la ville : pense à la surpopulation, aux bidonvilles, au chômage.
    · Pour le village : pense à l'abandon des champs, au vieillissement de la population.
    · Deux faces d'une même médaille : ville et village souffrent tous les deux.
  Corrigé :
    For the city, rural exodus causes overpopulation, the growth of slums and unemployment. For the village, it leads to the abandonment of farmland and an ageing population. Both city and village face serious problems because of this phenomenon.
    (Pour la ville, l'exode rural provoque la surpopulation, la croissance des bidonvilles et le chômage. Pour le village, il entraîne l'abandon des terres agricoles et le vieillissement de la population. La ville comme le village souffrent de ce phénomène.)

▶ HARKNESS 3 — Solutions possibles
  Q : What can be done to stop or reduce rural exodus?
      (Que peut-on faire pour stopper ou réduire l'exode rural ?)
  Guide :
    · Pense à ce qu'il faudrait améliorer au village pour que les gens restent.
    · Pense au rôle du gouvernement, des ONG, des communautés locales.
    · Utilise : invest (investir), build (construire), create (créer).
  Corrigé :
    To reduce rural exodus, governments must invest in villages by building schools, hospitals and roads. They must also create job opportunities in rural areas, for example by supporting farmers and developing agribusiness. If life improves in the village, people will have a reason to stay.
    (Pour réduire l'exode rural, les gouvernements doivent investir dans les villages en construisant des écoles, des hôpitaux et des routes. Ils doivent également créer des emplois en zone rurale, par exemple en soutenant les agriculteurs et en développant l'agro-industrie. Si la vie s'améliore au village, les gens auront une raison d'y rester.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre de la leçon : Rural Exodus (L'exode rural)
Compétence : C5 — Écoute simple · Ville ou Village

Objectifs de la leçon :
  À la fin de cette leçon, tu seras capable de :
  · Comprendre et utiliser le vocabulaire de l'exode rural en anglais.
  · Expliquer les causes et les conséquences de l'exode rural.
  · Utiliser le Conditional Type 1 (If + Present Simple → will + verbe) pour parler de solutions.
  · Produire un court texte ou dialogue sur ce thème.

Prérequis nécessaires :
  · Vocabulaire de la ville (Leçon 13) et du village (Leçon 14).
  · Le Simple Past (actions passées).
  · Le Present Simple (faits généraux).
  · Les comparatifs (bigger than, more than...).

────────────────────────────────────────────────────────

TABLEAU DES STRUCTURES FONDAMENTALES — Leçon 15

| Expression / Notion            | Valeur, emploi                              | Exemple anglais                                      | Traduction française                                      |
|-------------------------------|---------------------------------------------|------------------------------------------------------|-----------------------------------------------------------|
| Rural exodus                  | Nom du phénomène                            | Rural exodus is a serious problem.                   | L'exode rural est un problème sérieux.                    |
| To leave / left               | Quitter (irrégulier)                        | Many people left their village.                      | Beaucoup de gens ont quitté leur village.                 |
| To move to                    | S'installer dans / déménager vers           | They moved to Abidjan.                               | Ils se sont installés à Abidjan.                          |
| To look for                   | Chercher                                    | He is looking for a job.                             | Il cherche un emploi.                                     |
| To hope to + verbe            | Espérer faire quelque chose                 | She hopes to find work.                              | Elle espère trouver du travail.                           |
| Because / because of          | Parce que / à cause de                      | They left because of poverty.                        | Ils sont partis à cause de la pauvreté.                   |
| Conditional Type 1            | Si... alors... (futur possible)             | If the government invests, villages will grow.       | Si le gouvernement investit, les villages se développeront.|
| Push factors                  | Facteurs qui poussent à partir              | Poverty is a push factor.                            | La pauvreté est un facteur qui pousse à partir.           |
| Pull factors                  | Facteurs qui attirent vers la ville         | Jobs are pull factors.                               | Les emplois sont des facteurs d'attraction.               |
| Overpopulation                | Surpopulation                               | Cities face overpopulation.                          | Les villes font face à la surpopulation.                  |
| Slum / slums                  | Bidonville(s)                               | Many migrants live in slums.                         | Beaucoup de migrants vivent dans des bidonvilles.         |
| To improve                    | Améliorer / s'améliorer                     | Life must improve in rural areas.                    | La vie doit s'améliorer dans les zones rurales.           |

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

A. QU'EST-CE QUE L'EXODE RURAL ? / WHAT IS RURAL EXODUS?

Rural exodus is the movement of people from rural areas (villages) to urban areas (cities).
(L'exode rural est le déplacement de personnes des zones rurales — villages — vers les zones urbaines — villes.)

Ce phénomène touche toute l'Afrique et particulièrement la Côte d'Ivoire.
This phenomenon affects all of Africa and particularly Côte d'Ivoire.

┌─ À RETENIR ──────────────────────────────────────────┐
  Rural = qui concerne la campagne / le village
  Urban = qui concerne la ville
  Exodus = départ massif d'un groupe de personnes
  Rural exodus = départ massif des habitants des villages vers les villes
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

B. LES CAUSES / THE CAUSES

Il existe deux types de facteurs qui expliquent l'exode rural :

1. PUSH FACTORS — Les facteurs qui poussent à partir du village :
  · Poverty (la pauvreté) : farmers do not earn enough money.
    (Les agriculteurs ne gagnent pas assez d'argent.)
  · Lack of schools (manque d'écoles) : children cannot get a good education.
    (Les enfants ne peuvent pas avoir une bonne éducation.)
  · Lack of hospitals (manque d'hôpitaux) : people cannot get proper healthcare.
    (Les gens ne peuvent pas avoir accès à de bons soins de santé.)
  · Drought and bad harvests (sécheresse et mauvaises récoltes) : farming becomes impossible.
    (L'agriculture devient impossible.)
  · Lack of electricity and water (manque d'électricité et d'eau courante).

2. PULL FACTORS — Les facteurs qui attirent vers la ville :
  · Jobs and employment (emplois) : factories, companies, shops need workers.
    (Les usines, entreprises et boutiques ont besoin de travailleurs.)
  · Better schools and universities (meilleures écoles et universités).
  · Better hospitals and healthcare (meilleurs hôpitaux et soins de santé).
  · Modern life (vie moderne) : electricity, internet, transport.
    (Électricité, internet, transports.)
  · Higher salaries (salaires plus élevés).

⚠ ATTENTION
  Push factors = raisons de PARTIR (négatif pour le village).
  Pull factors = raisons d'ALLER en ville (positif pour la ville).
  Ces deux concepts sont souvent demandés au BEPC !
  Ne les confonds pas.

────────────────────────────────────────────────────────

C. LES CONSÉQUENCES / THE CONSEQUENCES

Pour la ville / For the city :
  · Overpopulation (surpopulation) : too many people in a small space.
  · Growth of slums (croissance des bidonvilles) : migrants cannot afford proper housing.
    (Les migrants ne peuvent pas se payer un logement décent.)
  · Unemployment (chômage) : there are not enough jobs for everyone.
    (Il n'y a pas assez d'emplois pour tout le monde.)
  · Pollution and traffic jams (pollution et embouteillages).

Pour le village / For the village :
  · Abandonment of farmland (abandon des terres agricoles) : no one left to farm.
    (Il ne reste personne pour cultiver.)
  · Ageing population (vieillissement de la population) : only old people stay.
    (Seules les personnes âgées restent.)
  · Loss of traditions (perte des traditions) : young people adopt city culture.
    (Les jeunes adoptent la culture urbaine.)
  · Decrease in food production (baisse de la production alimentaire).

┌─ À RETENIR ──────────────────────────────────────────┐
  L'exode rural crée des problèmes des DEUX côtés :
  · La ville devient surpeuplée et pauvre.
  · Le village se vide et vieillit.
  C'est un cercle vicieux (a vicious cycle).
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

D. LES SOLUTIONS / THE SOLUTIONS

Pour réduire l'exode rural, il faut investir dans les villages.
To reduce rural exodus, we must invest in villages.

Solutions possibles :
  · Build schools and hospitals in rural areas.
    (Construire des écoles et des hôpitaux dans les zones rurales.)
  · Create jobs in agriculture and agribusiness.
    (Créer des emplois dans l'agriculture et l'agro-industrie.)
  · Provide electricity and internet to villages.
    (Fournir l'électricité et internet aux villages.)
  · Support farmers with better tools and training.
    (Soutenir les agriculteurs avec de meilleurs outils et une formation.)

GRAMMAIRE : CONDITIONAL TYPE 1 (Si probable → résultat futur)

Formule :
  IF + sujet + verbe au Present Simple , sujet + WILL + verbe base

Exemples :
  · If the government builds more schools, children will stay in the village.
    (Si le gouvernement construit plus d'écoles, les enfants resteront au village.)
  · If farmers earn more money, they will not leave their land.
    (Si les agriculteurs gagnent plus d'argent, ils ne quitteront pas leurs terres.)
  · If we invest in villages, rural exodus will decrease.
    (Si nous investissons dans les villages, l'exode rural diminuera.)

⚠ ATTENTION — Faux ami !
  "Eventually" en anglais ne veut PAS dire "éventuellement" en français.
  · Eventually (EN) = finalement, un jour
  · Éventuellement (FR) = possibly (EN)
  Exemple : "Eventually, he found a job in Abidjan."
  = "Il a finalement trouvé un emploi à Abidjan." (PAS "éventuellement")

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Les causes et conséquences de l'exode rural]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Zone gauche (fond vert clair) : VILLAGE — label "Rural Area"
      Éléments : icône arbre, champ, maison simple
      Liste des PUSH FACTORS en rouge : poverty · no schools · no hospitals · drought
  · Zone centrale : grande flèche épaisse orange pointant vers la droite
      Label sur la flèche : "RURAL EXODUS"
  · Zone droite (fond gris/béton) : CITY — label "Urban Area"
      Éléments : immeubles, usine, embouteillage
      Liste des PULL FACTORS en vert : jobs · schools · hospitals · modern life
  · En bas à gauche (village) : conséquences en rouge sombre :
      "Ageing population · Abandoned farms · Loss of traditions"
  · En bas à droite (ville) : conséquences en orange foncé :
      "Overpopulation · Slums · Unemployment"
  · Légende : PUSH = raisons de partir | PULL = raisons d'arriver
Note générateur : draw.io ou Canva — utiliser 3 colonnes avec flèche centrale
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Carte mentale : Rural Exodus]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Nœud central (ovale orange) : "RURAL EXODUS"
  · Branche 1 (gauche, bleue) : "DEFINITION"
      → Sous-nœud : "Movement from village to city"
      → Sous-nœud : "A global phenomenon"
  · Branche 2 (haut, rouge) : "CAUSES"
      → Push : poverty / no schools / drought
      → Pull : jobs / education / modern life
  · Branche 3 (droite, verte) : "CONSEQUENCES"
      → Village : ageing / abandoned land
      → City : slums / unemployment
  · Branche 4 (bas, violette) : "SOLUTIONS"
      → Build schools · Create jobs · Invest in villages
  · Légende couleur : bleu = définition | rouge = causes | vert = effets | violet = solutions
Note générateur : MindMeister, Canva ou XMind
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Rural exodus (L'exode rural)
  The massive movement of people from rural areas (villages) to urban areas (cities) in search of better living and working conditions.
  (Le déplacement massif de personnes des zones rurales vers les zones urbaines à la recherche de meilleures conditions de vie et de travail.)

◆ DÉFINITION DPFC — Push factor (Facteur répulsif)
  A negative condition in the place of origin that forces people to leave, such as poverty or lack of services.
  (Une condition négative dans le lieu d'origine qui pousse les gens à partir, comme la pauvreté ou le manque de services.)

◆ DÉFINITION DPFC — Pull factor (Facteur attractif)
  A positive condition in the destination place that attracts people, such as employment opportunities or better schools.
  (Une condition positive dans le lieu de destination qui attire les gens, comme les opportunités d'emploi ou de meilleures écoles.)

◆ DÉFINITION DPFC — Overpopulation (Surpopulation)
  A situation where too many people live in a given area, leading to a lack of resources and services.
  (Une situation où trop de personnes vivent dans une zone donnée, entraînant un manque de ressources et de services.)

◆ DÉFINITION DPFC — Slum (Bidonville)
  A densely populated urban area characterized by poor housing, lack of clean water and sanitation.
  (Une zone urbaine densément peuplée caractérisée par des logements précaires, un manque d'eau potable et d'assainissement.)

◆ DÉFINITION DPFC — Conditional Type 1 (Conditionnel type 1)
  A grammatical structure used to describe a real and possible situation in the future: IF + Present Simple + WILL + base verb.
  (Une structure grammaticale utilisée pour décrire une situation réelle et possible dans le futur : SI + Présent Simple + WILL + verbe base.)

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

I. VOCABULAIRE CLÉS — RURAL EXODUS / L'EXODE RURAL

  Rural exodus = l'exode rural (départ des villages vers les villes)
  To leave / left = quitter (verbe irrégulier)
  To move to = s'installer à / déménager vers
  To look for = chercher
  To hope to = espérer
  Push factor = facteur répulsif (raison de partir)
  Pull factor = facteur attractif (raison d'aller en ville)
  Poverty = la pauvreté
  Unemployment = le chômage
  Overpopulation = la surpopulation
  Slum = le bidonville
  Ageing population = vieillissement de la population
  To invest = investir
  To improve = améliorer / s'améliorer

II. CAUSES DE L'EXODE RURAL / CAUSES OF RURAL EXODUS

  PUSH FACTORS (raisons de partir) :
    · Poverty / pauvreté
    · Lack of schools / manque d'écoles
    · Lack of hospitals / manque d'hôpitaux
    · Drought / sécheresse
    · Bad harvests / mauvaises récoltes

  PULL FACTORS (raisons d'aller en ville) :
    · Jobs / emplois
    · Better schools / meilleures écoles
    · Better healthcare / meilleurs soins de santé
    · Modern life / vie moderne
    · Higher salaries / salaires plus élevés

III. CONSÉQUENCES / CONSEQUENCES

  Pour le village / For the village :
    · Ageing population / vieillissement de la population
    · Abandoned farmland / abandon des terres agricoles
    · Loss of traditions / perte des traditions

  Pour la ville / For the city :
    · Overpopulation / surpopulation
    · Growth of slums / croissance des bidonvilles
    · Unemployment / chômage

IV. GRAMMAIRE — CONDITIONAL TYPE 1

  Formule : IF + Present Simple , WILL + verbe base
  Exemple 1 : If the government builds schools, children will stay in the village.
              (Si le gouvernement construit des écoles, les enfants resteront au village.)
  Exemple 2 : If farmers earn more, they will not leave.
              (Si les agriculteurs gagnent plus, ils ne partiront pas.)
  Exemple 3 : If we invest in villages, rural exodus will decrease.
              (Si nous investissons dans les villages, l'exode rural diminuera.)

V. FAUX AMI À RETENIR

  Eventually (EN) ≠ Éventuellement (FR)
  Eventually = finalement / un jour
  Possibly = éventuellement

────────────────────────────────────────────────────────

~ ANALOGIE CI 1
  Les push factors, c'est comme quand il pleut très fort à Bouaké et que tu cours chercher un abri à Abidjan. Ce n'est pas parce qu'Abidjan est parfaite — c'est parce que la pluie (la pauvreté, le manque d'école) te force à partir.

~ ANALOGIE CI 2
  Les pull factors, c'est comme l'odeur du attiéké et du poisson braisé au bord de la lagune à Abidjan. Cette odeur t'attire même si tu n'avais pas faim au départ. La ville attire les gens avec ses promesses d'emploi et de modernité.

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. Rural exodus = mouvement massif des villages vers les villes.
  2. Il existe deux types de causes : push factors (raisons de partir) et pull factors (raisons d'arriver).
  3. L'exode rural a des conséquences négatives pour le village (abandon, vieillissement) ET pour la ville (surpopulation, bidonvilles).
  4. Le Conditional Type 1 (IF + Present Simple + WILL) sert à proposer des solutions.
  5. Pour réduire l'exode rural, il faut investir dans les zones rurales (écoles, hôpitaux, emplois).

5 mots clés + définition courte :
  · Rural exodus : départ massif des villages vers les villes.
  · Push factor : raison négative qui pousse à quitter un lieu.
  · Pull factor : raison positive qui attire vers un nouveau lieu.
  · Slum : quartier pauvre et surpeuplé en ville.
  · Conditional Type 1 : structure IF + présent → WILL + verbe pour le futur probable.

3 erreurs fréquentes + correction :
  Erreur 1 : Confondre push et pull factors.
  ✔ Correction : Push = raison de PARTIR (négatif au village). Pull = raison d'ARRIVER (positif en ville).

  Erreur 2 : Écrire "If I will study, I will pass." (double will — faux !)
  ✔ Correction : "If I study, I will pass." — après IF, on utilise le Present Simple, JAMAIS will.

  Erreur 3 : Traduire "eventually" par "éventuellement".
  ✔ Correction : "Eventually" = "finalement". "Éventuellement" = "possibly".

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Push or Pull? · Notion ciblée : Push factors / Pull factors

  Énoncé :
  Read the following sentences and write P (Push factor) or PL (Pull factor) for each one.
  Lis les phrases suivantes et écris P (Push factor) ou PL (Pull factor) pour chacune.

  1. "There are no hospitals in my village." → ___
  2. "I can find a well-paid job in Abidjan." → ___
  3. "Our crops are destroyed by drought every year." → ___
  4. "The university in the city offers many courses." → ___
  5. "My children cannot go to school because there is none nearby." → ___

  MÉTHODE
  Étape 1 · Lis chaque phrase et identifie si elle décrit un problème au village ou un avantage en ville.
  Étape 2 · Si c'est un PROBLÈME au village → c'est un Push factor (raison de partir) → écris P.
  Étape 3 · Si c'est un AVANTAGE en ville → c'est un Pull factor (raison d'aller en ville) → écris PL.
  ✔ Vérification : Relis chaque phrase. Push = négatif au village. Pull = positif en ville.
  ✔ Conclusion :
    1. P — Pas d'hôpital au village = raison de partir.
    2. PL — Emploi bien payé en ville = raison d'y aller.
    3. P — Sécheresse et mauvaises récoltes = raison de partir.
    4. PL — Université en ville = raison d'y aller.
    5. P — Pas d'école au village = raison de partir.
  ✔ Ce que cet exercice mobilise : vocabulaire des causes de l'exode rural · distinction push/pull factors · compréhension de phrases en anglais · traduction implicite.

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Conditional Type 1 · Notion ciblée : IF + Present Simple + WILL

  Énoncé :
  Complete the sentences with the correct form of the verbs in brackets.
  Complète les phrases avec la forme correcte des verbes entre parenthèses.

  1. If the government ___ (build) more schools, children ___ (stay) in the village.
  2. If farmers ___ (earn) better salaries, they ___ (not leave) their land.
  3. If we ___ (invest) in rural areas, the quality of life ___ (improve).

  MÉTHODE
  Étape 1 · Rappelle-toi la formule : IF + verbe au Present Simple , WILL + verbe base.
  Étape 2 · Dans la partie IF : conjugue le verbe au Present Simple (ajoute -s à la 3ème personne du singulier si nécessaire).
  Étape 3 · Dans la partie résultat : utilise WILL + verbe à la forme de base (sans -s, sans -ed).
  Étape 4 · Pour la négation : WILL NOT = WON'T + verbe base.
  ✔ Vérification : Vérifie qu'il n'y a jamais de "will" dans la partie IF.
  ✔ Conclusion :
    1. If the government builds more schools, children will stay in the village.
       (Si le gouvernement construit plus d'écoles, les enfants resteront au village.)
    2. If farmers earn better salaries, they will not leave their land.
       (Si les agriculteurs gagnent de meilleurs salaires, ils ne quitteront pas leurs terres.)
    3. If we invest in rural areas, the quality of life will improve.
       (Si nous investissons dans les zones rurales, la qualité de vie s'améliorera.)
  ✔ Ce que cet exercice mobilise : Conditional Type 1 · Present Simple dans la proposition IF · WILL dans la proposition résultat · négation avec WON'T · traduction anglais-français.

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulary matching · Notions travaillées : Vocabulaire de l'exode rural

  Match each English word with its French translation.
  Relie chaque mot anglais à sa traduction française.

  1. Slum                   a. Chômage
  2. Unemployment           b. Vieillissement de la population
  3. Drought                c. Bidonville
  4. Ageing population      d. Sécheresse
  5. To invest              e. Investir

  GUIDE
  Étape 1 · Relis la trace écrite, section I (Vocabulaire clés).
  Étape 2 · Pour chaque mot anglais, cherche sa traduction dans ta liste de vocabulaire.
  Étape 3 · Relie par une flèche ou écris la lettre correspondante.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — True or False · Notions travaillées : Causes et conséquences de l'exode rural

  Read the sentences. Write TRUE or FALSE. Correct the false sentences in English.
  Lis les phrases. Écris VRAI ou FAUX. Corrige les phrases fausses en anglais.

  1. Push factors are positive reasons that attract people to the city.
  2. Rural exodus causes overpopulation in cities.
  3. When young people leave the village, the farmland is often abandoned.
  4. Pull factors include poverty and drought.
  5. Building hospitals in villages can help reduce rural exodus.

  GUIDE
  Étape 1 · Relis la définition de push factors et pull factors dans ta trace écrite.
  Étape 2 · Pour chaque phrase, demande-toi si elle correspond à ce que tu as appris.
  Étape 3 · Si FALSE : réécris la phrase correctement en anglais.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Conditional Type 1 · Notions travaillées : Grammaire — IF + Present Simple + WILL

  Make complete sentences using Conditional Type 1.
  Construis des phrases complètes au Conditional Type 1.

  1. (the government / provide electricity / people / not leave the village)
  2. (young people / get jobs in rural areas / cities / not become overcrowded)
  3. (we / support farmers / food production / increase)
  4. (schools / be built in villages / children / receive a good education)
  5. (the situation / not improve / rural exodus / continue)

  GUIDE
  Étape 1 · Rappelle-toi la formule : IF + Present Simple , WILL + verbe base.
  Étape 2 · Pour la phrase 5 : utilise la forme négative — IF + does not / do not improve.
  Étape 3 · Traduis chaque phrase en français après l'avoir écrite en anglais.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Short reading comprehension · Notions travaillées : Compréhension écrite + vocabulaire

  Read the text and answer the questions in English.
  Lis le texte et réponds aux questions en anglais.

  TEXT:
  "Kofi is 19 years old. He comes from a small village near Yamoussoukro. His father is a farmer, but the harvests have been very bad for three years because of drought. There is only one primary school in the village and no hospital. Last year, Kofi decided to move to Abidjan to look for work. He hopes to find a job in a factory. However, life in Abidjan is not easy. He lives in a small room in a poor neighbourhood with five other people."

  Questions :
  1. Where does Kofi come from?
  2. Why did Kofi leave his village? Give TWO reasons.
  3. What does Kofi hope to find in Abidjan?
  4. What problem does Kofi face in Abidjan?
  5. Identify ONE push factor and ONE pull factor in the text.

  GUIDE
  Étape 1 · Lis le texte une première fois sans t'arrêter pour comprendre l'idée générale.
  Étape 2 · Lis les questions, puis relis le texte en cherchant les informations précises.
  Étape 3 · Réponds en anglais avec des phrases complètes : sujet + verbe + complément.
  Étape 4 · Pour la question 5 : rappelle-toi la définition de push (négatif au village) et pull (positif en ville).

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Writing sentences · Notions travaillées : Expression écrite · causes et solutions · vocabulaire

  Write FIVE sentences in English about rural exodus. Use the words given.
  Écris CINQ phrases en anglais sur l'exode rural. Utilise les mots donnés.

  Mots à utiliser (un mot différent par phrase) :
  1. poverty
  2. slum
  3. invest
  4. ageing population
  5. because of

  GUIDE
  Étape 1 · Pour chaque mot, demande-toi s'il parle d'une cause, d'une conséquence ou d'une solution.
  Étape 2 · Construis une phrase complète avec sujet + verbe + complément.
  Étape 3 · Traduis chaque phrase en français après l'avoir écrite.
  Étape 4 · Relis tes phrases : vérifie l'accord des verbes et l'ordre des mots.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — My village, my city
  Durée conseillée : 30 minutes

  Sujet :
  Write a short text of 8 to 10 sentences in English.
  (Écris un court texte de 8 à 10 phrases en anglais.)

  Imagine you are Aminata, a 16-year-old girl from a village in the north of Côte d'Ivoire. You have just moved to Abidjan. Write about:
  · Why you left your village (at least 2 push factors).
  · What attracted you to Abidjan (at least 2 pull factors).
  · One problem you have found in Abidjan.
  · One hope you have for the future (use Conditional Type 1).

  Imagine que tu es Aminata, une fille de 16 ans venant d'un village du nord de la Côte d'Ivoire. Tu viens de t'installer à Abidjan. Écris sur :
  · Pourquoi tu as quitté ton village (au moins 2 push factors).
  · Ce qui t'a attirée à Abidjan (au moins 2 pull factors).
  · Un problème que tu as trouvé à Abidjan.
  · Un espoir pour l'avenir (utilise le Conditional Type 1).

  Contraintes obligatoires :
  · 8 à 10 phrases minimum.
  · Utiliser au moins 5 mots du vocabulaire de la leçon.
  · Utiliser AU MOINS UNE phrase au Conditional Type 1.
  · Écrire à la première personne (I...).
  · Traduire ton texte en français après l'avoir écrit en anglais.

  GUIDE (sans corrigé)
  Étape 1 · Relis la trace écrite : push factors, pull factors, vocabulaire clé.
  Étape 2 · Prépare tes idées au brouillon : note 2 push factors, 2 pull factors, 1 problème, 1 espoir.
  Étape 3 · Écris ton texte phrase par phrase. Commence par : "My name is Aminata. I come from..."
  Étape 4 · Intègre ton Conditional Type 1 pour l'espoir : "If..., I will..."
  Étape 5 · Relis et corrige : vérifie les verbes, la conjugaison, l'orthographe.
  Étape 6 · Traduis ton texte en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON : My village, my city

Proposition de corrigé (texte modèle en anglais) :

My name is Aminata. I am 16 years old and I come from a small village near Korhogo, in the north of Côte d'Ivoire. I decided to move to Abidjan last month.

I left my village because life was very hard there. First, there was no secondary school in my village, so I could not continue my education. Second, my family was very poor because of drought. Our crops were destroyed and my father could not earn enough money.

Abidjan attracted me because there are many schools and universities here. I also hoped to find a part-time job to help my family.

However, life in Abidjan is not easy. I live in a small room in a crowded neighbourhood. There is a lot of noise and it is very expensive.

But I remain hopeful. If I work hard, I will pass my BEPC exam. If the government invests in my village, maybe I will go back one day.

Traduction française complète :
Je m'appelle Aminata. J'ai 16 ans et je viens d'un petit village près de Korhogo, dans le nord de la Côte d'Ivoire. J'ai décidé de déménager à Abidjan le mois dernier.

J'ai quitté mon village parce que la vie y était très difficile. Premièrement, il n'y avait pas de collège dans mon village, donc je ne pouvais pas continuer mes études. Deuxièmement, ma famille était très pauvre à cause de la sécheresse. Nos récoltes étaient détruites et mon père ne pouvait pas gagner suffisamment d'argent.

Abidjan m'a attirée parce qu'il y a de nombreuses écoles et universités ici. J'espérais aussi trouver un emploi à temps partiel pour aider ma famille.

Cependant, la vie à Abidjan n'est pas facile. Je vis dans une petite chambre dans un quartier surpeuplé. Il y a beaucoup de bruit et c'est très cher.

Mais je garde espoir. Si je travaille dur, je réussirai mon BEPC. Si le gouvernement investit dans mon village, peut-être que j'y retournerai un jour.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1 : Vocabulary matching

  1. Slum → c. Bidonville
  2. Unemployment → a. Chômage
  3. Drought → d. Sécheresse
  4. Ageing population → b. Vieillissement de la population
  5. To invest → e. Investir

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2 : True or False

  1. FALSE — Push factors are NEGATIVE reasons that PUSH people to LEAVE the village. Pull factors are the positive reasons.
     (FAUX — Les push factors sont des raisons NÉGATIVES qui poussent les gens à QUITTER le village. Les pull factors sont les raisons positives.)

  2. TRUE — Rural exodus causes overpopulation in cities.
     (VRAI — L'exode rural provoque la surpopulation dans les villes.)

  3. TRUE — When young people leave the village, the farmland is often abandoned.
     (VRAI — Quand les jeunes quittent le village, les terres agricoles sont souvent abandonnées.)

  4. FALSE — Pull factors include jobs and better schools, NOT poverty and drought. Poverty and drought are PUSH factors.
     (FAUX — Les pull factors incluent les emplois et les meilleures écoles, PAS la pauvreté et la sécheresse. La pauvreté et la sécheresse sont des push factors.)

  5. TRUE — Building hospitals in villages can help reduce rural exodus.
     (VRAI — Construire des hôpitaux dans les villages peut aider à réduire l'exode rural.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3 : Conditional Type 1

  1. If the government provides electricity, people will not leave the village.
     (Si le gouvernement fournit l'électricité, les gens ne quitteront pas le village.)

  2. If young people get jobs in rural areas, cities will not become overcrowded.
     (Si les jeunes trouvent des emplois dans les zones rurales, les villes ne deviendront pas surpeuplées.)

  3. If we support farmers, food production will increase.
     (Si nous soutenons les agriculteurs, la production alimentaire augmentera.)

  4. If schools are built in villages, children will receive a good education.
     (Si des écoles sont construites dans les villages, les enfants recevront une bonne éducation.)

  5. If the situation does not improve, rural exodus will continue.
     (Si la situation ne s'améliore pas, l'exode rural continuera.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4 : Short reading comprehension

  1. Kofi comes from a small village near Yamoussoukro.
     (Kofi vient d'un petit village près de Yamoussoukro.)

  2. Kofi left his village because of drought (bad harvests) AND because there was no hospital / only one primary school.
     (Kofi a quitté son village à cause de la sécheresse et parce qu'il n'y avait pas d'hôpital / qu'une seule école primaire.)

  3. Kofi hopes to find a job in a factory.
     (Kofi espère trouver un emploi dans une usine.)

  4. Life in Abidjan is not easy: he lives in a small room with five other people in a poor neighbourhood.
     (La vie à Abidjan n'est pas facile : il vit dans une petite chambre avec cinq autres personnes dans un quartier pauvre.)

  5. Push factor: the harvests have been very bad because of drought. / There is no hospital.
     (Facteur répulsif : les récoltes sont mauvaises à cause de la sécheresse. / Il n'y a pas d'hôpital.)
     Pull factor: he hopes to find a job in a factory.
     (Facteur attractif : il espère trouver un emploi dans une usine.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5 : Writing sentences

  Proposition de corrigé (d'autres réponses sont possibles) :

  1. poverty :
     Many people leave their village because of poverty.
     (Beaucoup de gens quittent leur village à cause de la pauvreté.)

  2. slum :
     In Abidjan, many migrants from rural areas live in slums.
     (À Abidjan, de nombreux migrants des zones rurales vivent dans des bidonvilles.)

  3. invest :
     The government must invest in rural areas to stop rural exodus.
     (Le gouvernement doit investir dans les zones rurales pour stopper l'exode rural.)

  4. ageing population :
     When young people leave, the village is left with an ageing population.
     (Quand les jeunes partent, le village se retrouve avec une population vieillissante.)

  5. because of :
     Because of the lack of schools, many children move to the city.
     (À cause du manque d'écoles, de nombreux enfants partent en ville.)

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "Don't watch the clock; do what it does. Keep going."
  (Ne regarde pas l'horloge ; fais comme elle. Continue d'avancer.)

  — Sam Levenson

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
