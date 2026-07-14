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

COURS = """✅ Anglais LV1 · Leçon 18 — Tolerance · en cours de génération...

---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 18 · Classe de 3ème

TOLERANCE
La tolérance

Compétence C6 · Écoute élaborée — Droits humains / Human Rights

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 18 — TOLERANCE
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C6 · Écoute élaborée — Droits humains / Human Rights
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 17 : My duties)
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

CAPSULE DE REPRISE — Leçon 17 : My duties (Mes devoirs)

5 points clés à se rappeler :
  1. Un devoir est une responsabilité morale ou légale envers les autres, la société et l'environnement.
  2. Il existe 4 domaines de devoirs : à l'école · en tant que citoyen · envers les autres · envers l'environnement.
  3. HAVE TO / HAS TO = obligation externe imposée par une règle ou une loi.
  4. DON'T / DOESN'T HAVE TO = pas d'obligation — ce n'est pas nécessaire mais c'est permis.
  5. MUST NOT ≠ DON'T HAVE TO : MUST NOT = interdit · DON'T HAVE TO = pas nécessaire.

3 questions flash :
  Q1 : What is the difference between MUST NOT and DON'T HAVE TO?
       (Quelle est la différence entre MUST NOT et DON'T HAVE TO ?)
  R1 : MUST NOT = forbidden (interdit). DON'T HAVE TO = not necessary but allowed (pas nécessaire mais permis).

  Q2 : Complete: "She ___ wear her uniform. It is the school rule." (HAS TO / HAVE TO ?)
  R2 : She has to wear her uniform. (Sujet singulier "she" → HAS TO.)

  Q3 : Give ONE duty towards the environment. (Donne UN devoir envers l'environnement.)
  R3 : We should avoid throwing rubbish in the streets. / We have to protect forests.

Lien avec la leçon du jour :
  Dans la leçon 17, nous avons vu que chaque citoyen a des devoirs envers les autres et envers la société. Parmi ces devoirs figure le respect des différences. Aujourd'hui, nous allons aller plus loin et apprendre à parler de la tolérance en anglais — c'est le thème de la leçon 18, et c'est aussi la dernière leçon de la Compétence C6.

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

La Côte d'Ivoire est un pays de plus de 60 ethnies et de nombreuses religions : christianisme, islam, religions traditionnelles. À Abidjan, il n'est pas rare de voir une église et une mosquée côte à côte dans le même quartier. Dans les cours d'école, des élèves dioula, baoulé, bété, sénoufo et bien d'autres apprennent ensemble, jouent ensemble, grandissent ensemble.

Pourtant, la Côte d'Ivoire a aussi connu des périodes de crise liées à l'intolérance et à la division. C'est pourquoi la tolérance est une valeur fondamentale enseignée dans les écoles ivoiriennes.

Le mot d'ordre de la Côte d'Ivoire : "Union — Discipline — Travail."
L'union commence par la tolérance.
Unity begins with tolerance.

Lien DPFC : Ce thème est directement lié à l'éducation civique, à la paix et à la cohésion sociale, valeurs prioritaires du programme DPFC CI.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — Qu'est-ce que la tolérance ?
  Q : What does it mean to be tolerant? Is tolerance the same as accepting everything?
      (Qu'est-ce que cela signifie d'être tolérant ? La tolérance signifie-t-elle accepter tout ?)
  Guide :
    · La tolérance, c'est accepter que les autres soient différents de soi.
    · Mais tolérance ne veut pas dire tout accepter — on ne tolère pas la violence ou l'injustice.
    · Pense à la différence entre respecter une opinion et approuver un comportement nuisible.
  Corrigé :
    To be tolerant means to accept and respect people who are different from us, whether in terms of religion, culture, ethnicity or opinion. However, tolerance does not mean accepting everything. We do not have to tolerate violence, injustice or the violation of others' rights. True tolerance is about respecting differences while still standing up for what is right and just.
    (Être tolérant signifie accepter et respecter les personnes différentes de nous, que ce soit en termes de religion, de culture, d'ethnie ou d'opinion. Cependant, la tolérance ne signifie pas tout accepter. Nous n'avons pas à tolérer la violence, l'injustice ou la violation des droits des autres. La vraie tolérance consiste à respecter les différences tout en défendant ce qui est juste.)

▶ HARKNESS 2 — Tolérance et vivre-ensemble
  Q : How can tolerance help build a peaceful society?
      (Comment la tolérance peut-elle aider à construire une société pacifique ?)
  Guide :
    · Pense aux conflits qui naissent de l'intolérance : guerres ethniques, discriminations, exclusions.
    · Pense à ce que la tolérance apporte : dialogue, coopération, paix.
    · Utilise : peace (paix) · dialogue · cooperation (coopération) · diversity (diversité).
  Corrigé :
    Tolerance helps build a peaceful society because it encourages people to listen to each other, to cooperate and to resolve conflicts through dialogue rather than violence. When people accept diversity — different religions, cultures and opinions — they can live together harmoniously. In contrast, intolerance leads to discrimination, conflict and division. Côte d'Ivoire's motto, "Union — Discipline — Travail", reminds us that unity and peace begin with mutual respect and tolerance.
    (La tolérance aide à construire une société pacifique parce qu'elle encourage les gens à s'écouter, à coopérer et à résoudre les conflits par le dialogue plutôt que par la violence. Quand les gens acceptent la diversité — différentes religions, cultures et opinions — ils peuvent vivre ensemble harmonieusement. À l'inverse, l'intolérance conduit à la discrimination, aux conflits et à la division. La devise de la Côte d'Ivoire, "Union — Discipline — Travail", nous rappelle que l'unité et la paix commencent par le respect mutuel et la tolérance.)

▶ HARKNESS 3 — Tolérance à l'école
  Q : How can students practise tolerance at school every day?
      (Comment les élèves peuvent-ils pratiquer la tolérance à l'école au quotidien ?)
  Guide :
    · Pense aux situations concrètes : différences d'ethnie, de religion, d'opinion, de capacités.
    · Pense à des comportements tolérants : écouter, ne pas se moquer, inclure tout le monde.
    · Utilise : include (inclure) · listen (écouter) · mock (se moquer) · bully (intimider).
  Corrigé :
    Students can practise tolerance at school in many ways. They should listen to their classmates' opinions even when they disagree. They must not mock or bully students who are different from them, whether because of their ethnic group, religion, physical appearance or academic level. They should include everyone in group activities and help classmates who struggle. Small daily acts of kindness and respect are the foundation of a tolerant school community.
    (Les élèves peuvent pratiquer la tolérance à l'école de nombreuses façons. Ils devraient écouter les opinions de leurs camarades même quand ils ne sont pas d'accord. Ils ne doivent pas se moquer des élèves différents d'eux, que ce soit en raison de leur groupe ethnique, de leur religion, de leur apparence physique ou de leur niveau scolaire. Ils devraient inclure tout le monde dans les activités de groupe et aider les camarades en difficulté. Les petits actes quotidiens de gentillesse et de respect sont le fondement d'une communauté scolaire tolérante.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre de la leçon : Tolerance (La tolérance)
Compétence : C6 — Écoute élaborée · Droits humains / Human Rights

Objectifs de la leçon :
  À la fin de cette leçon, tu seras capable de :
  · Définir la tolérance et l'intolérance en anglais.
  · Expliquer l'importance de la tolérance dans la société et à l'école.
  · Utiliser les connecteurs logiques de cause, conséquence et opposition (because, so, but, however, although, therefore) pour construire un texte argumentatif.
  · Utiliser MUST / MUST NOT / SHOULD / HAVE TO pour exprimer des comportements tolérants et intolérants.
  · Produire un texte argumentatif court sur la tolérance.

Prérequis nécessaires :
  · MUST / MUST NOT / SHOULD (Leçon 16).
  · HAVE TO / DON'T HAVE TO (Leçon 17).
  · Le vocabulaire des droits et des devoirs (Leçons 16 et 17).
  · Le Present Simple et le Conditional Type 1 (Leçon 15).

────────────────────────────────────────────────────────

TABLEAU DES STRUCTURES FONDAMENTALES — Leçon 18

| Expression / Notion              | Valeur, emploi                                         | Exemple anglais                                                  | Traduction française                                                   |
|---------------------------------|--------------------------------------------------------|------------------------------------------------------------------|------------------------------------------------------------------------|
| Tolerance                       | La tolérance — nom                                     | Tolerance is essential for peace.                                | La tolérance est essentielle pour la paix.                             |
| Intolerance                     | L'intolérance — nom                                    | Intolerance leads to conflict.                                   | L'intolérance mène au conflit.                                         |
| To tolerate                     | Tolérer — verbe                                        | We must tolerate different opinions.                             | Nous devons tolérer les opinions différentes.                          |
| To respect                      | Respecter — verbe                                      | We should respect all cultures.                                  | Nous devrions respecter toutes les cultures.                           |
| To accept                       | Accepter — verbe                                       | Learning to accept differences is a sign of maturity.            | Apprendre à accepter les différences est un signe de maturité.         |
| Diversity                       | La diversité                                           | Diversity makes our society richer.                              | La diversité rend notre société plus riche.                            |
| Prejudice                       | Le préjugé                                             | Prejudice is often based on ignorance.                           | Le préjugé est souvent fondé sur l'ignorance.                          |
| Discrimination                  | La discrimination                                      | Discrimination must not be tolerated.                            | La discrimination ne doit pas être tolérée.                            |
| To mock / to bully              | Se moquer de / intimider                               | We must not mock or bully others.                                | Nous ne devons pas nous moquer des autres ni les intimider.            |
| However                         | Cependant / pourtant (opposition)                      | Tolerance is important. However, it has limits.                  | La tolérance est importante. Cependant, elle a des limites.            |
| Although                        | Bien que / même si (opposition dans une phrase)        | Although we are different, we can live together.                 | Bien que nous soyons différents, nous pouvons vivre ensemble.          |
| Therefore                       | Donc / par conséquent (conséquence formelle)           | We are all different; therefore, we must respect each other.     | Nous sommes tous différents ; par conséquent, nous devons nous respecter.|
| To coexist                      | Coexister / vivre ensemble                             | Different religions can coexist peacefully.                      | Différentes religions peuvent coexister pacifiquement.                 |
| Peaceful / peacefully           | Pacifique / pacifiquement                              | We must learn to live together peacefully.                       | Nous devons apprendre à vivre ensemble pacifiquement.                  |

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

A. QU'EST-CE QUE LA TOLÉRANCE ? / WHAT IS TOLERANCE?

Tolerance is the ability to accept and respect people who are different from us, whether in terms of religion, culture, ethnicity, opinion or way of life.
(La tolérance est la capacité d'accepter et de respecter des personnes différentes de nous, que ce soit en termes de religion, de culture, d'ethnie, d'opinion ou de mode de vie.)

The opposite of tolerance is intolerance.
(Le contraire de la tolérance est l'intolérance.)

Intolerance leads to prejudice, discrimination, conflict and violence.
(L'intolérance mène aux préjugés, à la discrimination, aux conflits et à la violence.)

┌─ À RETENIR ──────────────────────────────────────────┐
  Tolerance = tolérance (accepter les différences)
  Intolerance = intolérance (refus des différences)
  Diversity = diversité (variété de cultures, religions, opinions)
  Prejudice = préjugé (opinion négative sans fondement réel)
  Discrimination = discrimination (traitement injuste basé sur les différences)
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

B. POURQUOI LA TOLÉRANCE EST-ELLE IMPORTANTE ?
   WHY IS TOLERANCE IMPORTANT?

1. FOR PEACE AND SOCIAL COHESION (Pour la paix et la cohésion sociale) :
  · Tolerance allows people of different backgrounds to live together peacefully.
    (La tolérance permet à des personnes d'origines différentes de vivre ensemble pacifiquement.)
  · Without tolerance, societies fall apart because of conflicts and divisions.
    (Sans tolérance, les sociétés se désintègrent à cause des conflits et des divisions.)
  · In Côte d'Ivoire, with over 60 ethnic groups, tolerance is the foundation of national unity.
    (En Côte d'Ivoire, avec plus de 60 groupes ethniques, la tolérance est le fondement de l'unité nationale.)

2. FOR PERSONAL GROWTH (Pour l'épanouissement personnel) :
  · Meeting people who are different from us helps us learn and grow.
    (Rencontrer des personnes différentes de nous nous aide à apprendre et à grandir.)
  · Tolerance opens our minds and helps us understand the world better.
    (La tolérance ouvre notre esprit et nous aide à mieux comprendre le monde.)

3. FOR JUSTICE AND EQUALITY (Pour la justice et l'égalité) :
  · Tolerance means treating every person with dignity and respect.
    (La tolérance signifie traiter chaque personne avec dignité et respect.)
  · It means fighting against discrimination and prejudice.
    (Elle signifie lutter contre la discrimination et les préjugés.)

┌─ À RETENIR ──────────────────────────────────────────┐
  La tolérance est importante pour 3 raisons :
  1. Elle garantit la paix et la cohésion sociale.
  2. Elle favorise l'épanouissement personnel.
  3. Elle est le fondement de la justice et de l'égalité.
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

C. LES FORMES D'INTOLÉRANCE / FORMS OF INTOLERANCE

Intolerance can take many forms in everyday life :
(L'intolérance peut prendre de nombreuses formes dans la vie quotidienne :)

  · Racism (racisme) : judging someone based on their race or skin colour.
    (Juger quelqu'un en fonction de sa race ou de la couleur de sa peau.)
  · Religious intolerance (intolérance religieuse) : refusing to respect another person's religion.
    (Refuser de respecter la religion d'une autre personne.)
  · Ethnic discrimination (discrimination ethnique) : treating people differently because of their ethnic group.
    (Traiter les gens différemment en raison de leur groupe ethnique.)
  · Bullying (intimidation) : mocking or attacking someone because they are different.
    (Se moquer de quelqu'un ou l'attaquer parce qu'il est différent.)
  · Sexism (sexisme) : treating people unfairly based on their gender.
    (Traiter les gens injustement en raison de leur genre.)

⚠ ATTENTION
  All forms of intolerance are violations of human rights.
  (Toutes les formes d'intolérance sont des violations des droits humains.)
  We must not tolerate intolerance — but we must fight it with words, not violence.
  (Nous ne devons pas tolérer l'intolérance — mais nous devons la combattre avec des mots, pas avec la violence.)

────────────────────────────────────────────────────────

D. GRAMMAIRE — LES CONNECTEURS LOGIQUES

Les connecteurs logiques permettent de construire un texte argumentatif solide.
Ils relient les idées entre elles de façon claire et logique.
Ils sont indispensables pour le BEPC.

1. CAUSE (pourquoi ?)
  · Because + sujet + verbe :
    People fight because they do not understand each other.
    (Les gens se battent parce qu'ils ne se comprennent pas.)
  · Because of + nom :
    Conflicts arise because of intolerance.
    (Les conflits naissent à cause de l'intolérance.)

2. CONSÉQUENCE (donc, résultat)
  · So :
    We are all different, so we must learn to respect each other.
    (Nous sommes tous différents, donc nous devons apprendre à nous respecter.)
  · Therefore (plus formel) :
    Intolerance leads to violence; therefore, we must fight it.
    (L'intolérance mène à la violence ; par conséquent, nous devons la combattre.)

3. OPPOSITION (mais, cependant, bien que)
  · But :
    Tolerance is important, but it does not mean accepting injustice.
    (La tolérance est importante, mais elle ne signifie pas accepter l'injustice.)
  · However (plus formel — en début de phrase ou après un point-virgule) :
    Tolerance is a value. However, it has limits.
    (La tolérance est une valeur. Cependant, elle a des limites.)
  · Although + sujet + verbe (opposition dans une même phrase) :
    Although we are different, we can live together peacefully.
    (Bien que nous soyons différents, nous pouvons vivre ensemble pacifiquement.)

⚠ ATTENTION — Faux ami !
  "Sympathetic" en anglais ≠ "Sympathique" en français.
  · Sympathetic (EN) = compatissant, qui comprend la douleur des autres
  · Sympathique (FR) = nice / friendly (EN)
  Exemple : "She was very sympathetic when I told her about my problems."
  = "Elle était très compréhensive quand je lui ai parlé de mes problèmes."
  (PAS "elle était très sympathique")

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Tolérance vs Intolérance : causes et conséquences]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Deux colonnes principales séparées par une ligne centrale verticale
  · Colonne gauche (fond vert) : "TOLERANCE / LA TOLÉRANCE"
      Causes/sources : respect · open mind · education · diversity
      Conséquences (flèches vers le bas) :
        → Peace (paix)
        → Social cohesion (cohésion sociale)
        → Personal growth (épanouissement)
        → Justice and equality (justice et égalité)
  · Colonne droite (fond rouge pâle) : "INTOLERANCE / L'INTOLÉRANCE"
      Causes/sources : prejudice · ignorance · fear · discrimination
      Conséquences (flèches vers le bas) :
        → Conflict (conflit)
        → Division (division)
        → Violence
        → Injustice
  · Encadré central (fond jaune) entre les deux colonnes :
      "CHOICE / CHOIX" avec deux flèches pointant vers la gauche et vers la droite
  · Bas de page : "In Côte d'Ivoire : Union — Discipline — Travail"
Note générateur : Canva ou draw.io — 2 colonnes contrastées vert/rouge avec flèches
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Carte mentale des connecteurs logiques]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Nœud central (ovale bleu foncé) : "CONNECTEURS LOGIQUES"
  · Branche 1 (gauche, fond orange) : "CAUSE / POURQUOI ?"
      → because + sujet + verbe
      → because of + nom
      Exemple : "...because of intolerance"
  · Branche 2 (droite, fond vert) : "CONSÉQUENCE / RÉSULTAT"
      → so (informel)
      → therefore (formel)
      Exemple : "...therefore, we must act"
  · Branche 3 (bas, fond violet) : "OPPOSITION / CONTRASTE"
      → but (informel)
      → however (formel — début de phrase)
      → although + sujet + verbe (même phrase)
      Exemple : "Although we differ, we can coexist"
  · Légende : orange = cause | vert = conséquence | violet = opposition
Note générateur : MindMeister ou Canva — disposition en étoile à 3 branches
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Tolerance / La tolérance
  The ability to accept and respect people who are different from us in terms of religion, culture, ethnicity, opinion or way of life, without judging or discriminating against them.
  (La capacité d'accepter et de respecter des personnes différentes de nous en termes de religion, de culture, d'ethnie, d'opinion ou de mode de vie, sans les juger ni les discriminer.)

◆ DÉFINITION DPFC — Intolerance / L'intolérance
  The refusal to accept or respect differences between people, often leading to prejudice, discrimination and conflict.
  (Le refus d'accepter ou de respecter les différences entre les personnes, menant souvent aux préjugés, à la discrimination et aux conflits.)

◆ DÉFINITION DPFC — Diversity / La diversité
  The presence of a wide variety of cultures, religions, ethnicities, opinions and ways of life within a community or society.
  (La présence d'une grande variété de cultures, de religions, d'ethnies, d'opinions et de modes de vie au sein d'une communauté ou d'une société.)

◆ DÉFINITION DPFC — Prejudice / Le préjugé
  A negative opinion or judgment formed about a person or group without real knowledge or evidence, often based on stereotypes.
  (Une opinion ou un jugement négatif formé sur une personne ou un groupe sans connaissance ou preuve réelle, souvent basé sur des stéréotypes.)

◆ DÉFINITION DPFC — Connecteurs logiques / Logical connectors
  Words or phrases used to link ideas in a text and show the logical relationships between them: cause (because, because of), consequence (so, therefore), opposition (but, however, although).
  (Des mots ou des expressions utilisés pour relier des idées dans un texte et montrer les relations logiques entre elles : cause (because, because of), conséquence (so, therefore), opposition (but, however, although).)

◆ DÉFINITION DPFC — Coexistence / La coexistence
  The ability of people with different backgrounds, beliefs or cultures to live together peacefully in the same society.
  (La capacité de personnes d'origines, de croyances ou de cultures différentes à vivre ensemble pacifiquement dans la même société.)

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

I. VOCABULAIRE CLÉS — TOLERANCE / LA TOLÉRANCE

  Tolerance = tolérance (accepter les différences)
  Intolerance = intolérance (refus des différences)
  To tolerate = tolérer
  To accept = accepter
  To respect = respecter
  To coexist = coexister / vivre ensemble
  Diversity = diversité
  Prejudice = préjugé
  Discrimination = discrimination
  Racism = racisme
  Bullying = intimidation / harcèlement
  Sexism = sexisme
  To mock = se moquer de
  Peaceful / peacefully = pacifique / pacifiquement
  Therefore = par conséquent / donc (formel)
  However = cependant / pourtant (formel)
  Although = bien que / même si
  Sympathetic (EN) = compatissant / compréhensif (≠ sympathique en français !)

II. POURQUOI LA TOLÉRANCE EST-ELLE IMPORTANTE ?

  1. Pour la paix et la cohésion sociale :
     Tolerance allows people to live together peacefully.
     (La tolérance permet aux gens de vivre ensemble pacifiquement.)

  2. Pour l'épanouissement personnel :
     Meeting different people helps us learn and grow.
     (Rencontrer des personnes différentes nous aide à apprendre et à grandir.)

  3. Pour la justice et l'égalité :
     Tolerance means treating every person with dignity and respect.
     (La tolérance signifie traiter chaque personne avec dignité et respect.)

III. FORMES D'INTOLÉRANCE / FORMS OF INTOLERANCE

  · Racism (racisme) : judging someone based on race or skin colour.
  · Religious intolerance (intolérance religieuse) : refusing to respect another religion.
  · Ethnic discrimination (discrimination ethnique) : treating people differently because of their ethnic group.
  · Bullying (intimidation) : mocking or attacking someone because they are different.
  · Sexism (sexisme) : unfair treatment based on gender.

IV. GRAMMAIRE — CONNECTEURS LOGIQUES

  CAUSE :
    because + sujet + verbe → People fight because they do not understand each other.
    because of + nom → Conflicts arise because of intolerance.

  CONSÉQUENCE :
    so → We are different, so we must respect each other.
    therefore → Intolerance leads to violence; therefore, we must fight it.

  OPPOSITION :
    but → Tolerance is important, but it has limits.
    however → Tolerance is a value. However, it has limits.
    although + sujet + verbe → Although we are different, we can live together peacefully.

V. FAUX AMI À RETENIR

  Sympathetic (EN) ≠ Sympathique (FR)
  Sympathetic = compatissant, compréhensif
  Nice / friendly = sympathique
  Exemple : "My teacher was very sympathetic." = "Mon professeur était très compréhensif."

────────────────────────────────────────────────────────

~ ANALOGIE CI 1
  La tolérance, c'est comme le plat ivoirien de foutou avec sauce graine : il y a de l'igname, de la banane plantain, du manioc — tous différents — mais ensemble, ils forment quelque chose de délicieux. Une société tolérante, c'est pareil : des personnes d'ethnies, de religions et d'opinions différentes qui, ensemble, créent quelque chose de beau et de fort.

~ ANALOGIE CI 2
  Les connecteurs logiques, c'est comme les ponts d'Abidjan qui relient les différentes communes. Sans eux, chaque idée reste isolée, comme une île coupée du reste. BECAUSE relie la cause à l'effet, SO et THEREFORE montrent où on va, BUT et HOWEVER signalent un tournant inattendu.

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. La tolérance est la capacité d'accepter et de respecter les différences des autres : religion, culture, ethnie, opinion.
  2. L'intolérance mène aux préjugés, à la discrimination, aux conflits et à la violence.
  3. La tolérance est importante pour la paix, l'épanouissement personnel et la justice.
  4. Les connecteurs logiques structurent un texte argumentatif : BECAUSE (cause) · SO / THEREFORE (conséquence) · BUT / HOWEVER / ALTHOUGH (opposition).
  5. Tolérance ne signifie pas tout accepter : on ne tolère pas la violence ni l'injustice.

5 mots clés + définition courte :
  · Tolerance : capacité d'accepter et de respecter les différences.
  · Intolerance : refus des différences, source de conflits.
  · Diversity : variété de cultures, religions et opinions dans une société.
  · Prejudice : opinion négative sans fondement réel, basée sur l'ignorance.
  · Coexistence : capacité de vivre ensemble pacifiquement malgré les différences.

3 erreurs fréquentes + correction :
  Erreur 1 : Utiliser HOWEVER comme BUT au milieu d'une phrase.
  ✔ Correction : HOWEVER se place en début de phrase (après un point ou un point-virgule), jamais entre deux propositions comme BUT. Exemple correct : "Tolerance is important. However, it has limits." Exemple incorrect : "Tolerance is important however it has limits."

  Erreur 2 : Confondre BECAUSE et BECAUSE OF.
  ✔ Correction : BECAUSE est suivi d'un sujet + verbe. BECAUSE OF est suivi d'un nom ou d'un groupe nominal. Exemple : "He left because he was afraid." / "He left because of fear." (PAS "because of he was afraid.")

  Erreur 3 : Traduire "sympathetic" par "sympathique".
  ✔ Correction : Sympathetic (EN) = compatissant / compréhensif. Nice ou friendly = sympathique en français.

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Connecteurs logiques · Notion ciblée : Because / So / However / Although

  Énoncé :
  Fill in the blanks with the correct connector: BECAUSE, BECAUSE OF, SO, HOWEVER, ALTHOUGH.
  Complète les blancs avec le bon connecteur : BECAUSE, BECAUSE OF, SO, HOWEVER, ALTHOUGH.

  1. Côte d'Ivoire has over 60 ethnic groups, ___ tolerance is essential for national unity.
  2. ___ we are all different, we can live together peacefully.
  3. Some people are intolerant ___ they fear what they do not know.
  4. Tolerance is very important. ___, it does not mean accepting violence or injustice.
  5. Conflicts often arise ___ prejudice and ignorance.

  MÉTHODE
  Étape 1 · Lis chaque phrase et identifie la relation logique entre les deux parties : cause ? conséquence ? opposition ?
  Étape 2 · Cause → BECAUSE (+ sujet + verbe) ou BECAUSE OF (+ nom).
  Étape 3 · Conséquence → SO (informel, milieu de phrase) ou THEREFORE (formel, début de phrase).
  Étape 4 · Opposition → ALTHOUGH (+ sujet + verbe, même phrase) ou HOWEVER (début de nouvelle phrase).
  Étape 5 · Vérifie : HOWEVER commence toujours une nouvelle phrase ou suit un point-virgule.
  ✔ Vérification : Relis chaque phrase complétée. La relation logique est-elle claire ?
  ✔ Conclusion :
    1. SO — Côte d'Ivoire has over 60 ethnic groups, so tolerance is essential for national unity.
       (La Côte d'Ivoire a plus de 60 groupes ethniques, donc la tolérance est essentielle pour l'unité nationale.)
    2. ALTHOUGH — Although we are all different, we can live together peacefully.
       (Bien que nous soyons tous différents, nous pouvons vivre ensemble pacifiquement.)
    3. BECAUSE — Some people are intolerant because they fear what they do not know.
       (Certaines personnes sont intolérantes parce qu'elles craignent ce qu'elles ne connaissent pas.)
    4. HOWEVER — Tolerance is very important. However, it does not mean accepting violence or injustice.
       (La tolérance est très importante. Cependant, elle ne signifie pas accepter la violence ou l'injustice.)
    5. BECAUSE OF — Conflicts often arise because of prejudice and ignorance.
       (Les conflits naissent souvent à cause des préjugés et de l'ignorance.)
  ✔ Ce que cet exercice mobilise : connecteurs logiques · distinction BECAUSE / BECAUSE OF · SO / HOWEVER / ALTHOUGH · relations cause-conséquence-opposition · traduction anglais-français.

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Texte argumentatif · Notion ciblée : Production écrite sur la tolérance

  Énoncé :
  Read the following short text and answer the questions.
  Lis le court texte suivant et réponds aux questions.

  TEXT :
  "In our school, we come from different regions of Côte d'Ivoire. Some students are from the north, others from the south, the east or the west. We have different religions and different mother tongues. Although we are very different, we get along well because we respect each other. However, sometimes there are misunderstandings. When that happens, we talk and find solutions together. Tolerance is not just a word — it is a daily practice."

  Questions :
  1. How many regions are mentioned in the text? Name them.
     (Combien de régions sont mentionnées dans le texte ? Nomme-les.)
  2. Find ONE connector of opposition in the text. Write the full sentence.
     (Trouve UN connecteur d'opposition dans le texte. Écris la phrase entière.)
  3. Find ONE connector of cause in the text. Write the full sentence.
     (Trouve UN connecteur de cause dans le texte. Écris la phrase entière.)
  4. According to the text, how do students solve misunderstandings?
     (Selon le texte, comment les élèves résolvent-ils les malentendus ?)
  5. In your own words, explain what the last sentence means.
     (Avec tes propres mots, explique ce que signifie la dernière phrase.)

  MÉTHODE
  Étape 1 · Lis le texte une première fois pour comprendre l'idée générale.
  Étape 2 · Pour les questions 2 et 3 : cherche les connecteurs que tu connais (although, however, because, so...).
  Étape 3 · Pour la question 4 : cherche dans le texte comment les conflits sont résolus.
  Étape 4 · Pour la question 5 : explique avec tes propres mots, ne recopie pas la phrase.
  ✔ Vérification : Vérifie que tes réponses sont des phrases complètes en anglais.
  ✔ Conclusion :
    1. Four regions are mentioned : the north, the south, the east and the west.
       (Quatre régions sont mentionnées : le nord, le sud, l'est et l'ouest.)
    2. Connector of opposition : "Although we are very different, we get along well because we respect each other."
       OR : "However, sometimes there are misunderstandings."
       (Connecteur d'opposition : "Bien que nous soyons très différents, nous nous entendons bien parce que nous nous respectons." OU : "Cependant, parfois il y a des malentendus.")
    3. Connector of cause : "we get along well because we respect each other."
       (Connecteur de cause : "nous nous entendons bien parce que nous nous respectons.")
    4. When there are misunderstandings, students talk and find solutions together.
       (Quand il y a des malentendus, les élèves discutent et trouvent des solutions ensemble.)
    5. The last sentence means that tolerance is not just something we say — it is something we do every day, in our actions and choices.
       (La dernière phrase signifie que la tolérance n'est pas seulement quelque chose qu'on dit — c'est quelque chose qu'on pratique chaque jour, dans nos actions et nos choix.)
  ✔ Ce que cet exercice mobilise : compréhension écrite · identification des connecteurs logiques · cause et opposition · expression personnelle · traduction anglais-français.

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulary matching · Notions travaillées : Vocabulaire de la tolérance

  Match each English word with its French translation.
  Relie chaque mot anglais à sa traduction française.

  1. Prejudice               a. Coexister
  2. To coexist              b. Pacifiquement
  3. Diversity               c. Préjugé
  4. Peacefully              d. Intimidation / harcèlement
  5. Bullying                e. Diversité

  GUIDE
  Étape 1 · Relis la trace écrite, section I (Vocabulaire clés).
  Étape 2 · Pour chaque mot anglais, cherche sa traduction dans ta liste de vocabulaire.
  Étape 3 · Relie par une flèche ou écris la lettre correspondante.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — True or False · Notions travaillées : Compréhension du thème de la tolérance

  Read the sentences. Write TRUE or FALSE. Correct the false sentences in English.
  Lis les phrases. Écris VRAI ou FAUX. Corrige les phrases fausses en anglais.

  1. Tolerance means accepting everything, including violence and injustice.
  2. Intolerance can lead to discrimination and conflict.
  3. Diversity makes a society weaker.
  4. In Côte d'Ivoire, there are more than 60 ethnic groups.
  5. Prejudice is often based on real knowledge and facts.

  GUIDE
  Étape 1 · Relis la trace écrite, sections II et III.
  Étape 2 · Pour chaque phrase, vérifie si elle correspond à ce que tu as appris.
  Étape 3 · Si FALSE : réécris la phrase correctement en anglais.
  Étape 4 · Traduis chaque phrase corrigée en français.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Connecteurs logiques · Notions travaillées : Because / So / However / Although / Therefore

  Join the two sentences using the connector given in brackets.
  Relie les deux phrases en utilisant le connecteur indiqué entre parenthèses.

  1. People are often intolerant. They do not know other cultures. (BECAUSE)
  2. We are all different. We can build a strong country together. (ALTHOUGH)
  3. Intolerance causes division. We must fight it. (THEREFORE)
  4. Tolerance is a fundamental value. It is not always easy to practise. (HOWEVER)
  5. Konan respects all his classmates. He gets along well with everyone. (SO)

  GUIDE
  Étape 1 · Identifie la relation logique exprimée par chaque connecteur.
  Étape 2 · BECAUSE : place-le entre les deux idées (cause → effet).
  Étape 3 · ALTHOUGH : place-le au début de la première proposition (opposition dans une même phrase).
  Étape 4 · THEREFORE : place-le en début de la deuxième phrase après un point-virgule ou un point.
  Étape 5 · HOWEVER : commence une nouvelle phrase ou place après un point-virgule.
  Étape 6 · SO : place-le entre les deux propositions (virgule + so).
  Étape 7 · Traduis chaque phrase complétée en français.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Forms of intolerance · Notions travaillées : Identifier et nommer les formes d'intolérance

  Read each situation and identify the form of intolerance. Then write ONE sentence using MUST NOT to condemn it.
  Lis chaque situation et identifie la forme d'intolérance. Puis écris UNE phrase avec MUST NOT pour la condamner.

  1. Adjoa refuses to sit next to Mamadou because he is Muslim.
  2. Some boys in the class laugh at a girl because she asks a lot of questions.
  3. A student mocks another student because he speaks a different language at home.
  4. A teacher gives better marks to students from his own ethnic group.
  5. Some people say that girls should not go to university.

  GUIDE
  Étape 1 · Relis la trace écrite, section III (Formes d'intolérance).
  Étape 2 · Pour chaque situation, identifie la forme d'intolérance (racism, bullying, sexism, etc.).
  Étape 3 · Écris une phrase avec MUST NOT pour condamner ce comportement.
  Étape 4 · Traduis ta phrase en français.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Short writing · Notions travaillées : Production écrite · tolérance + connecteurs

  Write FIVE sentences about tolerance in Côte d'Ivoire.
  Écris CINQ phrases sur la tolérance en Côte d'Ivoire.

  Contraintes :
  · Phrase 1 : use BECAUSE or BECAUSE OF.
  · Phrase 2 : use ALTHOUGH.
  · Phrase 3 : use HOWEVER.
  · Phrase 4 : use THEREFORE.
  · Phrase 5 : use MUST or SHOULD + a word from the vocabulary list.

  GUIDE
  Étape 1 · Relis la trace écrite, section IV (Connecteurs logiques) et section I (Vocabulaire).
  Étape 2 · Pour chaque phrase, choisis une idée sur la tolérance en Côte d'Ivoire et applique le connecteur demandé.
  Étape 3 · Vérifie : BECAUSE est suivi d'un sujet + verbe. BECAUSE OF est suivi d'un nom.
  Étape 4 · Traduis chaque phrase en français.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — Is tolerance possible in a diverse society?
  Durée conseillée : 40 minutes

  Sujet :
  Write a short argumentative text of 10 to 12 sentences in English.
  (Écris un court texte argumentatif de 10 à 12 phrases en anglais.)

  Answer the following question: "Is tolerance possible in a diverse society like Côte d'Ivoire? Why?"
  (Réponds à la question suivante : "La tolérance est-elle possible dans une société diverse comme la Côte d'Ivoire ? Pourquoi ?")

  Your text must include :
  · An introduction that presents the question and your position (agree / disagree / nuanced).
  · At least TWO arguments in favour of tolerance (with examples from Côte d'Ivoire).
  · At least ONE obstacle to tolerance (a form of intolerance that exists).
  · At least ONE solution to promote tolerance.
  · A conclusion that restates your position.

  Contraintes obligatoires :
  · 10 à 12 phrases minimum.
  · Utiliser au moins 3 connecteurs logiques différents (because / although / however / therefore / so).
  · Utiliser au moins 2 modaux différents (must / should / have to).
  · Utiliser au moins 6 mots du vocabulaire de la leçon.
  · Traduire le texte en français après l'avoir écrit en anglais.

  GUIDE (sans corrigé)
  Étape 1 · Relis la trace écrite : tolérance, intolérance, formes d'intolérance, connecteurs.
  Étape 2 · Au brouillon : note ta position (oui, la tolérance est possible), 2 arguments, 1 obstacle, 1 solution.
  Étape 3 · Introduction : "Côte d'Ivoire is a diverse country with over 60 ethnic groups. In my opinion, tolerance is possible because..."
  Étape 4 · Développe chaque argument avec un connecteur logique.
  Étape 5 · Mentionne un obstacle : "However, intolerance still exists in the form of..."
  Étape 6 · Propose une solution : "To promote tolerance, we should / we have to..."
  Étape 7 · Conclus : "Therefore, I believe that..."
  Étape 8 · Relis, corrige, puis traduis en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON : Is tolerance possible in a diverse society?

Proposition de corrigé (texte modèle en anglais) :

Côte d'Ivoire is a diverse country with more than 60 ethnic groups, several religions and many different cultures. In my opinion, tolerance is not only possible — it is absolutely necessary.

First, tolerance is possible because Ivorian people have a long history of living together peacefully. In Abidjan, Christians and Muslims live side by side in the same neighbourhoods. Although they have different beliefs, they celebrate each other's festivals and help each other in difficult times.

Second, diversity is a strength, not a weakness. Because Côte d'Ivoire brings together so many different cultures, it is a rich and creative country. Tolerance allows us to benefit from this diversity.

However, intolerance still exists. Some people discriminate against others because of their ethnic group or religion. Bullying also happens in schools, where some students mock their classmates for being different. These behaviours must not be tolerated.

To promote tolerance, we should teach its values from an early age. Schools have to include tolerance and civic education in their programmes. If we educate young people to respect differences, our society will become more peaceful and united.

Therefore, I believe that tolerance is possible in Côte d'Ivoire. It requires daily effort, but it is the only path to a just and peaceful society. As our national motto says: "Union — Discipline — Travail." Unity begins with tolerance.

Traduction française complète :
La Côte d'Ivoire est un pays divers avec plus de 60 groupes ethniques, plusieurs religions et de nombreuses cultures différentes. À mon avis, la tolérance n'est pas seulement possible — elle est absolument nécessaire.

Premièrement, la tolérance est possible parce que les Ivoiriens ont une longue histoire de vie commune pacifique. À Abidjan, les chrétiens et les musulmans vivent côte à côte dans les mêmes quartiers. Bien qu'ils aient des croyances différentes, ils célèbrent les fêtes des uns et des autres et s'entraident dans les moments difficiles.

Deuxièmement, la diversité est une force, pas une faiblesse. Parce que la Côte d'Ivoire rassemble tant de cultures différentes, c'est un pays riche et créatif. La tolérance nous permet de bénéficier de cette diversité.

Cependant, l'intolérance existe encore. Certaines personnes discriminent les autres en raison de leur groupe ethnique ou de leur religion. L'intimidation se produit aussi dans les écoles, où certains élèves se moquent de leurs camarades parce qu'ils sont différents. Ces comportements ne doivent pas être tolérés.

Pour promouvoir la tolérance, nous devrions enseigner ses valeurs dès le plus jeune âge. Les écoles doivent inclure la tolérance et l'éducation civique dans leurs programmes. Si nous éduquons les jeunes à respecter les différences, notre société deviendra plus pacifique et unie.

Par conséquent, je crois que la tolérance est possible en Côte d'Ivoire. Elle demande un effort quotidien, mais c'est le seul chemin vers une société juste et pacifique. Comme le dit notre devise nationale : "Union — Discipline — Travail." L'unité commence par la tolérance.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1 : Vocabulary matching

  1. Prejudice → c. Préjugé
  2. To coexist → a. Coexister
  3. Diversity → e. Diversité
  4. Peacefully → b. Pacifiquement
  5. Bullying → d. Intimidation / harcèlement

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2 : True or False

  1. FALSE — Tolerance does not mean accepting everything. We must not tolerate violence or injustice.
     (FAUX — La tolérance ne signifie pas tout accepter. Nous ne devons pas tolérer la violence ou l'injustice.)

  2. TRUE — Intolerance can lead to discrimination and conflict.
     (VRAI — L'intolérance peut mener à la discrimination et aux conflits.)

  3. FALSE — Diversity makes a society richer and stronger, not weaker.
     (FAUX — La diversité rend une société plus riche et plus forte, pas plus faible.)

  4. TRUE — In Côte d'Ivoire, there are more than 60 ethnic groups.
     (VRAI — En Côte d'Ivoire, il y a plus de 60 groupes ethniques.)

  5. FALSE — Prejudice is often based on ignorance and stereotypes, NOT on real knowledge and facts.
     (FAUX — Le préjugé est souvent fondé sur l'ignorance et les stéréotypes, PAS sur des connaissances et des faits réels.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3 : Connecteurs logiques

  1. BECAUSE :
     People are often intolerant because they do not know other cultures.
     (Les gens sont souvent intolérants parce qu'ils ne connaissent pas les autres cultures.)

  2. ALTHOUGH :
     Although we are all different, we can build a strong country together.
     (Bien que nous soyons tous différents, nous pouvons construire ensemble un pays fort.)

  3. THEREFORE :
     Intolerance causes division; therefore, we must fight it.
     (L'intolérance cause la division ; par conséquent, nous devons la combattre.)

  4. HOWEVER :
     Tolerance is a fundamental value. However, it is not always easy to practise.
     (La tolérance est une valeur fondamentale. Cependant, il n'est pas toujours facile de la pratiquer.)

  5. SO :
     Konan respects all his classmates, so he gets along well with everyone.
     (Konan respecte tous ses camarades, donc il s'entend bien avec tout le monde.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4 : Forms of intolerance

  1. Form : Religious intolerance (intolérance religieuse).
     Sentence : We must not refuse to sit next to someone because of their religion.
     (Nous ne devons pas refuser de nous asseoir à côté de quelqu'un à cause de sa religion.)

  2. Form : Sexism (sexisme).
     Sentence : We must not mock girls for asking questions — curiosity is a quality for everyone.
     (Nous ne devons pas nous moquer des filles qui posent des questions — la curiosité est une qualité pour tout le monde.)

  3. Form : Ethnic discrimination / bullying (discrimination ethnique / intimidation).
     Sentence : We must not mock someone because of the language they speak at home.
     (Nous ne devons pas nous moquer de quelqu'un à cause de la langue qu'il parle à la maison.)

  4. Form : Ethnic discrimination (discrimination ethnique).
     Sentence : Teachers must not favour students from their own ethnic group — all students must be treated equally.
     (Les professeurs ne doivent pas favoriser les élèves de leur propre groupe ethnique — tous les élèves doivent être traités de façon égale.)

  5. Form : Sexism (sexisme).
     Sentence : We must not say that girls should not go to university — education is a right for everyone.
     (Nous ne devons pas dire que les filles ne devraient pas aller à l'université — l'éducation est un droit pour tout le monde.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5 : Short writing

  Propositions de corrigé (d'autres réponses correctes sont possibles) :

  Phrase 1 (BECAUSE / BECAUSE OF) :
     Conflicts often arise because of intolerance and ignorance.
     (Les conflits naissent souvent à cause de l'intolérance et de l'ignorance.)

  Phrase 2 (ALTHOUGH) :
     Although Côte d'Ivoire has many different ethnic groups, its people can live together peacefully.
     (Bien que la Côte d'Ivoire ait de nombreux groupes ethniques différents, ses habitants peuvent vivre ensemble pacifiquement.)

  Phrase 3 (HOWEVER) :
     Tolerance is taught at school. However, some students still bully their classmates.
     (La tolérance est enseignée à l'école. Cependant, certains élèves intimident encore leurs camarades.)

  Phrase 4 (THEREFORE) :
     Diversity makes our country richer; therefore, we must celebrate it, not fear it.
     (La diversité rend notre pays plus riche ; par conséquent, nous devons la célébrer, pas la craindre.)

  Phrase 5 (MUST / SHOULD + vocabulaire) :
     We must fight discrimination and prejudice in all their forms.
     (Nous devons combattre la discrimination et les préjugés sous toutes leurs formes.)

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "In diversity there is beauty and there is strength."
  (Dans la diversité, il y a de la beauté et il y a de la force.)

  — Maya Angelou

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
