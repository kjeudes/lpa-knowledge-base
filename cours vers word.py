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

COURS = """✅ Anglais LV1 · Leçon 17 — My duties · en cours de génération...

---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 17 · Classe de 3ème

MY DUTIES
Mes devoirs

Compétence C6 · Écoute élaborée — Droits humains / Human Rights

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 17 — MY DUTIES
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C6 · Écoute élaborée — Droits humains / Human Rights
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 16 : My rights)
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

CAPSULE DE REPRISE — Leçon 16 : My rights (Mes droits)

5 points clés à se rappeler :
  1. Un droit est universel : il appartient à tout être humain sans discrimination.
  2. Il existe trois catégories : droits civils et politiques · droits économiques et sociaux · droits de l'enfant.
  3. Ces droits sont protégés par la DUDH (1948), la CRC (1989) et la Constitution ivoirienne (2016).
  4. La Passive Voice : Sujet + AM/IS/ARE + participe passé (+ by + auteur).
  5. MUST = obligation forte · MUST NOT = interdiction · SHOULD = conseil moral.

3 questions flash :
  Q1 : What does "to be entitled to" mean? (Que signifie "to be entitled to" ?)
  R1 : To have the right to something. (Avoir droit à quelque chose.)

  Q2 : Transform into passive: "The law protects children."
  R2 : Children are protected by the law.

  Q3 : Fill in: "Adults ___ force children to work. It is forbidden."
  R3 : must not.

Lien avec la leçon du jour :
  Dans la leçon 16, nous avons vu que chaque personne possède des droits fondamentaux. Mais avoir des droits implique aussi d'assumer des responsabilités envers les autres et envers la société. Aujourd'hui, nous allons apprendre à exprimer nos devoirs en anglais — c'est le thème de la leçon 17 : My duties (Mes devoirs).

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

Konan Hervé est élève en 3ème au Lycée Moderne de Yopougon. Chaque matin, il salue ses professeurs, range sa salle de classe et aide ses camarades en difficulté. Son père lui a toujours dit : "Un bon citoyen ne demande pas seulement ce que la société peut faire pour lui — il demande aussi ce qu'il peut faire pour la société."

En Côte d'Ivoire, la Constitution de 2016 ne liste pas seulement des droits : elle énonce aussi des devoirs pour chaque citoyen. Parmi ces devoirs : respecter la loi, payer ses impôts, protéger l'environnement, aller à l'école, et défendre la patrie.

Avoir des devoirs, c'est reconnaître qu'on vit avec d'autres personnes et qu'on leur doit quelque chose.
Having duties means recognising that we live with other people and that we owe them something.

Lien DPFC : Ce thème est lié à l'éducation civique et à la citoyenneté responsable, valeurs centrales du programme ivoirien.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — Qu'est-ce qu'un devoir ?
  Q : What is the difference between a duty and an obligation?
      (Quelle est la différence entre un devoir et une obligation ?)
  Guide :
    · Un devoir est souvent moral : on le fait parce que c'est bien, même si personne ne nous oblige.
    · Une obligation est souvent légale : on doit le faire, sinon il y a une sanction.
    · Pense à des exemples : aider un camarade (devoir moral) vs payer ses impôts (obligation légale).
  Corrigé :
    A duty is often a moral responsibility that a person chooses to fulfil because it is right and good. An obligation, on the other hand, is something required by law or authority, with consequences if not respected. For example, helping an elderly person cross the street is a moral duty. Paying taxes is a legal obligation. Both are important for a well-functioning society.
    (Un devoir est souvent une responsabilité morale qu'une personne choisit d'accomplir parce que c'est juste et bien. Une obligation, en revanche, est quelque chose imposé par la loi ou l'autorité, avec des conséquences si elle n'est pas respectée. Par exemple, aider une personne âgée à traverser la rue est un devoir moral. Payer ses impôts est une obligation légale. Les deux sont importants pour une société bien organisée.)

▶ HARKNESS 2 — Devoirs à l'école
  Q : What are the duties of a student at school?
      (Quels sont les devoirs d'un élève à l'école ?)
  Guide :
    · Pense à ce qu'un bon élève doit faire envers ses professeurs, ses camarades et l'école.
    · Utilise : attend (assister) · respect (respecter) · study (étudier) · help (aider).
    · Pense aussi à ce qu'un élève NE doit PAS faire.
  Corrigé :
    A student has several duties at school. First, they must attend classes regularly and arrive on time. Second, they must respect their teachers and classmates. Third, they should study hard and complete their homework. Fourth, they must not cheat during exams or bully other students. Finally, they should help keep the school clean and contribute to a positive learning environment.
    (Un élève a plusieurs devoirs à l'école. Premièrement, il doit assister régulièrement aux cours et arriver à l'heure. Deuxièmement, il doit respecter ses professeurs et ses camarades. Troisièmement, il devrait travailler dur et faire ses devoirs. Quatrièmement, il ne doit pas tricher aux examens ni intimider d'autres élèves. Enfin, il devrait contribuer à garder l'école propre et à maintenir un environnement d'apprentissage positif.)

▶ HARKNESS 3 — Devoirs envers la société et l'environnement
  Q : Do we have duties towards the environment? Why?
      (Avons-nous des devoirs envers l'environnement ? Pourquoi ?)
  Guide :
    · Pense à la déforestation, à la pollution et aux problèmes environnementaux en CI.
    · Pense au lien DPFC : environnement et développement durable.
    · Utilise : protect (protéger) · preserve (préserver) · avoid (éviter) · pollute (polluer).
  Corrigé :
    Yes, we have important duties towards the environment. We must protect natural resources such as forests, rivers and oceans, because they are essential for life. We must avoid throwing rubbish in the streets or in rivers. In Côte d'Ivoire, protecting the forests of the west and the lagoon of Abidjan is a shared duty. If we do not protect the environment today, future generations will suffer the consequences.
    (Oui, nous avons d'importants devoirs envers l'environnement. Nous devons protéger les ressources naturelles telles que les forêts, les rivières et les océans, car ils sont essentiels à la vie. Nous devons éviter de jeter des ordures dans les rues ou dans les rivières. En Côte d'Ivoire, protéger les forêts de l'ouest et la lagune d'Abidjan est un devoir partagé. Si nous ne protégeons pas l'environnement aujourd'hui, les générations futures en subiront les conséquences.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre de la leçon : My Duties (Mes devoirs)
Compétence : C6 — Écoute élaborée · Droits humains / Human Rights

Objectifs de la leçon :
  À la fin de cette leçon, tu seras capable de :
  · Nommer et expliquer les principaux devoirs d'un citoyen en anglais.
  · Distinguer les devoirs moraux, civiques et environnementaux.
  · Utiliser HAVE TO et DON'T HAVE TO pour exprimer l'obligation et l'absence d'obligation.
  · Renforcer l'usage de MUST / MUST NOT / SHOULD vu en leçon 16.
  · Produire un texte argumentatif sur l'importance des devoirs dans la société.

Prérequis nécessaires :
  · MUST / MUST NOT / SHOULD (Leçon 16).
  · Le Present Simple (faits généraux).
  · La Passive Voice de base (Leçon 16).
  · Le vocabulaire des droits humains (Leçon 16).

────────────────────────────────────────────────────────

TABLEAU DES STRUCTURES FONDAMENTALES — Leçon 17

| Expression / Notion               | Valeur, emploi                                         | Exemple anglais                                                | Traduction française                                                |
|----------------------------------|--------------------------------------------------------|----------------------------------------------------------------|---------------------------------------------------------------------|
| Duty / duties                    | Devoir(s) — nom                                        | We all have duties towards our community.                      | Nous avons tous des devoirs envers notre communauté.                |
| Responsibility                   | Responsabilité                                         | It is our responsibility to protect the environment.           | C'est notre responsabilité de protéger l'environnement.             |
| Have to + verbe base             | Obligation externe (règle, loi, nécessité)             | I have to attend school every day.                             | Je dois aller à l'école tous les jours.                             |
| Don't / doesn't have to          | Absence d'obligation (pas nécessaire)                  | You don't have to shout to be respected.                       | Tu n'as pas besoin de crier pour être respecté.                     |
| Must + verbe base                | Obligation forte / personnelle                         | We must obey the law.                                          | Nous devons obéir à la loi.                                         |
| Should + verbe base              | Conseil / obligation morale                            | We should help people in need.                                 | Nous devrions aider les personnes dans le besoin.                   |
| To obey                          | Obéir                                                  | Citizens must obey the law.                                    | Les citoyens doivent obéir à la loi.                                |
| To respect                       | Respecter                                              | We must respect our elders.                                    | Nous devons respecter nos aînés.                                    |
| To protect                       | Protéger                                               | We have to protect the environment.                            | Nous devons protéger l'environnement.                               |
| To contribute to                 | Contribuer à                                           | Every citizen should contribute to national development.       | Chaque citoyen devrait contribuer au développement national.        |
| To avoid + verbe-ing             | Éviter de faire quelque chose                          | We should avoid polluting rivers.                              | Nous devrions éviter de polluer les rivières.                       |
| Civic duty                       | Devoir civique                                         | Voting is a civic duty.                                        | Voter est un devoir civique.                                        |
| Moral duty                       | Devoir moral                                           | Helping the poor is a moral duty.                              | Aider les pauvres est un devoir moral.                              |

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

A. QU'EST-CE QU'UN DEVOIR ? / WHAT IS A DUTY?

A duty is a moral or legal responsibility that a person must fulfil towards others, towards society or towards the environment.
(Un devoir est une responsabilité morale ou légale qu'une personne doit accomplir envers les autres, envers la société ou envers l'environnement.)

Rights and duties always go together: you cannot have one without the other.
(Les droits et les devoirs vont toujours ensemble : on ne peut pas avoir l'un sans l'autre.)

┌─ À RETENIR ──────────────────────────────────────────┐
  Duty = devoir (responsabilité morale ou légale)
  Responsibility = responsabilité
  Citizen = citoyen
  Civic = civique (relatif à la vie en société)
  Rights ↔ Duties : les deux faces d'une même médaille
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

B. LES TYPES DE DEVOIRS / TYPES OF DUTIES

1. DUTIES AT SCHOOL (Devoirs à l'école) :
  · Attend classes regularly. (Assister régulièrement aux cours.)
  · Arrive on time. (Arriver à l'heure.)
  · Respect teachers and classmates. (Respecter les professeurs et les camarades.)
  · Complete homework and study hard. (Faire ses devoirs et travailler dur.)
  · Not cheat during exams. (Ne pas tricher aux examens.)
  · Keep the classroom clean. (Garder la salle de classe propre.)

2. DUTIES AS A CITIZEN (Devoirs en tant que citoyen) :
  · Obey the law. (Obéir à la loi.)
  · Pay taxes. (Payer ses impôts.)
  · Vote in elections. (Voter aux élections.)
  · Respect the national flag and anthem. (Respecter le drapeau et l'hymne national.)
  · Help maintain peace and order. (Contribuer à maintenir la paix et l'ordre.)

3. DUTIES TOWARDS OTHERS (Devoirs envers les autres) :
  · Help people in need. (Aider les personnes dans le besoin.)
  · Respect elders. (Respecter les aînés.)
  · Not discriminate against anyone. (Ne pas discriminer qui que ce soit.)
  · Share resources fairly. (Partager équitablement les ressources.)

4. DUTIES TOWARDS THE ENVIRONMENT (Devoirs envers l'environnement) :
  · Protect forests and natural resources. (Protéger les forêts et les ressources naturelles.)
  · Avoid throwing rubbish in the streets. (Éviter de jeter des ordures dans les rues.)
  · Save water and electricity. (Économiser l'eau et l'électricité.)
  · Plant trees. (Planter des arbres.)

┌─ À RETENIR ──────────────────────────────────────────┐
  Il existe 4 domaines principaux de devoirs :
  1. À l'école / At school
  2. En tant que citoyen / As a citizen
  3. Envers les autres / Towards others
  4. Envers l'environnement / Towards the environment
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

C. GRAMMAIRE — HAVE TO / DON'T HAVE TO

HAVE TO exprime une obligation externe : une règle, une loi, une nécessité imposée de l'extérieur.
C'est différent de MUST, qui exprime souvent une conviction personnelle ou une obligation forte.

FORMULE :
  Sujet + HAVE TO / HAS TO + verbe base (obligation)
  Sujet + DON'T / DOESN'T HAVE TO + verbe base (pas d'obligation)

Exemples — HAVE TO (obligation) :
  · I have to go to school every day. It is the law.
    (Je dois aller à l'école tous les jours. C'est la loi.)
  · She has to wear her school uniform. It is the school rule.
    (Elle doit porter son uniforme scolaire. C'est le règlement de l'école.)
  · We have to pay taxes. It is our civic duty.
    (Nous devons payer des impôts. C'est notre devoir civique.)

Exemples — DON'T / DOESN'T HAVE TO (pas d'obligation — mais ce n'est pas interdit) :
  · You don't have to shout to be respected.
    (Tu n'es pas obligé de crier pour être respecté.)
  · He doesn't have to join the army, but he can if he wants.
    (Il n'est pas obligé de rejoindre l'armée, mais il peut s'il le souhaite.)

⚠ ATTENTION — Différence cruciale :
  MUST NOT = interdiction absolue (il ne faut PAS — c'est interdit)
  DON'T HAVE TO = absence d'obligation (ce n'est PAS nécessaire — mais c'est permis)

  Exemple :
  · You must not cheat. → C'est interdit. (Tricher est interdit.)
  · You don't have to study on Sunday. → Ce n'est pas obligatoire, mais tu peux.
    (Tu n'es pas obligé d'étudier le dimanche — mais tu peux si tu veux.)

⚠ ATTENTION — Faux ami !
  "Sensible" en anglais ≠ "Sensible" en français.
  · Sensible (EN) = raisonnable, plein de bon sens
  · Sensible (FR) = sensitive (EN) = sensible aux émotions
  Exemple : "It is sensible to obey the law."
  = "Il est raisonnable d'obéir à la loi." (PAS "il est sensible")

────────────────────────────────────────────────────────

D. DROITS ET DEVOIRS — LES DEUX FACES D'UNE MÊME MÉDAILLE
   RIGHTS AND DUTIES — TWO SIDES OF THE SAME COIN

Chaque droit correspond à un devoir :

  · Right to education → Duty to attend school and study seriously.
    (Droit à l'éducation → Devoir d'aller à l'école et de travailler sérieusement.)

  · Right to freedom of expression → Duty to respect others' opinions.
    (Droit à la liberté d'expression → Devoir de respecter les opinions des autres.)

  · Right to live in a clean environment → Duty to protect and preserve nature.
    (Droit à vivre dans un environnement sain → Devoir de protéger et préserver la nature.)

  · Right to security → Duty to obey the law and respect others.
    (Droit à la sécurité → Devoir d'obéir à la loi et de respecter autrui.)

  · Right to healthcare → Duty to contribute to the national health system (taxes).
    (Droit à la santé → Devoir de contribuer au système de santé national par les impôts.)

┌─ À RETENIR ──────────────────────────────────────────┐
  Une société fonctionne bien quand chacun respecte
  à la fois ses droits ET ses devoirs.
  Rights without duties = selfishness (égoïsme)
  Duties without rights = injustice
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Les 4 domaines de devoirs]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Nœud central (cercle orange) : "MY DUTIES / MES DEVOIRS"
  · Branche 1 (haut gauche, fond bleu) : "AT SCHOOL / À L'ÉCOLE"
      Sous-éléments : attend class · respect teachers · no cheating · study hard
  · Branche 2 (haut droite, fond vert) : "AS A CITIZEN / EN TANT QUE CITOYEN"
      Sous-éléments : obey the law · pay taxes · vote · respect flag & anthem
  · Branche 3 (bas gauche, fond violet) : "TOWARDS OTHERS / ENVERS LES AUTRES"
      Sous-éléments : help the needy · respect elders · no discrimination
  · Branche 4 (bas droite, fond marron clair) : "TOWARDS THE ENVIRONMENT / ENVERS L'ENVIRONNEMENT"
      Sous-éléments : protect forests · no littering · save water · plant trees
  · Flèches : de chaque branche vers le nœud central
  · Légende couleur : bleu = école | vert = civique | violet = social | marron = environnement
Note générateur : Canva ou MindMeister — disposition en croix autour du centre
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Tableau comparatif : MUST NOT vs DON'T HAVE TO]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Deux colonnes côte à côte, séparées par une ligne rouge épaisse
  · Colonne gauche (fond rouge pâle) : "MUST NOT / NE DOIT PAS"
      Sens : INTERDIT — c'est défendu
      Exemple 1 : You must not cheat during exams.
                  (Tu ne dois pas tricher aux examens.)
      Exemple 2 : We must not pollute rivers.
                  (Nous ne devons pas polluer les rivières.)
      Icône suggérée : panneau STOP ou croix rouge
  · Colonne droite (fond vert pâle) : "DON'T HAVE TO / N'EST PAS OBLIGÉ"
      Sens : PAS NÉCESSAIRE — mais permis
      Exemple 1 : You don't have to study on Sunday.
                  (Tu n'es pas obligé d'étudier le dimanche.)
      Exemple 2 : He doesn't have to join the army.
                  (Il n'est pas obligé de rejoindre l'armée.)
      Icône suggérée : panneau d'information ou point d'interrogation
  · Encadré bas : "ATTENTION : MUST NOT ≠ DON'T HAVE TO — erreur fréquente au BEPC !"
Note générateur : tableau Word, draw.io ou Canva — 2 colonnes contrastées rouge/vert
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Duty / Un devoir
  A moral or legal responsibility that a person must fulfil towards others, the community or the environment.
  (Une responsabilité morale ou légale qu'une personne doit accomplir envers les autres, la communauté ou l'environnement.)

◆ DÉFINITION DPFC — Civic duty / Devoir civique
  A responsibility that citizens have towards their country and society, such as obeying the law, paying taxes and voting.
  (Une responsabilité que les citoyens ont envers leur pays et leur société, comme obéir à la loi, payer des impôts et voter.)

◆ DÉFINITION DPFC — Moral duty / Devoir moral
  A responsibility based on ethical values and personal conscience, not necessarily imposed by law.
  (Une responsabilité fondée sur des valeurs éthiques et la conscience personnelle, pas nécessairement imposée par la loi.)

◆ DÉFINITION DPFC — Have to / Devoir (obligation externe)
  A modal expression used to describe an external obligation imposed by rules, laws or circumstances: Subject + have to / has to + base verb.
  (Une expression modale utilisée pour décrire une obligation externe imposée par des règles, des lois ou des circonstances : Sujet + have to / has to + verbe base.)

◆ DÉFINITION DPFC — Don't have to / Pas obligé de
  A modal expression indicating that something is not necessary or required, but is not forbidden either: Subject + don't / doesn't have to + base verb.
  (Une expression modale indiquant que quelque chose n'est pas nécessaire ou exigé, mais n't est pas non plus interdit : Sujet + don't / doesn't have to + verbe base.)

◆ DÉFINITION DPFC — Responsibility / Responsabilité
  The state of being accountable for one's actions and their consequences towards others and society.
  (L'état d'être responsable de ses actes et de leurs conséquences envers les autres et la société.)

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

I. VOCABULAIRE CLÉS — MY DUTIES / MES DEVOIRS

  Duty / duties = devoir(s)
  Responsibility = responsabilité
  Citizen = citoyen
  Civic duty = devoir civique
  Moral duty = devoir moral
  To obey = obéir
  To respect = respecter
  To protect = protéger
  To contribute to = contribuer à
  To avoid + verbe-ing = éviter de faire
  To pollute = polluer
  To preserve = préserver
  To plant = planter
  Rubbish = ordures / déchets
  Taxes = impôts
  Election = élection
  To vote = voter
  Elders = aînés / personnes âgées
  Sensible (EN) = raisonnable (≠ sensible en français !)

II. LES 4 DOMAINES DE DEVOIRS / THE 4 AREAS OF DUTIES

  1. At school (À l'école) :
     attend class · arrive on time · respect teachers · study hard · no cheating

  2. As a citizen (En tant que citoyen) :
     obey the law · pay taxes · vote · respect the flag and anthem

  3. Towards others (Envers les autres) :
     help the needy · respect elders · no discrimination · share fairly

  4. Towards the environment (Envers l'environnement) :
     protect forests · avoid littering · save water · plant trees

III. GRAMMAIRE — HAVE TO / DON'T HAVE TO

  HAVE TO / HAS TO + verbe base = obligation externe (règle, loi)
    I have to go to school. (Je dois aller à l'école — c'est la loi.)
    She has to wear her uniform. (Elle doit porter son uniforme — règlement.)

  DON'T / DOESN'T HAVE TO + verbe base = pas d'obligation (mais permis)
    You don't have to study on Sunday. (Tu n'es pas obligé d'étudier le dimanche.)
    He doesn't have to join the army. (Il n'est pas obligé de rejoindre l'armée.)

IV. DIFFÉRENCE ESSENTIELLE — MUST NOT vs DON'T HAVE TO

  MUST NOT = INTERDIT (il ne faut pas — c'est défendu)
    You must not cheat. (Tricher est interdit.)

  DON'T HAVE TO = PAS NÉCESSAIRE (ce n'est pas obligatoire — mais ce n'est pas interdit)
    You don't have to study on Sunday. (Ce n'est pas obligatoire — mais tu peux.)

V. DROITS ET DEVOIRS — DEUX FACES D'UNE MÊME MÉDAILLE

  Right to education → Duty to attend school and study.
  Right to free expression → Duty to respect others' opinions.
  Right to a clean environment → Duty to protect nature.
  Right to security → Duty to obey the law.

VI. FAUX AMI À RETENIR

  Sensible (EN) ≠ Sensible (FR)
  Sensible (EN) = raisonnable / plein de bon sens
  Sensible (FR) = sensitive (EN)
  Exemple : "It is sensible to obey the law." = "Il est raisonnable d'obéir à la loi."

────────────────────────────────────────────────────────

~ ANALOGIE CI 1
  HAVE TO, c'est comme le règlement intérieur du lycée affiché sur le mur de la direction à Bouaké : tu dois respecter les heures de cours, porter ton uniforme, remettre tes devoirs à temps. Ce n'est pas toi qui as décidé ces règles — elles s'imposent à toi de l'extérieur. C'est exactement ce qu'exprime HAVE TO.

~ ANALOGIE CI 2
  La différence entre MUST NOT et DON'T HAVE TO, c'est comme la différence entre "Il est interdit de traverser le Pont Houphouët-Boigny à pied" (MUST NOT — c'est dangereux et défendu) et "Tu n'es pas obligé de prendre le gbaka pour aller au marché Adjamé" (DON'T HAVE TO — tu peux y aller à pied si tu veux). L'un interdit, l'autre laisse le choix.

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. Un devoir est une responsabilité morale ou légale envers les autres, la société et l'environnement.
  2. Il existe 4 domaines de devoirs : à l'école · en tant que citoyen · envers les autres · envers l'environnement.
  3. HAVE TO / HAS TO = obligation externe imposée par une règle ou une loi.
  4. DON'T / DOESN'T HAVE TO = absence d'obligation (ce n'est pas nécessaire, mais c'est permis).
  5. MUST NOT ≠ DON'T HAVE TO : MUST NOT = interdit · DON'T HAVE TO = pas nécessaire.

5 mots clés + définition courte :
  · Duty : responsabilité morale ou légale envers les autres ou la société.
  · Civic duty : devoir du citoyen envers son pays (voter, payer ses impôts, obéir à la loi).
  · Have to : obligation externe imposée par une règle ou une loi.
  · Don't have to : absence d'obligation — ce n'est pas nécessaire mais c'est permis.
  · Responsibility : état d'être comptable de ses actes et de leurs conséquences.

3 erreurs fréquentes + correction :
  Erreur 1 : Confondre MUST NOT et DON'T HAVE TO.
  ✔ Correction : MUST NOT = interdit (You must not steal). DON'T HAVE TO = pas nécessaire (You don't have to come early — but you can).

  Erreur 2 : Oublier le -S à HAS TO à la 3ème personne du singulier.
  ✔ Correction : He / She / It HAS TO (pas "have to"). Exemple : "She has to wear her uniform." (PAS "She have to...")

  Erreur 3 : Traduire "sensible" par "sensible".
  ✔ Correction : Sensible (EN) = raisonnable. Sensitive (EN) = sensible (FR). Exemple : "A sensible citizen obeys the law." = "Un citoyen raisonnable obéit à la loi."

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Have to / Don't have to · Notion ciblée : Obligation externe et absence d'obligation

  Énoncé :
  Complete the sentences with HAVE TO, HAS TO, DON'T HAVE TO or DOESN'T HAVE TO.
  Complète les phrases avec HAVE TO, HAS TO, DON'T HAVE TO ou DOESN'T HAVE TO.

  1. Every citizen ___ obey the law. It is not optional.
  2. You ___ shout to make yourself heard. Speak calmly.
  3. Kofi ___ wear his school uniform every day. It is the school rule.
  4. Students ___ cheat. There are other ways to succeed.
     (Attention : cette phrase exprime une interdiction — utilise le bon modal !)
  5. In Côte d'Ivoire, adults ___ vote. It is both a right and a civic duty.

  MÉTHODE
  Étape 1 · Lis chaque phrase et identifie si elle exprime une obligation, une absence d'obligation ou une interdiction.
  Étape 2 · Obligation externe (règle/loi) → HAVE TO / HAS TO (selon le sujet).
  Étape 3 · Pas d'obligation — mais permis → DON'T HAVE TO / DOESN'T HAVE TO.
  Étape 4 · Interdiction absolue → MUST NOT (attention, phrase 4 est un piège !).
  Étape 5 · Vérifie : HAS TO = he/she/it · HAVE TO = I/you/we/they.
  ✔ Vérification : Relis chaque phrase complétée. Le sens est-il logique ?
  ✔ Conclusion :
    1. have to — Every citizen has to obey the law.
       (Chaque citoyen doit obéir à la loi.)
    2. don't have to — You don't have to shout to make yourself heard.
       (Tu n'as pas besoin de crier pour te faire entendre.)
    3. has to — Kofi has to wear his school uniform every day.
       (Kofi doit porter son uniforme scolaire tous les jours.)
    4. must not — Students must not cheat. (Piège : c'est une interdiction, pas une absence d'obligation.)
       (Les élèves ne doivent pas tricher.)
    5. have to — Adults have to vote. (Le sujet "adults" est pluriel → have to.)
       (Les adultes doivent voter.)
  ✔ Ce que cet exercice mobilise : HAVE TO / HAS TO · DON'T HAVE TO · MUST NOT · accord du sujet · distinction obligation / interdiction / absence d'obligation · traduction anglais-français.

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Rights and duties · Notion ciblée : Lien entre droits et devoirs

  Énoncé :
  Match each right with its corresponding duty. Then write a complete sentence linking them.
  Relie chaque droit à son devoir correspondant. Puis écris une phrase complète les reliant.

  Rights :
    1. The right to education
    2. The right to live in a clean environment
    3. The right to freedom of expression

  Duties :
    a. We have to protect and preserve nature.
    b. We have to attend school and study seriously.
    c. We must respect other people's opinions.

  MÉTHODE
  Étape 1 · Lis chaque droit et demande-toi : quelle responsabilité implique-t-il ?
  Étape 2 · Relie le droit au devoir correspondant.
  Étape 3 · Écris une phrase complète qui relie les deux avec "but" ou "so" ou "therefore".
  Étape 4 · Traduis chaque phrase en français.
  ✔ Vérification : Vérifie que ta phrase contient bien le droit ET le devoir.
  ✔ Conclusion :
    1 → b : We have the right to education, so we have to attend school and study seriously.
            (Nous avons le droit à l'éducation, donc nous devons aller à l'école et travailler sérieusement.)
    2 → a : We have the right to live in a clean environment, but we also have to protect and preserve nature.
            (Nous avons le droit de vivre dans un environnement sain, mais nous devons aussi protéger et préserver la nature.)
    3 → c : We have the right to freedom of expression, but we must respect other people's opinions.
            (Nous avons le droit à la liberté d'expression, mais nous devons respecter les opinions des autres.)
  ✔ Ce que cet exercice mobilise : lien droits-devoirs · HAVE TO / MUST · connecteurs logiques (so, but, therefore) · production de phrases complexes · traduction anglais-français.

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulary matching · Notions travaillées : Vocabulaire des devoirs

  Match each English word with its French translation.
  Relie chaque mot anglais à sa traduction française.

  1. To obey                   a. Préserver
  2. Rubbish                   b. Voter
  3. To preserve               c. Ordures / déchets
  4. To vote                   d. Aîné(s)
  5. Elders                    e. Obéir

  GUIDE
  Étape 1 · Relis la trace écrite, section I (Vocabulaire clés).
  Étape 2 · Pour chaque mot anglais, cherche sa traduction dans ta liste de vocabulaire.
  Étape 3 · Relie par une flèche ou écris la lettre correspondante.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Must not or Don't have to? · Notions travaillées : Distinction MUST NOT / DON'T HAVE TO

  Choose the correct option: MUST NOT or DON'T HAVE TO.
  Choisis la bonne option : MUST NOT ou DON'T HAVE TO.

  1. You ___ throw rubbish in the river. It pollutes the water.
  2. Students ___ wear their school uniform on Saturdays. The school is closed.
  3. Citizens ___ break the law. There are consequences.
  4. You ___ be rich to be a good citizen. Everyone can contribute.
  5. We ___ discriminate against anyone. It is a violation of human rights.

  GUIDE
  Étape 1 · Pour chaque phrase, demande-toi : est-ce que c'est INTERDIT ou simplement PAS NÉCESSAIRE ?
  Étape 2 · Interdit → MUST NOT.
  Étape 3 · Pas nécessaire (mais permis) → DON'T HAVE TO.
  Étape 4 · Traduis chaque phrase complétée en français.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Classifying duties · Notions travaillées : Les 4 domaines de devoirs

  Classify the following duties into the correct category.
  Classe les devoirs suivants dans la bonne catégorie.

  Duties to classify :
    respect teachers · plant trees · pay taxes · help the needy ·
    save water · vote in elections · arrive on time · not discriminate against anyone

  Categories :
    A. At school (À l'école)
    B. As a citizen (En tant que citoyen)
    C. Towards others (Envers les autres)
    D. Towards the environment (Envers l'environnement)

  GUIDE
  Étape 1 · Relis la trace écrite, section II (Les 4 domaines de devoirs).
  Étape 2 · Pour chaque devoir, demande-toi : dans quel contexte s'applique-t-il ?
  Étape 3 · Note la lettre correspondante (A, B, C ou D) devant chaque devoir.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Have to / Has to · Notions travaillées : Obligation externe + accord du sujet

  Rewrite the sentences using HAVE TO or HAS TO.
  Réécris les phrases en utilisant HAVE TO ou HAS TO.

  1. It is obligatory for students to attend all classes.
  2. It is compulsory for Awa to submit her homework on time.
  3. It is necessary for all citizens to respect the national flag.
  4. It is the rule for Konan to wear his uniform every day.
  5. It is important for us to protect the environment.

  GUIDE
  Étape 1 · Identifie le sujet de chaque phrase.
  Étape 2 · Si le sujet est he/she/it ou un nom singulier → HAS TO.
  Étape 3 · Si le sujet est I/you/we/they ou un nom pluriel → HAVE TO.
  Étape 4 · Écris la phrase complète : Sujet + HAVE TO / HAS TO + verbe base.
  Étape 5 · Traduis chaque phrase en français.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Short reading comprehension · Notions travaillées : Compréhension écrite + devoirs + modaux

  Read the text and answer the questions in English.
  Lis le texte et réponds aux questions en anglais.

  TEXT:
  "Aminata is 15 years old and lives in San Pedro. Every morning, she wakes up early, does her household chores and then goes to school. She never misses a class. She says: 'I have to go to school because my future depends on it. I also have to help my mother at home. But I don't have to do everything alone — my brothers help too. As citizens, we must obey the law and respect our neighbours. We should also avoid throwing rubbish in the sea, because San Pedro is a port city and the sea is our greatest treasure.'"

  Questions :
  1. Where does Aminata live?
  2. What are TWO duties Aminata mentions at home and at school?
  3. Find ONE sentence with HAVE TO in the text. Write it and translate it.
  4. Find ONE sentence with DON'T HAVE TO. Write it and translate it.
  5. What environmental duty does Aminata mention? Why is it important?

  GUIDE
  Étape 1 · Lis le texte une première fois pour comprendre la situation générale.
  Étape 2 · Pour les questions 3 et 4 : cherche les expressions HAVE TO et DON'T HAVE TO dans le texte.
  Étape 3 · Pour la question 5 : cherche la phrase sur l'environnement et explique pourquoi ce devoir est important selon Aminata.
  Étape 4 · Réponds en anglais avec des phrases complètes : sujet + verbe + complément.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — A responsible citizen
  Durée conseillée : 35 minutes

  Sujet :
  Write a short text of 8 to 10 sentences in English.
  (Écris un court texte de 8 à 10 phrases en anglais.)

  Imagine you are giving a short speech at your school in Abidjan during Citizenship Day (La Journée de la Citoyenneté). You want to convince your classmates to be responsible citizens.

  Your speech must include :
  · A greeting and introduction ("Dear classmates, today I want to talk about...")
  · At least TWO duties at school (use HAVE TO or MUST)
  · At least ONE duty as a citizen (use HAVE TO or MUST)
  · At least ONE duty towards the environment (use SHOULD or HAVE TO)
  · ONE sentence using DON'T HAVE TO to show that being a good citizen is not complicated
  · A conclusion that links rights and duties

  Contraintes obligatoires :
  · 8 à 10 phrases minimum.
  · Utiliser HAVE TO, MUST, SHOULD et DON'T HAVE TO (au moins une fois chacun).
  · Utiliser au moins 5 mots du vocabulaire de la leçon.
  · Écrire à la première personne (I / we).
  · Traduire le discours en français après l'avoir écrit en anglais.

  GUIDE (sans corrigé)
  Étape 1 · Relis la trace écrite : les 4 domaines de devoirs + HAVE TO / DON'T HAVE TO / MUST / SHOULD.
  Étape 2 · Au brouillon, note 2 devoirs à l'école, 1 devoir civique, 1 devoir environnemental et 1 exemple de DON'T HAVE TO.
  Étape 3 · Commence par une salutation : "Dear classmates, today I want to talk about our duties as citizens."
  Étape 4 · Développe tes idées phrase par phrase avec les bons modaux.
  Étape 5 · Conclus avec un lien droits-devoirs : "If we respect our duties, everyone will enjoy their rights."
  Étape 6 · Relis et corrige : accord de HAS TO / HAVE TO · verbe base après les modaux.
  Étape 7 · Traduis ton discours en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON : A responsible citizen

Proposition de corrigé (discours modèle en anglais) :

Dear classmates, today I want to talk about our duties as responsible citizens.

First, we have to attend all our classes and arrive on time. A student who misses class cannot learn and cannot build a better future. We also have to respect our teachers and our classmates, because a good school is a place of mutual respect.

As citizens, we must obey the law and respect the rules of our community. In Côte d'Ivoire, every adult must vote, because our vote shapes the future of our country.

We should also take care of our environment. We should avoid throwing rubbish in the streets or in the lagoon. Abidjan is a beautiful city — we have to keep it clean.

But being a good citizen doesn't have to be difficult. Small actions matter: greeting your neighbour, helping an elderly person, picking up a piece of rubbish from the ground.

In conclusion, rights and duties go together. If we all fulfil our duties, everyone will be able to enjoy their rights. Let us be the change we want to see in Côte d'Ivoire.

Thank you.

Traduction française complète :
Chers camarades, aujourd'hui je veux vous parler de nos devoirs en tant que citoyens responsables.

Premièrement, nous devons assister à tous nos cours et arriver à l'heure. Un élève qui manque les cours ne peut pas apprendre et ne peut pas construire un meilleur avenir. Nous devons aussi respecter nos professeurs et nos camarades, parce qu'une bonne école est un lieu de respect mutuel.

En tant que citoyens, nous devons obéir à la loi et respecter les règles de notre communauté. En Côte d'Ivoire, chaque adulte doit voter, parce que notre vote façonne l'avenir de notre pays.

Nous devrions aussi prendre soin de notre environnement. Nous devrions éviter de jeter des ordures dans les rues ou dans la lagune. Abidjan est une belle ville — nous devons la garder propre.

Mais être un bon citoyen n'a pas besoin d'être difficile. Les petits gestes comptent : saluer son voisin, aider une personne âgée, ramasser un déchet par terre.

En conclusion, droits et devoirs vont ensemble. Si nous accomplissons tous nos devoirs, chacun pourra jouir de ses droits. Soyons le changement que nous voulons voir en Côte d'Ivoire.

Merci.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1 : Vocabulary matching

  1. To obey → e. Obéir
  2. Rubbish → c. Ordures / déchets
  3. To preserve → a. Préserver
  4. To vote → b. Voter
  5. Elders → d. Aîné(s)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2 : Must not or Don't have to?

  1. MUST NOT — You must not throw rubbish in the river. It pollutes the water.
     (Tu ne dois pas jeter des ordures dans la rivière. Cela pollue l'eau.) → C'est interdit.

  2. DON'T HAVE TO — Students don't have to wear their school uniform on Saturdays.
     (Les élèves ne sont pas obligés de porter leur uniforme le samedi.) → Pas nécessaire.

  3. MUST NOT — Citizens must not break the law. There are consequences.
     (Les citoyens ne doivent pas enfreindre la loi. Il y a des conséquences.) → C'est interdit.

  4. DON'T HAVE TO — You don't have to be rich to be a good citizen.
     (Tu n'as pas besoin d'être riche pour être un bon citoyen.) → Pas nécessaire.

  5. MUST NOT — We must not discriminate against anyone.
     (Nous ne devons pas discriminer qui que ce soit.) → C'est interdit.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3 : Classifying duties

  A. At school (À l'école) :
     · Respect teachers (respecter les professeurs)
     · Arrive on time (arriver à l'heure)

  B. As a citizen (En tant que citoyen) :
     · Pay taxes (payer ses impôts)
     · Vote in elections (voter aux élections)

  C. Towards others (Envers les autres) :
     · Help the needy (aider les personnes dans le besoin)
     · Not discriminate against anyone (ne pas discriminer qui que ce soit)

  D. Towards the environment (Envers l'environnement) :
     · Plant trees (planter des arbres)
     · Save water (économiser l'eau)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4 : Have to / Has to

  1. Students have to attend all classes.
     (Les élèves doivent assister à tous les cours.)
     → Sujet pluriel "students" → HAVE TO.

  2. Awa has to submit her homework on time.
     (Awa doit remettre ses devoirs à temps.)
     → Sujet singulier "Awa" → HAS TO.

  3. All citizens have to respect the national flag.
     (Tous les citoyens doivent respecter le drapeau national.)
     → Sujet pluriel "all citizens" → HAVE TO.

  4. Konan has to wear his uniform every day.
     (Konan doit porter son uniforme tous les jours.)
     → Sujet singulier "Konan" → HAS TO.

  5. We have to protect the environment.
     (Nous devons protéger l'environnement.)
     → Sujet "we" → HAVE TO.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5 : Short reading comprehension

  1. Aminata lives in San Pedro.
     (Aminata vit à San Pedro.)

  2. Two duties mentioned :
     · At home : she has to help her mother / do household chores.
       (À la maison : elle doit aider sa mère / faire les tâches ménagères.)
     · At school : she has to go to school / she never misses a class.
       (À l'école : elle doit aller à l'école / elle ne manque jamais un cours.)

  3. Sentence with HAVE TO :
     "I have to go to school because my future depends on it."
     (Je dois aller à l'école parce que mon avenir en dépend.)
     OU : "I also have to help my mother at home."
     (Je dois aussi aider ma mère à la maison.)

  4. Sentence with DON'T HAVE TO :
     "I don't have to do everything alone — my brothers help too."
     (Je ne suis pas obligée de tout faire seule — mes frères aident aussi.)

  5. Environmental duty : Aminata says we should avoid throwing rubbish in the sea.
     It is important because San Pedro is a port city and the sea is their greatest treasure — it provides food, income and beauty to the community.
     (Devoir environnemental : Aminata dit qu'on devrait éviter de jeter des ordures dans la mer. C'est important parce que San Pedro est une ville portuaire et que la mer est leur plus grand trésor — elle fournit de la nourriture, des revenus et de la beauté à la communauté.)

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "Ask not what your country can do for you — ask what you can do for your country."
  (Ne demande pas ce que ton pays peut faire pour toi — demande ce que tu peux faire pour ton pays.)

  — John F. Kennedy, Discours inaugural, 1961

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
