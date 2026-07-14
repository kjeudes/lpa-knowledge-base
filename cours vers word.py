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

COURS = """✅ Anglais LV1 · Leçon 21 — HIV/AIDS · en cours de génération...

---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 21 · Classe de 3ème

HIV/AIDS
*Le VIH/SIDA*

Compétence C7 · Hygiène et Santé / Health

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 21 — HIV/AIDS
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C7 · Hygiène et Santé / Health
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 20 : Hygiene at school)
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

CAPSULE DE REPRISE — Leçon 20 : Hygiene at school (L'hygiène à l'école)

5 points clés à retenir :
  1. L'hygiène à l'école concerne tous les élèves, enseignants et le personnel.
  2. MUST exprime une obligation absolue : "Students must wash their hands."
  3. MUST NOT exprime une interdiction stricte : "You must not litter."
  4. SHOULD exprime un conseil : "You should keep your uniform clean."
  5. Must et should sont invariables — jamais de -s à la 3ème personne.

3 questions flash :
  Q1 : What does MUST express?
  R1 : Must expresses a strong obligation or a strict rule.
       (Must exprime une obligation forte ou une règle stricte.)

  Q2 : Give one hygiene rule using MUST NOT.
  R2 : You must not litter in the classroom.
       (Tu ne dois pas jeter des ordures en classe.)

  Q3 : What is the difference between MUST and SHOULD?
  R3 : Must is for strict obligations ; should is for advice and recommendations.
       (Must est pour les obligations strictes ; should est pour les conseils.)

Lien avec la leçon du jour :
  La leçon 20 nous a montré comment l'hygiène prévient les maladies à l'école.
  Aujourd'hui, nous allons étudier une maladie très grave qui touche de nombreuses
  personnes en Côte d'Ivoire et dans le monde : le VIH/SIDA.
  Nous allons apprendre à en parler en anglais, à comprendre ses modes de transmission
  et à connaître les moyens de prévention.

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

Adjoa est une élève de 3ème au Lycée Moderne de Bouaké.
Un jour, son professeur de SVT invite une infirmière, Mme Koffi Ama,
à venir parler du VIH/SIDA devant la classe.
Mme Koffi Ama explique que la Côte d'Ivoire est l'un des pays d'Afrique
de l'Ouest les plus touchés par cette maladie.
Elle dit que beaucoup de jeunes ne connaissent pas les modes de transmission
et prennent des risques sans le savoir.
Elle distribue des brochures de l'organisation ONUSIDA et encourage les élèves
à se faire dépister au Centre de Santé d'Abobo.

→ Qu'est-ce que le VIH ? Qu'est-ce que le SIDA ?
→ Comment se transmet cette maladie ?
→ Comment peut-on se protéger ?
→ C'est exactement ce que nous allons apprendre à expliquer en anglais aujourd'hui.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — What is HIV/AIDS?
  Q : What do you know about HIV/AIDS?
      (Que sais-tu sur le VIH/SIDA ?)
  Guide :
    · Pense à ce que signifient les lettres HIV et AIDS.
    · Pense à la différence entre être séropositif (HIV+) et avoir le SIDA (AIDS).
    · Pense aux personnes touchées dans le monde et en Côte d'Ivoire.
  Corrigé :
    HIV stands for Human Immunodeficiency Virus. It is a virus that attacks
    the immune system, which is the body's defence system against diseases.
    AIDS stands for Acquired Immunodeficiency Syndrome. AIDS is the most
    advanced stage of HIV infection, when the body can no longer fight
    even simple infections. A person can live with HIV for many years
    without developing AIDS if they receive proper medical treatment.
    In Côte d'Ivoire, HIV/AIDS remains a major public health challenge,
    especially among young people.
    (HIV signifie Virus de l'Immunodéficience Humaine. C'est un virus qui attaque
    le système immunitaire, qui est le système de défense du corps contre les maladies.
    AIDS signifie Syndrome d'Immunodéficience Acquise. Le SIDA est le stade le plus
    avancé de l'infection par le VIH, quand le corps ne peut plus combattre même
    les infections simples. Une personne peut vivre avec le VIH pendant de nombreuses
    années sans développer le SIDA si elle reçoit un traitement médical approprié.
    En Côte d'Ivoire, le VIH/SIDA reste un défi majeur de santé publique,
    surtout chez les jeunes.)

▶ HARKNESS 2 — Transmission and prevention
  Q : How is HIV transmitted and how can we prevent it?
      (Comment le VIH se transmet-il et comment peut-on le prévenir ?)
  Guide :
    · Pense aux voies de transmission : sang, rapports sexuels, mère à enfant.
    · Pense à ce qui NE transmet PAS le VIH : poignée de main, partage de repas.
    · Pense aux moyens de prévention : préservatif, dépistage, fidélité.
  Corrigé :
    HIV is transmitted through three main routes : unprotected sexual intercourse,
    contact with infected blood (through shared needles or contaminated blood
    transfusions), and from an infected mother to her child during pregnancy,
    childbirth or breastfeeding.
    HIV is NOT transmitted through everyday contact such as handshakes, hugging,
    sharing food, using the same toilet or mosquito bites.
    Prevention includes using condoms, getting tested regularly, being faithful
    to one partner, avoiding sharing needles, and following medical advice
    during pregnancy.
    (Le VIH se transmet par trois voies principales : les rapports sexuels non
    protégés, le contact avec du sang infecté (par des aiguilles partagées ou
    des transfusions sanguines contaminées), et d'une mère infectée à son enfant
    pendant la grossesse, l'accouchement ou l'allaitement.
    Le VIH ne se transmet PAS par les contacts quotidiens comme les poignées de main,
    les câlins, le partage de nourriture, l'utilisation des mêmes toilettes
    ou les piqûres de moustiques.
    La prévention comprend l'utilisation de préservatifs, les dépistages réguliers,
    la fidélité à un partenaire, l'évitement des aiguilles partagées et le suivi
    des conseils médicaux pendant la grossesse.)

▶ HARKNESS 3 — Fighting stigma
  Q : Why should people living with HIV/AIDS be treated with respect?
      (Pourquoi les personnes vivant avec le VIH/SIDA doivent-elles être
      traitées avec respect ?)
  Guide :
    · Pense à ce qu'est la stigmatisation (stigma) et ses conséquences.
    · Pense aux droits humains des personnes séropositives.
    · Pense à l'impact de l'exclusion sociale sur la lutte contre le SIDA.
  Corrigé :
    People living with HIV/AIDS must be treated with respect and dignity
    because HIV is not a punishment and anyone can be infected.
    Stigma — which means discrimination and rejection — pushes people away
    from getting tested and seeking treatment. This makes the epidemic worse.
    In Côte d'Ivoire and across Africa, fighting stigma is as important as
    fighting the virus itself. A person living with HIV who receives support
    from family and community can live a long, productive life with proper treatment.
    Tolerance and solidarity are essential weapons against HIV/AIDS.
    (Les personnes vivant avec le VIH/SIDA doivent être traitées avec respect
    et dignité car le VIH n'est pas une punition et n'importe qui peut être infecté.
    La stigmatisation — qui signifie discrimination et rejet — éloigne les gens
    du dépistage et du traitement. Cela aggrave l'épidémie.
    En Côte d'Ivoire et dans toute l'Afrique, lutter contre la stigmatisation
    est aussi important que lutter contre le virus lui-même. Une personne vivant
    avec le VIH qui reçoit le soutien de sa famille et de sa communauté peut
    vivre une vie longue et productive avec un traitement approprié.
    La tolérance et la solidarité sont des armes essentielles contre le VIH/SIDA.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre : HIV/AIDS — Le VIH/SIDA
Compétence : C7 · Hygiène et Santé / Health

Objectifs de la leçon :
  → Comprendre et expliquer ce qu'est le VIH/SIDA en anglais.
  → Identifier les modes de transmission et les moyens de prévention.
  → Utiliser le Passive Voice pour décrire comment le VIH est transmis.
  → Utiliser les modaux MUST / SHOULD pour formuler des conseils de prévention.
  → Enrichir son vocabulaire sur la santé, les maladies et la prévention.

Prérequis nécessaires :
  → Vocabulaire de base : disease (maladie) · sick (malade) · health (santé)
  → Les modaux MUST / SHOULD vus en Leçon 20
  → Le Present Simple pour les faits généraux
  → Notions de base sur les maladies vues en Leçon 19 (Endemic diseases)

┌─ TABLEAU DES STRUCTURES FONDAMENTALES ──────────────────┐
  Expression / Notion         | Valeur, emploi et exemple traduit
  ────────────────────────────|──────────────────────────────────
  HIV is transmitted by...    | Voie de transmission (Passive Voice)
                              | "HIV is transmitted by unprotected sex."
                              | (Le VIH est transmis par les rapports
                              |  sexuels non protégés.)
  ────────────────────────────|──────────────────────────────────
  HIV is NOT transmitted by.. | Idée reçue à corriger
                              | "HIV is not transmitted by a handshake."
                              | (Le VIH ne se transmet pas par une
                              |  poignée de main.)
  ────────────────────────────|──────────────────────────────────
  must + base verbale         | Obligation / prévention obligatoire
                              | "You must use a condom."
                              | (Tu dois utiliser un préservatif.)
  ────────────────────────────|──────────────────────────────────
  should + base verbale       | Conseil de santé
                              | "You should get tested regularly."
                              | (Tu devrais te faire dépister régulièrement.)
  ────────────────────────────|──────────────────────────────────
  People living with HIV...   | Formulation respectueuse
                              | "People living with HIV deserve respect."
                              | (Les personnes vivant avec le VIH méritent
                              |  le respect.)
  ────────────────────────────|──────────────────────────────────
  It is important to...       | Souligner l'importance d'une action
                              | "It is important to get tested."
                              | (Il est important de se faire dépister.)
  ────────────────────────────|──────────────────────────────────
  to be afraid of             | Avoir peur de
                              | "We should not be afraid of HIV+ people."
                              | (Nous ne devrions pas avoir peur des
                              |  personnes séropositives.)
└──────────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

TEXTE DE BASE — À lire et comprendre

  "HIV stands for Human Immunodeficiency Virus. It is a virus that destroys
  the immune system. AIDS is the most advanced stage of HIV infection.
  HIV is transmitted through unprotected sexual intercourse, contact with
  infected blood, and from mother to child during pregnancy or breastfeeding.
  However, HIV is not transmitted through handshakes, sharing food, hugging
  or mosquito bites.
  There is no cure for HIV yet, but antiretroviral drugs can help people
  living with HIV to lead a normal life.
  Prevention is the best weapon against HIV/AIDS. You must use condoms,
  get tested regularly and avoid sharing needles.
  People living with HIV deserve respect and solidarity, not discrimination."

  TRADUCTION FRANÇAISE :
  "HIV signifie Virus de l'Immunodéficience Humaine. C'est un virus qui détruit
  le système immunitaire. Le SIDA est le stade le plus avancé de l'infection par le VIH.
  Le VIH se transmet par les rapports sexuels non protégés, le contact avec du sang
  infecté, et de la mère à l'enfant pendant la grossesse ou l'allaitement.
  Cependant, le VIH ne se transmet pas par les poignées de main, le partage de
  nourriture, les câlins ou les piqûres de moustiques.
  Il n'existe pas encore de remède contre le VIH, mais les médicaments antirétroviraux
  peuvent aider les personnes vivant avec le VIH à mener une vie normale.
  La prévention est la meilleure arme contre le VIH/SIDA. Tu dois utiliser des
  préservatifs, te faire dépister régulièrement et éviter de partager des aiguilles.
  Les personnes vivant avec le VIH méritent le respect et la solidarité,
  pas la discrimination."

────────────────────────────────────────────────────────

┌─ À RETENIR — HIV vs AIDS ───────────────────────────────┐
  HIV (VIH) :
  → C'est le VIRUS. On peut être porteur du VIH sans être malade.
  → HIV = Human Immunodeficiency Virus
  → VIH = Virus de l'Immunodéficience Humaine
  → Prononciation : [eɪtʃ aɪ viː]

  AIDS (SIDA) :
  → C'est la MALADIE au stade avancé, causée par le VIH non traité.
  → AIDS = Acquired Immunodeficiency Syndrome
  → SIDA = Syndrome d'Immunodéficience Acquise
  → Prononciation : [eɪdz]

  RÈGLE IMPORTANTE :
  On dit "a person LIVING WITH HIV" (une personne vivant avec le VIH)
  et NON "an HIV person" — la formulation respectueuse est obligatoire.
└──────────────────────────────────────────────────────────┘

⚠ ATTENTION — Faux amis et pièges
  · "preservative" ≠ "préservatif"
    → preservative = conservateur alimentaire (dans la nourriture)
    → préservatif = condom [ˈkɒndəm]
  · "immune" se prononce [ɪˈmjuːn] — le "e" final ne se prononce pas.
  · "acquired" se prononce [əˈkwaɪərd] — le "qu" se dit "kw".
  · Stigma [ˈstɪɡmə] = stigmatisation — mot important dans ce sujet.

────────────────────────────────────────────────────────

VOCABULAIRE ESSENTIEL — HIV/AIDS

  ANGLAIS                  | PRONONCIATION           | TRADUCTION FR
  ─────────────────────────|─────────────────────────|──────────────────────────
  HIV                      | [eɪtʃ aɪ viː]           | VIH
  AIDS                     | [eɪdz]                  | SIDA
  virus                    | [ˈvaɪrəs]               | virus
  immune system            | [ɪˈmjuːn ˈsɪstəm]      | système immunitaire
  infection                | [ɪnˈfekʃən]             | infection
  transmission             | [trænsˈmɪʃən]           | transmission
  prevention               | [prɪˈvenʃən]            | prévention
  condom                   | [ˈkɒndəm]               | préservatif
  blood                    | [blʌd]                  | sang
  needle                   | [ˈniːdl]                | aiguille
  pregnancy                | [ˈpreɡnənsi]            | grossesse
  breastfeeding            | [ˈbrestfiːdɪŋ]          | allaitement
  antiretroviral           | [ˌæntɪˌretrəˈvaɪərəl]  | antirétroviral
  treatment                | [ˈtriːtmənt]            | traitement
  cure                     | [kjʊər]                 | remède / guérison
  stigma                   | [ˈstɪɡmə]              | stigmatisation
  discrimination           | [dɪˌskrɪmɪˈneɪʃən]    | discrimination
  solidarity               | [ˌsɒlɪˈdærɪti]         | solidarité
  to get tested            | [ɡet ˈtestɪd]           | se faire dépister
  awareness                | [əˈweənəs]              | sensibilisation / prise de conscience
  to transmit              | [trænsˈmɪt]             | transmettre
  unprotected              | [ˌʌnprəˈtektɪd]        | non protégé(e)
  to avoid                 | [əˈvɔɪd]               | éviter
  to deserve               | [dɪˈzɜːv]              | mériter

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Modes de transmission du VIH / How HIV is transmitted]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Deux zones distinctes côte à côte, séparées par une ligne verticale centrale :

  ZONE GAUCHE (fond rouge clair) — titre : HOW HIV IS TRANSMITTED
    → Élément 1 : icône goutte de sang — "Through infected blood"
                  (Par le sang infecté)
    → Élément 2 : icône couple — "Through unprotected sexual intercourse"
                  (Par les rapports sexuels non protégés)
    → Élément 3 : icône mère et enfant — "From mother to child during
                  pregnancy, childbirth or breastfeeding"
                  (De la mère à l'enfant pendant la grossesse,
                  l'accouchement ou l'allaitement)

  ZONE DROITE (fond vert clair) — titre : HOW HIV IS NOT TRANSMITTED
    → Élément 1 : icône poignée de main barrée — "NOT by handshakes"
                  (Pas par les poignées de main)
    → Élément 2 : icône assiette barrée — "NOT by sharing food"
                  (Pas par le partage de nourriture)
    → Élément 3 : icône moustique barré — "NOT by mosquito bites"
                  (Pas par les piqûres de moustiques)
    → Élément 4 : icône câlin barré — "NOT by hugging"
                  (Pas par les câlins)

  · Titre général en haut : HIV — TRANSMISSION
  · Chaque élément porte sa traduction française en italique en dessous
  · Symbole ✔ dans la zone rouge · Symbole ✘ dans la zone verte
Note générateur : Canva (Two-column infographic) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Carte mentale : HIV/AIDS · Prevention & Solidarity]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Centre (cercle rouge) : HIV/AIDS
  · 4 branches principales :
      Branche 1 (rouge foncé) : WHAT IT IS
        → virus / immune system / AIDS = advanced stage
      Branche 2 (orange) : HOW IT SPREADS
        → infected blood / unprotected sex / mother to child
      Branche 3 (verte) : HOW TO PREVENT IT
        → use condoms / get tested / avoid shared needles / be faithful
      Branche 4 (bleue) : HOW TO FIGHT STIGMA
        → respect / solidarity / no discrimination / support
  · Traductions françaises en italique sous chaque sous-élément
  · Légende : rouge = danger · vert = protection · bleu = solidarité
Note générateur : Canva (Mind Map) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — HIV (VIH)
  EN : HIV stands for Human Immunodeficiency Virus. It is a virus that
  attacks and weakens the human immune system, making the body unable
  to fight infections and diseases.
  FR : VIH signifie Virus de l'Immunodéficience Humaine. C'est un virus
  qui attaque et affaiblit le système immunitaire humain, rendant le corps
  incapable de combattre les infections et les maladies.

◆ DÉFINITION DPFC — AIDS (SIDA)
  EN : AIDS stands for Acquired Immunodeficiency Syndrome. It is the most
  advanced and serious stage of HIV infection, when the immune system is
  severely damaged and the body can no longer defend itself against infections.
  FR : SIDA signifie Syndrome d'Immunodéficience Acquise. C'est le stade
  le plus avancé et le plus grave de l'infection par le VIH, quand le système
  immunitaire est gravement endommagé et que le corps ne peut plus se défendre
  contre les infections.

◆ DÉFINITION DPFC — Transmission (Transmission)
  EN : Transmission refers to the way a virus or disease passes from one
  person to another. HIV is transmitted through blood, unprotected sex
  and from mother to child.
  FR : La transmission désigne la façon dont un virus ou une maladie passe
  d'une personne à une autre. Le VIH se transmet par le sang, les rapports
  sexuels non protégés et de la mère à l'enfant.

◆ DÉFINITION DPFC — Prevention (Prévention)
  EN : Prevention is the set of actions taken to stop a disease from
  spreading or to protect oneself from infection before it occurs.
  FR : La prévention est l'ensemble des actions prises pour empêcher
  une maladie de se propager ou pour se protéger d'une infection avant qu'elle survienne.

◆ DÉFINITION DPFC — Stigma (Stigmatisation)
  EN : Stigma refers to negative attitudes, discrimination and rejection
  directed at people living with HIV/AIDS. Stigma is harmful because it
  prevents people from getting tested and treated.
  FR : La stigmatisation désigne les attitudes négatives, la discrimination
  et le rejet dirigés contre les personnes vivant avec le VIH/SIDA.
  La stigmatisation est nuisible car elle empêche les gens de se faire
  dépister et traiter.

◆ DÉFINITION DPFC — Antiretroviral treatment / ARV (Traitement antirétroviral)
  EN : Antiretroviral treatment is a combination of medicines that reduce
  the amount of HIV in the body. It does not cure HIV but allows people
  living with HIV to live a long and healthy life.
  FR : Le traitement antirétroviral est une combinaison de médicaments qui
  réduit la quantité de VIH dans le corps. Il ne guérit pas le VIH mais
  permet aux personnes vivant avec le VIH de vivre longtemps et en bonne santé.

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

  ANGLAIS LV1 · Leçon 21 · 3ème
  HIV/AIDS — Le VIH/SIDA
  Compétence C7 · Hygiène et Santé · Prof. Evelyne ASSI

  I. VOCABULAIRE ESSENTIEL
  ─────────────────────────
  HIV [eɪtʃ aɪ viː] = VIH (Virus de l'Immunodéficience Humaine)
  AIDS [eɪdz] = SIDA (Syndrome d'Immunodéficience Acquise)
  virus = virus
  immune system = système immunitaire
  transmission = transmission
  blood [blʌd] = sang
  condom [ˈkɒndəm] = préservatif
  needle = aiguille
  pregnancy = grossesse
  treatment = traitement
  cure = remède / guérison
  stigma = stigmatisation
  discrimination = discrimination
  solidarity = solidarité
  to get tested = se faire dépister
  to avoid = éviter
  awareness = sensibilisation

  II. CE QU'EST LE VIH/SIDA
  ──────────────────────────
  · HIV is a virus that destroys the immune system.
    (Le VIH est un virus qui détruit le système immunitaire.)
  · AIDS is the most advanced stage of HIV infection.
    (Le SIDA est le stade le plus avancé de l'infection par le VIH.)
  · There is no cure for HIV yet, but ARV treatment helps people
    living with HIV to lead a normal life.
    (Il n'existe pas encore de remède contre le VIH, mais le traitement
    ARV aide les personnes vivant avec le VIH à mener une vie normale.)

  III. MODES DE TRANSMISSION
  ───────────────────────────
  HIV IS transmitted through (Le VIH SE transmet par) :
    · Unprotected sexual intercourse (les rapports sexuels non protégés)
    · Contact with infected blood (le contact avec du sang infecté)
    · From mother to child during pregnancy, childbirth or breastfeeding
      (de la mère à l'enfant pendant la grossesse, l'accouchement ou l'allaitement)

  HIV is NOT transmitted through (Le VIH ne se transmet PAS par) :
    · Handshakes (les poignées de main)
    · Sharing food (le partage de nourriture)
    · Hugging (les câlins)
    · Mosquito bites (les piqûres de moustiques)
    · Using the same toilet (l'utilisation des mêmes toilettes)

  IV. PRÉVENTION
  ───────────────
  · You must use a condom. (Tu dois utiliser un préservatif.)
  · You must avoid sharing needles. (Tu dois éviter de partager des aiguilles.)
  · You should get tested regularly. (Tu devrais te faire dépister régulièrement.)
  · You should be faithful to one partner.
    (Tu devrais être fidèle à un seul partenaire.)
  · Pregnant women should follow medical advice to protect their baby.
    (Les femmes enceintes devraient suivre les conseils médicaux pour protéger leur bébé.)

  V. SOLIDARITÉ
  ──────────────
  · People living with HIV deserve respect and solidarity.
    (Les personnes vivant avec le VIH méritent le respect et la solidarité.)
  · We must not discriminate against people living with HIV.
    (Nous ne devons pas discriminer les personnes vivant avec le VIH.)
  · Stigma makes the epidemic worse — solidarity makes it better.
    (La stigmatisation aggrave l'épidémie — la solidarité l'améliore.)

~ ANALOGIE CI
  En Côte d'Ivoire, chaque 1er décembre, c'est la Journée Mondiale de Lutte
  contre le SIDA. À Abidjan, des associations comme AIBEF et ACAD organisent
  des marches et des séances de dépistage gratuit dans les quartiers.
  Des milliers de jeunes ivoiriens portent le ruban rouge pour montrer
  leur solidarité avec les personnes vivant avec le VIH.
  "Wearing the red ribbon is not enough — knowing the facts is the first step."
  (Porter le ruban rouge ne suffit pas — connaître les faits est la première étape.)

~ ANALOGIE CI 2
  Imagine un feu de brousse en saison sèche dans le nord de la Côte d'Ivoire.
  Si tu ne l'éteins pas à temps, il se propage partout.
  Le VIH, c'est pareil : sans prévention, il se propage.
  Mais avec les bons outils — préservatif, dépistage, traitement ARV —
  on peut le contrôler, comme on contrôle un feu avec de l'eau.
  "Prevention is the water that stops the fire."
  (La prévention est l'eau qui arrête le feu.)

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. HIV est le virus ; AIDS est la maladie au stade avancé — ce sont deux choses différentes.
  2. Le VIH se transmet par le sang, les rapports sexuels non protégés et de mère à enfant.
  3. Le VIH ne se transmet PAS par les poignées de main, les câlins ou les piqûres de moustiques.
  4. Il n'existe pas de remède, mais le traitement ARV permet de vivre normalement.
  5. La stigmatisation aggrave l'épidémie — le respect et la solidarité sont essentiels.

5 mots clés + définition courte :
  · HIV → virus qui détruit le système immunitaire
  · AIDS → stade avancé de l'infection par le VIH
  · Transmission → façon dont le virus passe d'une personne à une autre
  · Prevention → ensemble des actions pour éviter la contamination
  · Stigma → rejet et discrimination des personnes séropositives

3 erreurs fréquentes + correction :
  Erreur 1 : Confondre HIV et AIDS.
  Correction : HIV est le virus / AIDS est la maladie à un stade avancé.
               On peut vivre avec HIV sans avoir AIDS si on est traité.

  Erreur 2 : Croire que le VIH se transmet par les moustiques ou le partage de nourriture.
  Correction : "HIV is NOT transmitted by mosquito bites or sharing food."
               Ce sont des idées reçues — la science le prouve.

  Erreur 3 : Dire "an AIDS person" pour parler d'une personne séropositive.
  Correction : On dit "a person living with HIV" — formulation respectueuse et correcte.

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Vrai ou Faux ? · Notion ciblée : Compréhension des modes de transmission

  Énoncé :
  Lis chaque affirmation. Écris VRAI (True) ou FAUX (False).
  Si c'est FAUX, écris la phrase correcte en anglais et traduis-la en français.

  a) HIV is transmitted by sharing food with an infected person.
  b) AIDS is the most advanced stage of HIV infection.
  c) HIV is transmitted through unprotected sexual intercourse.
  d) You can get HIV from a mosquito bite.
  e) Antiretroviral drugs can cure HIV completely.

  MÉTHODE
  Étape 1 · Lis chaque affirmation et rappelle-toi les modes de transmission vus en classe.
  Étape 2 · Décide : est-ce VRAI ou FAUX selon les faits scientifiques ?
  Étape 3 · Si c'est FAUX, réécris la phrase correcte en anglais.
  Étape 4 · Traduis la phrase correcte en français.
  ✔ Vérification : Relis la trace écrite pour confirmer tes réponses.
  ✔ Conclusion :
    a) FAUX — HIV is NOT transmitted by sharing food with an infected person.
              (Le VIH ne se transmet PAS par le partage de nourriture.)
    b) VRAI — AIDS is the most advanced stage of HIV infection.
              (Le SIDA est le stade le plus avancé de l'infection par le VIH.)
    c) VRAI — HIV is transmitted through unprotected sexual intercourse.
              (Le VIH se transmet par les rapports sexuels non protégés.)
    d) FAUX — You cannot get HIV from a mosquito bite.
              (On ne peut pas attraper le VIH par une piqûre de moustique.)
    e) FAUX — Antiretroviral drugs cannot cure HIV, but they help people
              living with HIV to lead a normal life.
              (Les médicaments antirétroviraux ne guérissent pas le VIH,
              mais ils aident les personnes vivant avec le VIH à mener
              une vie normale.)
  ✔ Ce que cet exercice mobilise :
    · Compréhension des modes de transmission du VIH
    · Correction d'idées reçues en anglais
    · Structure "HIV is / is not transmitted by..."
    · Traduction anglais → français
    · Esprit critique et éducation à la santé

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Conseils de prévention · Notion ciblée : Production écrite + Modaux + Prévention

  Énoncé :
  Koffi, 15 ans, lycéen à Abidjan, doit préparer un exposé de 6 phrases
  sur la prévention du VIH/SIDA pour sa classe.
  Aide-le à écrire cet exposé en anglais, en utilisant MUST, MUST NOT et SHOULD.
  Traduis chaque phrase en français.

  MÉTHODE
  Étape 1 · Pense aux différents moyens de prévention : préservatif, dépistage,
            aiguilles, fidélité, suivi médical pendant la grossesse.
  Étape 2 · Pour chaque conseil, choisis le modal approprié :
            → Règle stricte de santé publique : MUST
            → Interdiction : MUST NOT
            → Conseil recommandé : SHOULD
  Étape 3 · Construis chaque phrase : Sujet + Modal + Base verbale + Complément.
  Étape 4 · Traduis chaque phrase en français.
  ✔ Vérification : Vérifie que tu as utilisé MUST, MUST NOT et SHOULD au moins une fois chacun.
  ✔ Conclusion — Exposé de Koffi :
    1. To prevent HIV, you must use a condom during sexual intercourse.
       (Pour prévenir le VIH, tu dois utiliser un préservatif lors des rapports sexuels.)
    2. You must not share needles or syringes with other people.
       (Tu ne dois pas partager des aiguilles ou des seringues avec d'autres personnes.)
    3. You should get tested for HIV regularly, especially if you are sexually active.
       (Tu devrais te faire dépister régulièrement pour le VIH, surtout si tu es sexuellement actif.)
    4. You should be faithful to one partner to reduce the risk of infection.
       (Tu devrais être fidèle à un seul partenaire pour réduire le risque d'infection.)
    5. Pregnant women must follow medical advice to prevent transmission to their baby.
       (Les femmes enceintes doivent suivre les conseils médicaux pour prévenir
       la transmission à leur bébé.)
    6. We must not discriminate against people living with HIV — they deserve our respect.
       (Nous ne devons pas discriminer les personnes vivant avec le VIH —
       elles méritent notre respect.)
  ✔ Ce que cet exercice mobilise :
    · Production écrite guidée sur un sujet DPFC (VIH/SIDA)
    · Modaux MUST / MUST NOT / SHOULD dans un contexte de santé publique
    · Vocabulaire de la prévention et de la solidarité
    · Traduction anglais → français
    · Éducation civique et à la santé (contenu DPFC intégré)

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulaire du VIH/SIDA · Notions travaillées : Lexique, traduction, sigles

  Partie A — Que signifient ces sigles ? Écris leur développé complet en anglais
  et leur traduction en français.
    1. HIV = _____________ → Traduction FR : _____________
    2. AIDS = _____________ → Traduction FR : _____________
    3. ARV = _____________ → Traduction FR : _____________

  Partie B — Relie chaque mot anglais à sa traduction française.
    1. blood               A. aiguille
    2. needle              B. sang
    3. stigma              C. dépistage
    4. testing             D. solidarité
    5. solidarity          E. stigmatisation

  GUIDE
  Étape 1 · Pour la Partie A, rappelle-toi les définitions complètes vues en classe.
  Étape 2 · Pour la Partie B, relis le vocabulaire essentiel de la trace écrite.
  Étape 3 · Vérifie l'orthographe de chaque mot.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Transmission : Vrai ou Faux · Notions travaillées : Modes de transmission, idées reçues

  Lis chaque affirmation. Écris TRUE (Vrai) ou FALSE (Faux).
  Pour chaque réponse FALSE, écris la phrase correcte en anglais.

  a) HIV is transmitted by using the same toilet as an infected person.
  b) HIV can be transmitted from an infected mother to her baby during breastfeeding.
  c) You can get HIV by hugging a person living with HIV.
  d) HIV is transmitted through contact with infected blood.
  e) Sharing food with an HIV positive person can transmit the virus.

  GUIDE
  Étape 1 · Relis la section "HIV is / is not transmitted" de ta trace écrite.
  Étape 2 · Pour chaque affirmation, décide si elle est vraie ou fausse selon les faits.
  Étape 3 · Pour les fausses, réécris la phrase correctement.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Compléter un texte lacunaire · Notions travaillées : Vocabulaire, structures de la leçon

  Complète le texte avec les mots suivants. Traduis le texte complet en français.
  (Mots à utiliser : immune system · transmitted · condom · tested · discrimination
                     cure · solidarity · AIDS · virus · prevention)

  "HIV is a _______ that attacks the _______. When the body can no longer
  defend itself, HIV becomes _______. HIV is _______ through unprotected
  sexual intercourse, infected blood and from mother to child.
  There is no _______ for HIV yet, but antiretroviral drugs help patients
  to live normally. The best weapon is _______: you must use a _______
  and you should get _______ regularly.
  People living with HIV do not deserve _______. They need _______ and respect."

  GUIDE
  Étape 1 · Lis le texte entier une première fois pour comprendre le sens général.
  Étape 2 · Pour chaque blanc, choisis le mot qui convient grammaticalement et logiquement.
  Étape 3 · Relis le texte complet pour vérifier la cohérence.
  Étape 4 · Traduis le texte complet en français.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Questions de compréhension écrite · Notions travaillées : Compréhension, structures interrogatives

  Lis ce texte, puis réponds aux questions en anglais avec des phrases complètes.
  Traduis chaque réponse en français.

  TEXTE :
  "Fatou is a 16-year-old girl from Abidjan. Last week, her school organised
  an HIV/AIDS awareness day. A doctor came to explain what HIV is and how it
  spreads. Fatou learned that HIV is transmitted through unprotected sex,
  infected blood and from mother to child. She also learned that HIV is NOT
  transmitted through handshakes or sharing food. The doctor encouraged all
  students to get tested and to always use condoms. He also said that people
  living with HIV deserve respect and should never be discriminated against.
  Fatou left school that day feeling informed and ready to share what she had
  learned with her family."

  Questions :
  Q1 : How old is Fatou?
  Q2 : What did the school organise last week?
  Q3 : Name two ways HIV is transmitted according to the text.
  Q4 : Why should people living with HIV never be discriminated against?
  Q5 : What did Fatou decide to do after the awareness day?

  GUIDE
  Étape 1 · Lis le texte une première fois pour en comprendre le sens général.
  Étape 2 · Lis chaque question et repère dans le texte la partie qui y répond.
  Étape 3 · Formule ta réponse avec une phrase complète en anglais.
  Étape 4 · Traduis ta réponse en français.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Production écrite : Lutte contre la stigmatisation · Notions travaillées : Expression écrite, solidarité, modaux

  Dans ton lycée, un élève a appris qu'un de ses camarades vit avec le VIH.
  Il commence à se moquer de lui et à l'exclure.
  Écris un texte de 10 à 12 phrases pour convaincre cet élève de changer
  son comportement. Tu dois :
    · Expliquer ce qu'est le VIH (2 phrases)
    · Expliquer comment le VIH ne se transmet pas (2 phrases)
    · Dire pourquoi la stigmatisation est dangereuse (2 phrases)
    · Donner 2 obligations avec MUST / MUST NOT
    · Donner 2 conseils avec SHOULD
    · Terminer par une phrase de solidarité

  GUIDE
  Étape 1 · Commence par adresser ton camarade : "Dear classmate, I need to talk to you about..."
  Étape 2 · Explique brièvement ce qu'est le VIH.
  Étape 3 · Corrige les idées reçues sur la transmission.
  Étape 4 · Explique pourquoi la stigmatisation aggrave la situation.
  Étape 5 · Utilise MUST / MUST NOT pour les obligations.
  Étape 6 · Utilise SHOULD pour les conseils.
  Étape 7 · Termine par un message de solidarité.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — An awareness speech about HIV/AIDS
  Durée conseillée : 35 minutes

  Énoncé :
  Le 1er décembre, Journée Mondiale de Lutte contre le SIDA, ton lycée
  t'a choisi pour prononcer un discours devant tous les élèves.
  Écris ce discours en anglais (15 à 18 phrases). Tu dois obligatoirement :
    · Expliquer ce qu'est le VIH et ce qu'est le SIDA (2-3 phrases)
    · Décrire les 3 modes de transmission du VIH (3 phrases)
    · Corriger 3 idées reçues sur la transmission (3 phrases)
    · Donner 3 conseils de prévention avec MUST / MUST NOT / SHOULD
    · Appeler à la solidarité envers les personnes vivant avec le VIH (2 phrases)
    · Terminer par une phrase d'espoir

  Contraintes obligatoires :
    ✔ Utiliser MUST au moins 2 fois
    ✔ Utiliser MUST NOT au moins 2 fois
    ✔ Utiliser SHOULD au moins 2 fois
    ✔ Utiliser au moins 10 mots de vocabulaire vus en classe
    ✔ Traduire le discours complet en français
    ✔ Format discours : formule d'adresse / corps / conclusion

  GUIDE
  Étape 1 · Commence par la formule d'adresse :
            "Dear students, teachers and guests, today is World AIDS Day..."
  Étape 2 · Définis le VIH et le SIDA.
  Étape 3 · Explique les 3 modes de transmission avec "HIV is transmitted through..."
  Étape 4 · Corrige les idées reçues avec "HIV is NOT transmitted by..."
  Étape 5 · Donne tes conseils de prévention avec MUST, MUST NOT et SHOULD.
  Étape 6 · Lance un appel à la solidarité.
  Étape 7 · Termine par une phrase d'espoir et traduis tout le discours en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON

  "Dear students, teachers and guests,
  Today is the 1st of December — World AIDS Day — and I am honoured to speak
  to you about a subject that concerns us all.
  (Chers élèves, enseignants et invités, aujourd'hui, c'est le 1er décembre
  — la Journée Mondiale de Lutte contre le SIDA — et je suis honoré de vous
  parler d'un sujet qui nous concerne tous.)

  First, let me explain what HIV and AIDS are.
  HIV stands for Human Immunodeficiency Virus. It is a virus that attacks
  and destroys our immune system, which is our body's defence against diseases.
  AIDS is the most advanced stage of HIV infection, when the body can no longer
  fight even simple infections.
  (Premièrement, laissez-moi vous expliquer ce que sont le VIH et le SIDA.
  HIV signifie Virus de l'Immunodéficience Humaine. C'est un virus qui attaque
  et détruit notre système immunitaire, qui est la défense de notre corps contre
  les maladies. Le SIDA est le stade le plus avancé de l'infection par le VIH,
  quand le corps ne peut plus combattre même les infections simples.)

  HIV is transmitted in three main ways.
  First, HIV is transmitted through unprotected sexual intercourse.
  Second, HIV is transmitted through contact with infected blood,
  for example through shared needles.
  Third, HIV is transmitted from an infected mother to her child during
  pregnancy, childbirth or breastfeeding.
  (Le VIH se transmet de trois façons principales.
  Premièrement, le VIH se transmet par les rapports sexuels non protégés.
  Deuxièmement, le VIH se transmet par le contact avec du sang infecté,
  par exemple par des aiguilles partagées.
  Troisièmement, le VIH se transmet d'une mère infectée à son enfant
  pendant la grossesse, l'accouchement ou l'allaitement.)

  Now, let me correct some false ideas.
  HIV is NOT transmitted by handshakes — you cannot get HIV by touching
  a person living with HIV.
  HIV is NOT transmitted by sharing food or drinking from the same glass.
  HIV is NOT transmitted by mosquito bites or by using the same toilet.
  (Maintenant, laissez-moi corriger quelques fausses idées.
  Le VIH ne se transmet PAS par les poignées de main — on ne peut pas
  attraper le VIH en touchant une personne vivant avec le VIH.
  Le VIH ne se transmet PAS par le partage de nourriture ou en buvant
  dans le même verre.
  Le VIH ne se transmet PAS par les piqûres de moustiques ou en utilisant
  les mêmes toilettes.)

  Here is how we can protect ourselves.
  You must use a condom every time you have sexual intercourse.
  You must not share needles or syringes with anyone.
  You should get tested for HIV regularly — knowing your status is the
  first step to protecting yourself and others.
  (Voici comment nous pouvons nous protéger.
  Tu dois utiliser un préservatif à chaque rapport sexuel.
  Tu ne dois pas partager des aiguilles ou des seringues avec qui que ce soit.
  Tu devrais te faire dépister régulièrement pour le VIH — connaître ton statut
  est la première étape pour te protéger toi et les autres.)

  Finally, I want to talk about solidarity.
  People living with HIV deserve our respect and our support.
  We must not discriminate against them — stigma only makes things worse
  and pushes people away from getting help.
  (Enfin, je veux parler de solidarité.
  Les personnes vivant avec le VIH méritent notre respect et notre soutien.
  Nous ne devons pas les discriminer — la stigmatisation ne fait qu'aggraver
  les choses et éloigne les gens de l'aide dont ils ont besoin.)

  Together, with knowledge, prevention and solidarity,
  we can win the fight against HIV/AIDS.
  Thank you.
  (Ensemble, avec la connaissance, la prévention et la solidarité,
  nous pouvons gagner la lutte contre le VIH/SIDA.
  Merci.)"

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1

  Partie A :
    1. HIV = Human Immunodeficiency Virus → VIH = Virus de l'Immunodéficience Humaine
    2. AIDS = Acquired Immunodeficiency Syndrome → SIDA = Syndrome d'Immunodéficience Acquise
    3. ARV = Antiretroviral → ARV = Antirétroviral

  Partie B :
    1 — B : blood = sang
    2 — A : needle = aiguille
    3 — E : stigma = stigmatisation
    4 — C : testing = dépistage
    5 — D : solidarity = solidarité

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2

  a) FALSE — HIV is NOT transmitted by using the same toilet as an infected person.
             (Le VIH ne se transmet PAS en utilisant les mêmes toilettes qu'une personne infectée.)

  b) TRUE — HIV can be transmitted from an infected mother to her baby during breastfeeding.
            (Le VIH peut se transmettre d'une mère infectée à son bébé pendant l'allaitement.)

  c) FALSE — HIV is NOT transmitted by hugging a person living with HIV.
             (Le VIH ne se transmet PAS en étreignant une personne vivant avec le VIH.)

  d) TRUE — HIV is transmitted through contact with infected blood.
            (Le VIH se transmet par le contact avec du sang infecté.)

  e) FALSE — Sharing food with an HIV positive person cannot transmit the virus.
             (Partager de la nourriture avec une personne séropositive ne peut pas transmettre le virus.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3

  Texte complété :
  "HIV is a VIRUS that attacks the IMMUNE SYSTEM. When the body can no longer
  defend itself, HIV becomes AIDS. HIV is TRANSMITTED through unprotected
  sexual intercourse, infected blood and from mother to child.
  There is no CURE for HIV yet, but antiretroviral drugs help patients
  to live normally. The best weapon is PREVENTION: you must use a CONDOM
  and you should get TESTED regularly.
  People living with HIV do not deserve DISCRIMINATION. They need SOLIDARITY and respect."

  Traduction française complète :
  "Le VIH est un VIRUS qui attaque le SYSTÈME IMMUNITAIRE. Quand le corps
  ne peut plus se défendre, le VIH devient le SIDA. Le VIH est TRANSMIS par
  les rapports sexuels non protégés, le sang infecté et de la mère à l'enfant.
  Il n'existe pas encore de REMÈDE contre le VIH, mais les médicaments
  antirétroviraux aident les patients à vivre normalement. La meilleure arme
  est la PRÉVENTION : tu dois utiliser un PRÉSERVATIF et tu devrais te faire
  DÉPISTER régulièrement.
  Les personnes vivant avec le VIH ne méritent pas la DISCRIMINATION.
  Elles ont besoin de SOLIDARITÉ et de respect."

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4

  Q1 : How old is Fatou?
  R1 : Fatou is 16 years old.
       (Fatou a 16 ans.)

  Q2 : What did the school organise last week?
  R2 : The school organised an HIV/AIDS awareness day.
       (L'école a organisé une journée de sensibilisation sur le VIH/SIDA.)

  Q3 : Name two ways HIV is transmitted according to the text.
  R3 : According to the text, HIV is transmitted through unprotected sex
       and through infected blood.
       (Selon le texte, le VIH se transmet par les rapports sexuels non protégés
       et par le sang infecté.)

  Q4 : Why should people living with HIV never be discriminated against?
  R4 : People living with HIV should never be discriminated against because
       they deserve respect.
       (Les personnes vivant avec le VIH ne devraient jamais être discriminées
       car elles méritent le respect.)

  Q5 : What did Fatou decide to do after the awareness day?
  R5 : Fatou decided to share what she had learned with her family.
       (Fatou a décidé de partager ce qu'elle avait appris avec sa famille.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5

  Exemple de texte :

  "Dear classmate,
  I need to talk to you about something important.
  (Cher camarade, je dois te parler de quelque chose d'important.)

  HIV is a virus that attacks the immune system. It can be controlled with
  medication, and people living with HIV can lead a completely normal life.
  (Le VIH est un virus qui attaque le système immunitaire. Il peut être contrôlé
  avec des médicaments, et les personnes vivant avec le VIH peuvent mener
  une vie tout à fait normale.)

  HIV is not transmitted by handshakes or by being near someone who is HIV positive.
  You cannot get HIV by sharing food, using the same classroom or sitting
  next to our classmate.
  (Le VIH ne se transmet pas par les poignées de main ou en étant près
  d'une personne séropositive. On ne peut pas attraper le VIH en partageant
  de la nourriture, en utilisant la même salle de classe ou en s'asseyant
  à côté de notre camarade.)

  Stigma is dangerous because it pushes people living with HIV away from
  getting tested and treated. When we discriminate, we make the epidemic worse,
  not better.
  (La stigmatisation est dangereuse car elle éloigne les personnes vivant
  avec le VIH du dépistage et du traitement. Quand nous discriminons, nous
  aggravons l'épidémie, pas le contraire.)

  You must stop mocking our classmate — this behaviour is unacceptable.
  You must not exclude him from our group because of his health condition.
  (Tu dois arrêter de te moquer de notre camarade — ce comportement est inacceptable.
  Tu ne dois pas l'exclure de notre groupe à cause de son état de santé.)

  You should try to understand what he is going through and offer him your support.
  You should talk to him normally, just as you would with any other classmate.
  (Tu devrais essayer de comprendre ce qu'il traverse et lui offrir ton soutien.
  Tu devrais lui parler normalement, comme tu le ferais avec n'importe quel
  autre camarade.)

  Remember : solidarity is our greatest strength.
  Together, we can make our school a place of respect and dignity for everyone.
  (Souviens-toi : la solidarité est notre plus grande force.
  Ensemble, nous pouvons faire de notre école un endroit de respect et de
  dignité pour tous.)"

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "The greatest weapon against HIV is not a medicine — it is knowledge,
  compassion and solidarity."
  (La plus grande arme contre le VIH n'est pas un médicament —
  c'est la connaissance, la compassion et la solidarité.)

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
